from __future__ import annotations

import os
import shutil
import time
from collections import deque
from dataclasses import dataclass
from datetime import datetime

import serial
from openpyxl.utils.exceptions import InvalidFileException
from PySide6.QtCore import QThread, Signal
from serial.tools import list_ports

from excel_writer import (
    append_raw_csv_row,
    check_output_available,
    create_raw_csv,
    create_run_output_dir,
    create_stable_excel,
    raw_csv_to_excel,
    refresh_stable_excel,
    resolve_output_path,
)
from filter_logic import (
    WAVELENGTH_CODES,
    ZERO_VALUE_EPSILON,
    absorbance_from_transmittance,
    air_window_is_stable,
    calibrate_intensity,
    channel_gain_calibration,
    channel_indices,
    filter_is_inserted,
    filter_window_is_stable,
    window_diagnostics,
)
from serial_protocol import (
    SERIAL_BAUD,
    SERIAL_BYTESIZE,
    SERIAL_PARITY,
    SERIAL_STOPBITS,
    build_cmd,
    frame_to_hex,
    parse_absorbance_frame,
    parse_frames,
)


def list_serial_port_names() -> list[str]:
    return [port.device for port in list_ports.comports()]


class SaveOutputError(OSError):
    """Excel 输出保存失败。"""


@dataclass(frozen=True)
class TestConfig:
    port: str
    wavelength: int
    channel_group: int
    test_mode: str
    output: str
    stable_output: str
    filter_stable_window_frames: int
    filter_stable_cv_limit: float
    filter_stable_slope_limit: float
    filter_stable_range_limit: float
    air_tolerance: float
    air_warmup_seconds: float
    air_stable_window_frames: int
    air_stable_cv_limit: float
    air_stable_slope_limit: float
    air_stable_range_limit: float
    no_start: bool
    keep_light: bool


class SerialTestWorker(QThread):
    status_changed = Signal(str)
    log_added = Signal(str)
    values_received = Signal(object)
    baseline_progress_changed = Signal(int, int)
    baseline_value_changed = Signal(int, float)
    channel_state_changed = Signal(int, str)
    stable_value_changed = Signal(int, float, float, float, float)
    row_status_changed = Signal(int, int)
    output_paths_ready = Signal(str, str)
    test_finished = Signal()

    def __init__(self, config: TestConfig, parent=None) -> None:
        super().__init__(parent)
        self.config = config
        self._running = False

    def stop(self) -> None:
        self._running = False

    def run(self) -> None:
        self._running = True
        buffer = bytearray()
        target_indices = channel_indices(self.config.channel_group)
        filter_mode = self.config.test_mode == "filter"
        filter_window_count = self.config.filter_stable_window_frames
        filter_cv_limit = self.config.filter_stable_cv_limit / 100
        filter_slope_limit = self.config.filter_stable_slope_limit / 100
        filter_range_limit = self.config.filter_stable_range_limit / 100
        air_tolerance = self.config.air_tolerance / 100
        air_window_count = self.config.air_stable_window_frames
        air_cv_limit = self.config.air_stable_cv_limit / 100
        air_slope_limit = self.config.air_stable_slope_limit / 100
        air_range_limit = self.config.air_stable_range_limit / 100
        had_error = False
        final_status = "状态：未运行"

        if self.config.test_mode not in ("air", "filter"):
            self.status_changed.emit("状态：运行模式配置无效")
            self.log_added.emit("test_mode 必须是 air 或 filter")
            self.test_finished.emit()
            return

        if filter_mode and (
            filter_window_count < 2
            or filter_cv_limit < 0
            or filter_slope_limit < 0
            or filter_range_limit < 0
        ):
            self.status_changed.emit("状态：滤光片判稳配置无效")
            self.log_added.emit("滤光片判稳配置必须为非负数，滤光片窗口帧数必须大于等于 2")
            self.test_finished.emit()
            return

        if filter_mode and (
            self.config.air_warmup_seconds < 0
            or not 0 <= air_tolerance < 1
            or air_window_count < 2
            or air_cv_limit < 0
            or air_slope_limit < 0
            or air_range_limit < 0
        ):
            self.status_changed.emit("状态：空气判稳配置无效")
            self.log_added.emit(
                "空气判稳配置必须为非负数，air_tolerance 必须小于 100，空气窗口帧数必须大于等于 2"
            )
            self.test_finished.emit()
            return

        output_dir = create_run_output_dir()
        raw_output_path = resolve_output_path(output_dir, self.config.output)
        stable_output_path = resolve_output_path(output_dir, self.config.stable_output)
        self.output_paths_ready.emit(os.path.abspath(raw_output_path), os.path.abspath(stable_output_path))

        if filter_mode and os.path.abspath(raw_output_path) == os.path.abspath(stable_output_path):
            self.status_changed.emit("状态：原始数据文件和稳定值文件不能使用同一个路径")
            self.test_finished.emit()
            return

        if not check_output_available(raw_output_path):
            self.status_changed.emit("状态：原始数据文件被占用，请关闭后重试")
            self.test_finished.emit()
            return
        if filter_mode and not check_output_available(stable_output_path):
            self.status_changed.emit("状态：稳定值文件被占用，请关闭后重试")
            self.test_finished.emit()
            return

        raw_csv_file, raw_csv_writer, raw_csv_path = create_raw_csv(raw_output_path, target_indices)
        stable_wb = stable_ws = None
        stable_values = None
        stable_records = None
        baseline_values = None

        if filter_mode:
            stable_wb, stable_ws = create_stable_excel(stable_output_path, target_indices)
            stable_values = [[] for _ in range(24)]
            stable_records = [[] for _ in range(24)]

        row_count = 0
        stable_count = 0
        start_cmd = build_cmd(1, 0x50, (WAVELENGTH_CODES[self.config.wavelength] << 8) | 0x01)
        stop_cmd = build_cmd(1, 0x50, 0)

        ser = None
        try:
            ser = serial.Serial(
                self.config.port,
                SERIAL_BAUD,
                bytesize=SERIAL_BYTESIZE,
                parity=SERIAL_PARITY,
                stopbits=SERIAL_STOPBITS,
                timeout=0.1,
                write_timeout=1,
            )
            ser.reset_input_buffer()
            self.status_changed.emit(f"状态：已连接 {self.config.port}")
            self.log_added.emit(f"串口配置: {SERIAL_BAUD} 8N1")

            if not self.config.no_start:
                ser.write(start_cmd)
                self.log_added.emit(f"已发送启动命令: {frame_to_hex(start_cmd)}")

            if filter_mode:
                self.status_changed.emit("状态：请保持空气状态，正在建立空气基底")
                self.log_added.emit(
                    f"空气判稳配置: 最短暖机 {self.config.air_warmup_seconds:g} 秒，"
                    f"单个 {air_window_count} 帧窗口，"
                    f"CV ≤ {self.config.air_stable_cv_limit:g}%，"
                    f"绝对斜率 ≤ {self.config.air_stable_slope_limit:g} 个百分点/帧，"
                    f"极差 ≤ {self.config.air_stable_range_limit:g} 个百分点"
                )
                self.log_added.emit(
                    f"滤光片判稳配置: 单个 {filter_window_count} 帧窗口，"
                    f"CV ≤ {self.config.filter_stable_cv_limit:g}%，"
                    f"绝对斜率 ≤ {self.config.filter_stable_slope_limit:g} 个百分点/帧，"
                    f"极差 ≤ {self.config.filter_stable_range_limit:g} 个百分点；"
                    f"透过率低于 {100 - self.config.air_tolerance:g}% 时自动识别为插片"
                )
                baseline_values = self._collect_air_baseline(
                    ser,
                    target_indices,
                    self.config.air_warmup_seconds,
                    air_window_count,
                    air_cv_limit,
                    air_slope_limit,
                    air_range_limit,
                )
                if not self._running:
                    return
                ser.reset_input_buffer()
                buffer.clear()
                self.log_added.emit(
                    "空气基底采集完成，已清空基底阶段数据；已开始持续跟踪稳定空气基底，现在可以插入滤光片"
                )
                self.status_changed.emit("状态：基底完成，可以插入滤光片")
            else:
                for idx in target_indices:
                    self.channel_state_changed.emit(idx, "ready")
                self.status_changed.emit("状态：空气记录模式，正在持续采样")

            stable_windows = [deque(maxlen=filter_window_count) for _ in range(24)]
            air_windows = [deque(maxlen=air_window_count) for _ in range(24)]
            waiting_for_air = [False for _ in range(24)]
            filter_detected = [False for _ in range(24)]
            measurement_baselines = [None for _ in range(24)]
            measurement_air_means = [None for _ in range(24)]
            measurement_calibration_factors = [None for _ in range(24)]

            while self._running:
                data = ser.read(4096)
                if data:
                    buffer.extend(data)

                for frame in parse_frames(buffer):
                    values = parse_absorbance_frame(frame)
                    if values is None:
                        continue

                    self.values_received.emit(values)
                    try:
                        append_raw_csv_row(raw_csv_file, raw_csv_writer, values, target_indices)
                    except (OSError, PermissionError, ValueError) as exc:
                        raise SaveOutputError(f"原始 CSV 缓存保存失败: {exc}") from exc
                    row_count += 1

                    if not filter_mode:
                        self.row_status_changed.emit(row_count, stable_count)
                        continue

                    stable_changed = False
                    for idx in target_indices:
                        value = values[idx]
                        if value <= ZERO_VALUE_EPSILON:
                            stable_windows[idx].clear()
                            air_windows[idx].clear()
                            if waiting_for_air[idx]:
                                continue
                            if filter_detected[idx]:
                                filter_detected[idx] = False
                                measurement_baselines[idx] = None
                                measurement_air_means[idx] = None
                                measurement_calibration_factors[idx] = None
                                self.channel_state_changed.emit(idx, "ready")
                            continue

                        if waiting_for_air[idx]:
                            air_ratio = value / baseline_values[idx]
                            if abs(air_ratio - 1.0) > air_tolerance:
                                air_windows[idx].clear()
                                continue

                            air_windows[idx].append(value)
                            if len(air_windows[idx]) < air_window_count:
                                continue

                            if not air_window_is_stable(
                                list(air_windows[idx]),
                                baseline_values[idx],
                                air_cv_limit,
                                air_slope_limit,
                                air_range_limit,
                                air_window_count,
                            ):
                                continue

                            baseline_values[idx] = round(sum(air_windows[idx]) / air_window_count, 6)
                            waiting_for_air[idx] = False
                            filter_detected[idx] = False
                            measurement_baselines[idx] = None
                            measurement_air_means[idx] = None
                            measurement_calibration_factors[idx] = None
                            stable_windows[idx].clear()
                            air_windows[idx].clear()
                            self.baseline_value_changed.emit(idx, baseline_values[idx])
                            self.channel_state_changed.emit(idx, "ready")
                            self.log_added.emit(
                                f"CH{idx + 1} 已回到稳定空气，空气基底更新为 {baseline_values[idx]:.6f}，"
                                "已恢复持续基底跟踪，可再次插入滤光片"
                            )
                            continue

                        active_baseline = baseline_values[idx]
                        if not filter_detected[idx]:
                            air_ratio = value / active_baseline
                            if abs(air_ratio - 1.0) <= air_tolerance:
                                stable_windows[idx].clear()
                                air_windows[idx].append(value)
                                if len(air_windows[idx]) >= air_window_count:
                                    if air_window_is_stable(
                                        list(air_windows[idx]),
                                        active_baseline,
                                        air_cv_limit,
                                        air_slope_limit,
                                        air_range_limit,
                                        air_window_count,
                                    ):
                                        new_baseline = round(sum(air_windows[idx]) / air_window_count, 6)
                                        if new_baseline != baseline_values[idx]:
                                            baseline_values[idx] = new_baseline
                                            self.baseline_value_changed.emit(idx, new_baseline)
                                continue

                            air_windows[idx].clear()
                            if not filter_is_inserted(value, active_baseline, air_tolerance):
                                stable_windows[idx].clear()
                                continue

                            calibration = channel_gain_calibration(baseline_values, target_indices, idx)
                            if calibration is None:
                                stable_windows[idx].clear()
                                continue

                            air_mean, calibration_factor = calibration
                            measurement_baselines[idx] = active_baseline
                            measurement_air_means[idx] = air_mean
                            measurement_calibration_factors[idx] = calibration_factor
                            measurement_baseline = active_baseline
                            ratio = value / measurement_baseline
                            filter_detected[idx] = True
                            self.channel_state_changed.emit(idx, "detecting")
                            self.log_added.emit(
                                f"CH{idx + 1} 检测到滤光片: 当前值 {value:.6f}, "
                                f"已冻结空气本底 {measurement_baseline:.6f}, "
                                f"空气均值 {air_mean:.6f}, 校准系数 {calibration_factor:.6f}, "
                                f"比例 {ratio * 100:.2f}%"
                            )
                        else:
                            measurement_baseline = measurement_baselines[idx] or active_baseline
                            ratio = value / measurement_baseline
                            if not filter_is_inserted(value, measurement_baseline, air_tolerance):
                                stable_windows[idx].clear()
                                filter_detected[idx] = False
                                measurement_baselines[idx] = None
                                measurement_air_means[idx] = None
                                measurement_calibration_factors[idx] = None
                                self.channel_state_changed.emit(idx, "ready")
                                continue

                        stable_windows[idx].append(value)
                        if len(stable_windows[idx]) < filter_window_count:
                            continue

                        window_values = list(stable_windows[idx])
                        baseline_value = measurement_baselines[idx] or baseline_values[idx]
                        if filter_window_is_stable(
                            window_values,
                            baseline_value,
                            filter_cv_limit,
                            filter_slope_limit,
                            filter_range_limit,
                            filter_window_count,
                        ):
                            final_value = round(sum(window_values) / filter_window_count, 6)
                            air_mean = measurement_air_means[idx]
                            calibration_factor = measurement_calibration_factors[idx]
                            calibrated_value = round(
                                calibrate_intensity(final_value, calibration_factor),
                                6,
                            )
                            ratio = final_value / baseline_value
                            absorbance = absorbance_from_transmittance(ratio)
                            if absorbance is None:
                                stable_windows[idx].clear()
                                continue
                            stable_values[idx].append(final_value)
                            diagnostics = window_diagnostics(window_values, baseline_value)
                            stable_records[idx].append(
                                {
                                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3],
                                    "wavelength": self.config.wavelength,
                                    "attempt": len(stable_values[idx]),
                                    "baseline": baseline_value,
                                    "air_mean": air_mean,
                                    "calibration_factor": calibration_factor,
                                    "stable_value": final_value,
                                    "calibrated_value": calibrated_value,
                                    "ratio": ratio,
                                    "absorbance": absorbance,
                                    "sample_count": diagnostics["sample_count"],
                                    "ratio_sample_std": diagnostics["ratio_sample_std"],
                                    "ratio_cv": diagnostics["ratio_cv"],
                                    "ratio_range": diagnostics["ratio_range"],
                                    "ratio_slope_per_frame": diagnostics["ratio_slope_per_frame"],
                                    "ratio_cv_limit": filter_cv_limit,
                                    "ratio_slope_limit": filter_slope_limit,
                                    "ratio_range_limit": filter_range_limit,
                                    "air_tolerance": air_tolerance,
                                    "result": "通过",
                                }
                            )
                            waiting_for_air[idx] = True
                            measurement_baselines[idx] = None
                            measurement_air_means[idx] = None
                            measurement_calibration_factors[idx] = None
                            stable_windows[idx].clear()
                            air_windows[idx].clear()
                            stable_changed = True
                            ratio_percent = ratio * 100
                            self.channel_state_changed.emit(idx, "waiting_air")
                            self.stable_value_changed.emit(
                                idx,
                                final_value,
                                calibrated_value,
                                ratio_percent,
                                absorbance,
                            )
                            self.log_added.emit(
                                f"CH{idx + 1} 滤光片稳定: 原始 {final_value:.6f}, "
                                f"校准 {calibrated_value:.6f}, 比例 {ratio_percent:.2f}%, "
                                f"吸光度 {absorbance:.6f}；请拔出滤光片"
                            )

                    if stable_changed:
                        stable_count = sum(bool(values) for values in stable_values)
                        try:
                            refresh_stable_excel(
                                stable_wb,
                                stable_ws,
                                stable_output_path,
                                stable_values,
                                target_indices,
                                stable_records,
                            )
                        except (OSError, PermissionError, InvalidFileException) as exc:
                            raise SaveOutputError(f"稳定值保存失败: {exc}") from exc

                    self.row_status_changed.emit(row_count, stable_count)

        except SaveOutputError as exc:
            had_error = True
            final_status = f"状态：保存异常：{exc}"
            if self._running:
                self.status_changed.emit(final_status)
                self.log_added.emit(f"保存异常: {exc}")
                self.log_added.emit("请确认 Excel/WPS 没有打开输出文件，资源管理器未预览该文件，并避免杀毒或同步软件占用 data 目录")
        except Exception as exc:
            had_error = True
            final_status = f"状态：串口异常：{exc}"
            if self._running:
                self.status_changed.emit(final_status)
                self.log_added.emit(f"串口异常: {exc}")
        finally:
            self._running = False
            try:
                if ser is not None and ser.is_open and not self.config.keep_light and not self.config.no_start:
                    ser.write(stop_cmd)
                    self.log_added.emit(f"已发送关闭命令: {frame_to_hex(stop_cmd)}")
            except Exception as exc:
                self.log_added.emit(f"发送关闭命令失败: {exc}")
            finally:
                if ser is not None and ser.is_open:
                    ser.close()

            try:
                raw_csv_file.close()
            except OSError as exc:
                self.log_added.emit(f"关闭原始 CSV 缓存失败: {exc}")

            should_keep_raw = row_count > 0
            should_keep_stable = filter_mode and stable_count > 0
            raw_excel_saved = False

            if should_keep_raw:
                try:
                    raw_csv_to_excel(raw_csv_path, raw_output_path)
                    raw_excel_saved = True
                except (OSError, PermissionError, ValueError, InvalidFileException) as exc:
                    had_error = True
                    final_status = f"状态：保存异常：原始 Excel 生成失败: {exc}"
                    self.log_added.emit(f"原始 Excel 生成失败: {exc}")

            if filter_mode and stable_wb is not None and should_keep_stable:
                try:
                    refresh_stable_excel(
                        stable_wb,
                        stable_ws,
                        stable_output_path,
                        stable_values,
                        target_indices,
                        stable_records,
                    )
                except (OSError, PermissionError, InvalidFileException) as exc:
                    self.log_added.emit(f"保存稳定值失败: {exc}")

            if should_keep_raw and raw_excel_saved:
                self.log_added.emit(f"原始数据已保存至: {os.path.abspath(raw_output_path)}")
            if should_keep_raw:
                self.log_added.emit(f"原始 CSV 缓存已保留: {os.path.abspath(raw_csv_path)}")
            if filter_mode and should_keep_stable:
                self.log_added.emit(f"稳定值已保存至: {os.path.abspath(stable_output_path)}")

            if had_error or not should_keep_raw or (filter_mode and not should_keep_stable):
                self._cleanup_empty_outputs(
                    output_dir,
                    raw_output_path,
                    stable_output_path if filter_mode else None,
                    keep_raw=should_keep_raw,
                    keep_stable=should_keep_stable,
                    raw_csv_path=raw_csv_path,
                )
            self.status_changed.emit(final_status)
            self.test_finished.emit()

    def _cleanup_empty_outputs(
        self,
        output_dir: str,
        raw_output_path: str,
        stable_output_path: str | None,
        keep_raw: bool,
        keep_stable: bool,
        raw_csv_path: str | None = None,
    ) -> None:
        removed_files = []

        if not keep_raw and self._remove_file(raw_output_path):
            removed_files.append(os.path.basename(raw_output_path))
        if raw_csv_path is not None and not keep_raw and self._remove_file(raw_csv_path):
            removed_files.append(os.path.basename(raw_csv_path))

        if stable_output_path is not None and not keep_stable and self._remove_file(stable_output_path):
            removed_files.append(os.path.basename(stable_output_path))

        if removed_files:
            self.log_added.emit(f"未产生有效数据，已删除空文件: {', '.join(removed_files)}")

        try:
            temp_csv_dir = os.path.join(output_dir, "_temp_csv")
            if os.path.isdir(temp_csv_dir) and not os.listdir(temp_csv_dir):
                shutil.rmtree(temp_csv_dir)
            if os.path.isdir(output_dir) and not os.listdir(output_dir):
                shutil.rmtree(output_dir)
                self.log_added.emit(f"本次记录目录为空，已删除: {os.path.abspath(output_dir)}")
        except OSError as exc:
            self.log_added.emit(f"清理空记录目录失败: {exc}")

    def _remove_file(self, path: str) -> bool:
        try:
            if os.path.exists(path):
                os.remove(path)
                return True
        except OSError as exc:
            self.log_added.emit(f"删除空文件失败: {os.path.abspath(path)}: {exc}")
        return False

    def _collect_air_baseline(
        self,
        ser,
        target_indices: list[int],
        warmup_seconds: float,
        air_window_count: int,
        air_cv_limit: float,
        air_slope_limit: float,
        air_range_limit: float,
    ):
        buffer = bytearray()
        baseline_windows = [deque(maxlen=air_window_count) for _ in range(24)]
        baseline_values = [None for _ in range(24)]
        last_done_count = -1
        warmup_started_at = time.monotonic()
        warmup_finished = warmup_seconds <= 0

        self.baseline_progress_changed.emit(0, len(target_indices))
        if not warmup_finished:
            self.status_changed.emit(
                f"状态：请保持空气状态，空气基底至少暖机 {warmup_seconds:g} 秒"
            )
            self.log_added.emit(f"空气基底暖机开始，至少等待 {warmup_seconds:g} 秒")
        while self._running:
            data = ser.read(4096)
            if data:
                buffer.extend(data)

            for frame in parse_frames(buffer):
                values = parse_absorbance_frame(frame)
                if values is None:
                    continue

                self.values_received.emit(values)
                if not warmup_finished:
                    if time.monotonic() - warmup_started_at < warmup_seconds:
                        continue
                    warmup_finished = True
                    self.status_changed.emit("状态：请保持空气状态，正在判定空气基底稳定性")
                    self.log_added.emit("空气基底最短暖机时间已到，开始稳定窗口判定")

                for idx in target_indices:
                    if baseline_values[idx] is not None:
                        continue

                    value = values[idx]
                    if value <= ZERO_VALUE_EPSILON:
                        baseline_windows[idx].clear()
                        continue

                    baseline_windows[idx].append(value)
                    if len(baseline_windows[idx]) < air_window_count:
                        continue

                    window_values = list(baseline_windows[idx])
                    if not air_window_is_stable(
                        window_values,
                        None,
                        air_cv_limit,
                        air_slope_limit,
                        air_range_limit,
                        air_window_count,
                    ):
                        continue

                    baseline_values[idx] = round(sum(window_values) / air_window_count, 6)
                    self.baseline_value_changed.emit(idx, baseline_values[idx])
                    self.channel_state_changed.emit(idx, "ready")

                done_count = sum(baseline_values[idx] is not None for idx in target_indices)
                if done_count != last_done_count:
                    self.baseline_progress_changed.emit(done_count, len(target_indices))
                    last_done_count = done_count

                if done_count == len(target_indices):
                    return baseline_values

        return baseline_values
