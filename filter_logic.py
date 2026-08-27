import math


WAVELENGTH_CODES = {
    410: 0,
    460: 1,
    520: 2,
    550: 3,
    590: 4,
    630: 5,
}

WAVELENGTH_COLORS = {
    410: "紫色",
    460: "蓝色",
    520: "绿色",
    550: "绿色",
    590: "橙色",
    630: "红色",
}

ZERO_VALUE_EPSILON = 1e-7


def channel_gain_calibration(
    baseline_values: list[float | None],
    target_indices: list[int],
    channel_index: int,
):
    """按所选通道空气基底算术平均值计算当前通道的强度校准系数。"""
    selected_baselines = [baseline_values[idx] for idx in target_indices]
    if (
        not selected_baselines
        or any(value is None or value <= ZERO_VALUE_EPSILON for value in selected_baselines)
        or channel_index not in target_indices
    ):
        return None

    air_mean = sum(selected_baselines) / len(selected_baselines)
    channel_baseline = baseline_values[channel_index]
    return air_mean, air_mean / channel_baseline


def calibrate_intensity(value: float, calibration_factor: float):
    """使用插片时冻结的通道增益系数校准滤光片稳定强度。"""
    return value * calibration_factor


def absorbance_from_transmittance(transmittance: float):
    """按 A=-log10(T) 将透过率比例转换为吸光度。"""
    if transmittance <= 0:
        return None
    return -math.log10(transmittance)


def window_diagnostics(values: list[float], baseline: float):
    """计算空气和滤光片判稳共用的归一化窗口指标。"""
    sample_count = len(values)
    if not values or baseline <= ZERO_VALUE_EPSILON:
        return {
            "sample_count": sample_count,
            "ratio_sample_std": None,
            "ratio_cv": None,
            "ratio_range": None,
            "ratio_slope_per_frame": None,
        }

    ratios = [value / baseline for value in values]
    ratio_mean = sum(ratios) / sample_count
    ratio_range = max(ratios) - min(ratios)

    if sample_count < 2:
        ratio_sample_std = None
        ratio_cv = None
        ratio_slope = None
    else:
        ratio_sample_std = (
            sum((ratio - ratio_mean) ** 2 for ratio in ratios) / (sample_count - 1)
        ) ** 0.5
        ratio_cv = ratio_sample_std / abs(ratio_mean) if ratio_mean else None

        frame_mean = (sample_count - 1) / 2
        slope_divisor = sum((frame - frame_mean) ** 2 for frame in range(sample_count))
        ratio_slope = (
            sum((frame - frame_mean) * (ratio - ratio_mean) for frame, ratio in enumerate(ratios))
            / slope_divisor
        )

    return {
        "sample_count": sample_count,
        "ratio_sample_std": ratio_sample_std,
        "ratio_cv": ratio_cv,
        "ratio_range": ratio_range,
        "ratio_slope_per_frame": ratio_slope,
    }


def air_window_is_stable(
    values: list[float],
    baseline: float | None,
    cv_limit: float,
    slope_limit: float,
    range_limit: float,
    required_count: int,
):
    """按归一化 CV、斜率和极差判断空气窗口是否稳定。"""
    if len(values) < required_count:
        return False

    reference_baseline = baseline
    if reference_baseline is None:
        reference_baseline = sum(values) / len(values)
    if reference_baseline <= ZERO_VALUE_EPSILON:
        return False

    diagnostics = window_diagnostics(values, reference_baseline)
    ratio_cv = diagnostics["ratio_cv"]
    ratio_slope = diagnostics["ratio_slope_per_frame"]
    ratio_range = diagnostics["ratio_range"]
    return (
        ratio_cv is not None
        and ratio_slope is not None
        and ratio_range is not None
        and ratio_cv <= cv_limit
        and abs(ratio_slope) <= slope_limit
        and ratio_range <= range_limit
    )


def filter_window_is_stable(
    values: list[float],
    baseline: float,
    cv_limit: float,
    slope_limit: float,
    range_limit: float,
    required_count: int,
):
    """按透过率 CV、斜率和极差判断滤光片窗口是否稳定。"""
    if len(values) < required_count or baseline <= ZERO_VALUE_EPSILON:
        return False

    diagnostics = window_diagnostics(values, baseline)
    ratio_cv = diagnostics["ratio_cv"]
    ratio_slope = diagnostics["ratio_slope_per_frame"]
    ratio_range = diagnostics["ratio_range"]
    return (
        ratio_cv is not None
        and ratio_slope is not None
        and ratio_range is not None
        and ratio_cv <= cv_limit
        and abs(ratio_slope) <= slope_limit
        and ratio_range <= range_limit
    )


def filter_is_inserted(value: float, baseline: float, air_tolerance: float):
    """透过率低于空气范围下限时，判定为滤光片已插入。"""
    if value <= ZERO_VALUE_EPSILON or baseline <= ZERO_VALUE_EPSILON:
        return False
    return value / baseline < 1.0 - air_tolerance


def channel_indices(channel_group: int):
    if channel_group == 1:
        return list(range(12))
    if channel_group == 2:
        return list(range(12, 24))
    return list(range(24))


def channel_group_label(channel_group: int):
    if channel_group == 1:
        return "CH1-CH12"
    if channel_group == 2:
        return "CH13-CH24"
    return "CH1-CH24"
