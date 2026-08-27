import configparser
import os

from app_paths import get_app_dir


CONFIG_FILE_NAME = "config.ini"

DEFAULT_OPTIONS = {
    "port": "",
    "wavelength": 520,
    "channel_group": 0,
    "test_mode": "filter",
    "output": "串口数据记录.xlsx",
    "stable_output": "滤光片稳定值.xlsx",
    "filter_stable_window_frames": 10,
    "filter_stable_cv_limit": 0.25,
    "filter_stable_slope_limit": 0.02,
    "filter_stable_range_limit": 0.25,
    "air_tolerance": 5.0,
    "air_warmup_seconds": 30.0,
    "air_stable_window_frames": 20,
    "air_stable_cv_limit": 0.25,
    "air_stable_slope_limit": 0.03,
    "air_stable_range_limit": 0.50,
    "no_start": False,
    "keep_light": False,
}


def get_config_path():
    return os.path.join(get_app_dir(), CONFIG_FILE_NAME)


def load_config(path: str):
    if not os.path.exists(path):
        return {}

    with open(path, "r", encoding="utf-8") as file:
        text = file.read()

    if "[" not in text:
        text = "[settings]\n" + text

    parser = configparser.ConfigParser()
    parser.read_string(text)
    if not parser.has_section("settings"):
        return {}

    values = {key.replace("-", "_"): value for key, value in parser.items("settings")}
    if "test_mode" not in values and "filter_ratio" in values:
        try:
            values["test_mode"] = "air" if int(values["filter_ratio"]) == 0 else "filter"
        except ValueError:
            pass
    values.pop("filter_ratio", None)
    values.pop("ratio_tolerance", None)
    return values


def merge_config(args, parser):
    config_path = get_config_path()
    config_values = load_config(config_path)

    for name, default_value in DEFAULT_OPTIONS.items():
        cli_value = getattr(args, name, None)
        if cli_value is not None:
            value = cli_value
        elif name in config_values:
            value = convert_config_value(name, config_values[name], parser)
        else:
            value = default_value

        setattr(args, name, value)

    validate_args(args, parser)
    args.config_path = config_path
    args.config_loaded = bool(config_values)
    return args


def convert_config_value(name: str, value: str, parser):
    try:
        if name in (
            "wavelength",
            "channel_group",
            "filter_stable_window_frames",
            "air_stable_window_frames",
        ):
            return int(value)
        if name in (
            "filter_stable_cv_limit",
            "filter_stable_slope_limit",
            "filter_stable_range_limit",
            "air_tolerance",
            "air_warmup_seconds",
            "air_stable_cv_limit",
            "air_stable_slope_limit",
            "air_stable_range_limit",
        ):
            return float(value)
        if name in ("no_start", "keep_light"):
            return parse_bool(value)
        if name == "test_mode":
            return value.strip().lower()
    except ValueError:
        parser.error(f"config.ini 中 {name} 的值无效: {value}")

    return value


def parse_bool(value: str):
    normalized = value.strip().lower()
    if normalized in ("1", "true", "yes", "on", "y"):
        return True
    if normalized in ("0", "false", "no", "off", "n"):
        return False
    raise ValueError(value)


def validate_args(args, parser):
    if args.wavelength not in (410, 460, 520, 550, 590, 630):
        parser.error("wavelength 必须是 410/460/520/550/590/630")
    if args.channel_group not in (0, 1, 2):
        parser.error("channel_group 必须是 0/1/2")
    if args.test_mode not in ("air", "filter"):
        parser.error("test_mode 必须是 air 或 filter")
    if args.filter_stable_window_frames < 2:
        parser.error("filter_stable_window_frames 必须大于等于 2")
    if args.filter_stable_cv_limit < 0:
        parser.error("filter_stable_cv_limit 不能小于 0")
    if args.filter_stable_slope_limit < 0:
        parser.error("filter_stable_slope_limit 不能小于 0")
    if args.filter_stable_range_limit < 0:
        parser.error("filter_stable_range_limit 不能小于 0")
    if not 0 <= args.air_tolerance < 100:
        parser.error("air_tolerance 必须大于等于 0 且小于 100")
    if args.air_warmup_seconds < 0:
        parser.error("air_warmup_seconds 不能小于 0")
    if args.air_stable_window_frames < 2:
        parser.error("air_stable_window_frames 必须大于等于 2")
    if args.air_stable_cv_limit < 0:
        parser.error("air_stable_cv_limit 不能小于 0")
    if args.air_stable_slope_limit < 0:
        parser.error("air_stable_slope_limit 不能小于 0")
    if args.air_stable_range_limit < 0:
        parser.error("air_stable_range_limit 不能小于 0")
