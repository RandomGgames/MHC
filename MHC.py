import hashlib
import json
import logging
import os
import pathlib
import send2trash
import shutil
import socket
import sys
import time
import toml
import traceback
import typing

from datetime import datetime

logger = logging.getLogger(__name__)


def load_config(config_path: str = "config.toml") -> dict:
    logger.debug("Loading config...")
    required_keys = {"cache", "media_dir"}
    logger.debug("Reading config file...")
    config = toml.load(config_path)
    required_keys = {
        "cache": str,
        "media_dir": str,
        "media_extensions": list,
        "ignore_files_with": list
    }

    logger.debug("Validating config...")
    for key, expected_type in required_keys.items():
        if key not in config or not isinstance(config[key], expected_type):
            raise ValueError(f"config.toml is missing or has incorrect type for key '{key}' (expected {expected_type.__name__})")
    if not all(key in config for key in required_keys):
        raise ValueError(f"config.toml is missing required key(s): {', '.join(sorted(list(set(required_keys) - set(config.keys()))))}")
    logger.debug("Config loaded successfully.")
    return config


def load_cache(config: dict) -> dict:
    logger.debug("Loading cache...")
    if os.path.exists(config["cache"]):
        try:
            with open(config["cache"]) as f:
                cache = json.load(f)
                logger.debug(f"Loaded cache.")
        except json.JSONDecodeError as e:
            logger.error(f"Failed to load cache from {config['cache']} due to {e}.")
            raise
    else:
        logger.debug(f"Cache file {config['cache']} does not exist.")
        cache = {}
        save_cache(cache, config)
    return cache


def save_cache(cache: dict, config: dict) -> None:
    logger.debug("Saving cache...")
    try:
        cache_dir = pathlib.Path(config["cache"]).parent
        if not cache_dir.exists():
            cache_dir.mkdir(parents=True)
            logger.debug(f"Created cache directory {cache_dir}.")
        with open(config["cache"], "w") as f:
            json.dump(cache, f, indent=4)
            logger.debug(f"Saved cache.")
    except Exception as e:
        logger.error(f"Failed to save cache to {config['cache']} due to {e}.")
        raise


def generate_hash(file_path: typing.Union[str, pathlib.Path], algorithm="sha256") -> str:
    logger.debug(f"Generating hash for {file_path}...")
    with open(file_path, "rb") as f:
        digest = hashlib.file_digest(f, algorithm)
    hash = digest.hexdigest()
    logger.debug(f"Generated hash '{hash}'.")
    return hash


def get_media_files(config: dict) -> typing.Iterable[pathlib.Path]:
    media_dir = pathlib.Path(config["media_dir"]).resolve()
    logger.debug(f"Searching for media files in '{media_dir}'...")
    for file in media_dir.rglob("*"):
        if file.is_file():
            # logger.debug(f"Found file: {file}")
            if file.suffix[1:] in config["media_extensions"]:
                # logger.debug(f"File has media extension: {file}")
                if not any(phrase.lower() in file.stem.lower() for phrase in config["ignore_files_with"]):
                    # logger.debug(f"File does not have any ignore phrases: {file}")
                    yield file


def get_file_data(file_path: typing.Union[str, pathlib.Path]) -> tuple[int, int, int]:
    logger.debug(f"Getting file data for {file_path}...")
    file_path = pathlib.Path(file_path)
    modified_time = file_path.stat().st_mtime_ns
    created_time = file_path.stat().st_birthtime_ns
    size = file_path.stat().st_size
    logger.debug(f"Got file data: {modified_time = }, {created_time = }, {size = }")
    return modified_time, created_time, size


def main() -> None:
    config = load_config()
    cache = load_cache(config)

    logger.debug(f"{cache=}")

    for path in get_media_files(config):
        hash = generate_hash(path)
        modified_time, created_time, size = get_file_data(path)

def setup_logging(
        logger: logging.Logger,
        log_file_path: typing.Union[str, pathlib.Path],
        number_of_logs_to_keep: typing.Union[int, None] = None,
        console_logging_level: int = logging.DEBUG,
        file_logging_level: int = logging.DEBUG,
        log_message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s] [%(name)s]: %(message)s",
        date_format: str = "%Y-%m-%d %H:%M:%S") -> None:
    # Ensure log_dir is a Path object
    log_file_path = pathlib.Path(log_file_path)
    log_dir = log_file_path.parent
    log_dir.mkdir(parents=True, exist_ok=True)  # Create logs dir if it does not exist

    # Limit # of logs in logs folder
    if number_of_logs_to_keep is not None:
        log_files = sorted([f for f in log_dir.glob("*.log")], key=lambda f: f.stat().st_mtime)
        if len(log_files) >= number_of_logs_to_keep:
            for file in log_files[:len(log_files) - number_of_logs_to_keep + 1]:
                file.unlink()

    logger.setLevel(file_logging_level)  # Set the overall logging level

    # File Handler for date-based log file
    file_handler_date = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler_date.setLevel(file_logging_level)
    file_handler_date.setFormatter(logging.Formatter(log_message_format, datefmt=date_format))
    logger.addHandler(file_handler_date)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(logging.Formatter(log_message_format, datefmt=date_format))
    logger.addHandler(console_handler)

    # Set specific logging levels if needed
    # logging.getLogger("requests").setLevel(logging.INFO)


if __name__ == "__main__":
    pc_name = socket.gethostname()
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    script_name = pathlib.Path(__file__).stem
    log_dir = pathlib.Path(f"{script_name} Logs")
    log_file_name = f"{timestamp}_{pc_name}.log"
    log_file_path = log_dir / log_file_name
    setup_logging(logger, log_file_path, number_of_logs_to_keep=10, log_message_format="%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s")

    error = 0
    try:
        start_time = time.perf_counter()
        logger.info("Starting operation...")
        main()
        end_time = time.perf_counter()
        duration = end_time - start_time
        logger.info(f"Completed operation in {duration:.4f}s.")
    except Exception as e:
        logger.warning(f"A fatal error has occurred: {repr(e)}\n{traceback.format_exc()}")
        error = 1
    finally:
        sys.exit(error)
