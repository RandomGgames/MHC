import bisect
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
    config = toml.load(config_path)

    logger.debug("Validating config...")
    required_keys = {
        "cache": str,
        "media_dir": str,
        "media_extensions": list,
        "ignore_files_with": list
    }
    for key, expected_type in required_keys.items():
        if key not in config or not isinstance(config[key], expected_type):
            raise ValueError(
                f"config.toml is missing or has incorrect type for key '{key}' (expected {expected_type.__name__})")
    if not all(key in config for key in required_keys):
        raise ValueError(
            f"config.toml is missing required key(s): {', '.join(sorted(list(set(required_keys) - set(config.keys()))))}")
    logger.debug("Config loaded successfully.")
    return config


def load_cache(config: dict) -> dict:
    logger.debug("Loading cache...")
    if os.path.exists(config["cache"]):
        try:
            logger.debug(f"Reading cache file...")
            with open(config["cache"]) as f:
                cache = json.load(f)
                logger.debug(f"Read cache file.")
        except json.JSONDecodeError as e:
            logger.error(
                f"Failed to load cache from {config['cache']} due to {e}.")
            cache = {}
    else:
        logger.debug(f"Cache file '{config['cache']}' does not exist. Generating new cache...")
        cache = {}
    cache.setdefault("files", {})
    cache.setdefault("hashes", {})

    logger.debug("Validating cache...")
    for file_path in list(cache["files"].keys()):
        if not pathlib.Path(file_path).exists():
            logger.debug(f"Removing non-existent file {file_path} from cache.")
            del cache["files"][file_path]
    logger.debug("Cache validated successfully.")

    logger.debug("Cache loaded.")
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


def get_media_files(config: dict) -> typing.Iterable[pathlib.Path]:
    media_dir = pathlib.Path(config["media_dir"]).resolve()
    logger.debug(f"Searching for media files in '{media_dir}'...")
    for file_path in media_dir.rglob("*"):
        if file_path.is_file():
            if file_path.suffix[1:] in config["media_extensions"]:
                if not any(phrase.lower() in file_path.stem.lower() for phrase in config["ignore_files_with"]):
                    logger.debug(f"Found media file: {file_path}")
                    yield file_path


def get_file_data(file_path: typing.Union[str, pathlib.Path]) -> tuple[int, int, int]:
    logger.debug(f"Getting file data...")
    file_path = pathlib.Path(file_path)
    modified_time = file_path.stat().st_mtime_ns
    created_time = file_path.stat().st_birthtime_ns
    size = file_path.stat().st_size
    logger.debug(f"Got file data: {modified_time=}, {created_time=}, {size=}")
    return modified_time, created_time, size


def generate_hash(file_path: typing.Union[str, pathlib.Path], algorithm="sha256") -> str:
    logger.debug(f"Generating hash for {file_path}...")
    with open(file_path, "rb") as f:
        digest = hashlib.file_digest(f, algorithm)
    hash = digest.hexdigest()
    logger.debug(f"Generated hash '{hash}'.")
    return hash


def build_files_cache(cache: dict, config: dict) -> None:
    logger.debug("Building files cache...")

    changes_since_save = 0
    total_changes = 0

    for idx, file in enumerate(get_media_files(config), start=1):
        file_path = str(file)
        modified_time, created_time, size = get_file_data(file_path)

        # If file is not in cache, add it
        if file_path not in cache["files"]:
            hash = generate_hash(file_path)
            cache["files"][file_path] = {
                "modified_time": modified_time,
                "created_time": created_time,
                "size": size,
                "hash": hash
            }
            logger.debug(f"Added file to cache.")
            changes_since_save += 1
            total_changes += 1

        # If file is in cache, check if it has changed
        else:
            old = cache["files"][file_path]
            if (
                modified_time != old["modified_time"]
                or created_time != old["created_time"]
                or size != old["size"]
            ):
                hash = generate_hash(file_path)
                cache["files"][file_path] = {
                    "modified_time": modified_time,
                    "created_time": created_time,
                    "size": size,
                    "hash": hash
                }
                logger.debug(f"Updated file in cache.")
            # If file has not changed, do nothing
                changes_since_save += 1
                total_changes += 1
            else:
                logger.debug(f"File has not changed. Skipping.")

        # Periodic save every 100 changes
        if changes_since_save >= 100:
            logger.debug("Saving cache (checkpoint)...")
            save_cache(cache, config)
            changes_since_save = 0

    # Final save if any unsaved changes remain
    if changes_since_save > 0:
        logger.debug("Final save of unsaved changes...")
        save_cache(cache, config)

    logger.debug(f"Finished building cache. Total changes: {total_changes}")


def build_hashes_cache(cache: dict, config: dict) -> None:
    cache["hashes"] = {}

    for file_path, file_data in cache["files"].items():
        file_hash = file_data["hash"]

        # Sort key: created_time asc, modified_time asc, size desc
        key = (file_data["created_time"], file_data["modified_time"], -file_data["size"])
        entry = (key, file_path, file_data)

        if file_hash not in cache["hashes"]:
            cache["hashes"][file_hash] = [entry]
        else:
            keys_list = [e[0] for e in cache["hashes"][file_hash]]
            idx = bisect.bisect(keys_list, key)
            cache["hashes"][file_hash].insert(idx, entry)

    # Strip to final form
    for file_hash, entries in cache["hashes"].items():
        cache["hashes"][file_hash] = [{path: data} for (_, path, data) in entries]

    save_cache(cache, config)


def rename_files_to_hashes(cache: dict, config: dict) -> None:
    logger.debug("Renaming files to hashes...")

    # Group files by hash (cache['hashes'] is already ordered properly)
    hash_groups = cache["hashes"]

    # Step 1: Build mapping of final desired names
    rename_plan = {}
    for file_hash, entries in hash_groups.items():
        for i, entry in enumerate(entries, start=1):
            file_path = list(entry.keys())[0]
            ext = pathlib.Path(file_path).suffix.lower()
            if i == 1:
                new_name = f"{file_hash}{ext}"
            else:
                new_name = f"{file_hash} (copy {i}){ext}"
            rename_plan[file_path] = str(pathlib.Path(file_path).with_name(new_name))

    # Step 2: Resolve conflicts with temporary renames
    for old, new in list(rename_plan.items()):
        new_path = pathlib.Path(new)
        if new_path.exists() and new not in rename_plan.values():
            tmp_name = str(new_path.with_name(
                f"{new_path.stem}_{int(time.time() * 1000)}{new_path.suffix}"
            ))
            logger.debug(f"Temporarily renaming '{new}' → '{tmp_name}'")
            os.rename(new, tmp_name)

            # Update cache for the temporarily moved file
            cache["files"][tmp_name] = cache["files"].pop(str(new_path))

    # Step 3: Apply renames according to plan
    for old, new in rename_plan.items():
        if old != new:
            logger.debug(f"Renaming '{old}' → '{new}'")
            os.rename(old, new)

            # Update cache
            cache["files"][new] = cache["files"].pop(old)
            logger.debug(f"Renamed and cache updated.")
        else:
            logger.debug(f"Skipping rename of '{old}'.")

    save_cache(cache, config)


def main() -> None:
    config = load_config()
    cache = load_cache(config)
    build_files_cache(cache, config)
    build_hashes_cache(cache, config)
    rename_files_to_hashes(cache, config)


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
    # Create logs dir if it does not exist
    log_dir.mkdir(parents=True, exist_ok=True)

    # Limit # of logs in logs folder
    if number_of_logs_to_keep is not None:
        log_files = sorted([f for f in log_dir.glob("*.log")],
                           key=lambda f: f.stat().st_mtime)
        if len(log_files) >= number_of_logs_to_keep:
            for file in log_files[:len(log_files) - number_of_logs_to_keep + 1]:
                file.unlink()

    logger.setLevel(file_logging_level)  # Set the overall logging level

    # File Handler for date-based log file
    file_handler_date = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler_date.setLevel(file_logging_level)
    file_handler_date.setFormatter(logging.Formatter(
        log_message_format, datefmt=date_format))
    logger.addHandler(file_handler_date)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(logging.Formatter(
        log_message_format, datefmt=date_format))
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
    setup_logging(logger, log_file_path, number_of_logs_to_keep=10,
                  log_message_format="%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s")

    error = 0
    try:
        start_time = time.perf_counter()
        logger.info("Starting operation...")
        main()
        end_time = time.perf_counter()
        duration = end_time - start_time
        logger.info(f"Completed operation in {duration:.4f}s.")
    except Exception as e:
        logger.warning(
            f"A fatal error has occurred: {repr(e)}\n{traceback.format_exc()}")
        error = 1
    finally:
        sys.exit(error)
