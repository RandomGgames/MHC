"""
A python script that hashes media and assists in removing duplicates.
"""

import bisect
import hashlib
import json
import logging
import os
import pathlib
import re
import shutil
import socket
import sys
import time
import tomllib
import traceback
import typing
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Iterable, Pattern

import win32com.client

import send2trash

logger = logging.getLogger(__name__)


__version__ = "1.1.0"  # Major.Minor.Patch


def read_toml(file_path: Path | str) -> dict:
    """
    Reads a TOML file and returns its contents as a dictionary.

    Args:
        file_path (Path | str): The file path of the TOML file to read.

    Returns:
        dict: The contents of the TOML file as a dictionary.

    Raises:
        FileNotFoundError: If the TOML file does not exist.
        OSError: If the file cannot be read.
        tomllib.TOMLDecodeError (or toml.TomlDecodeError): If the file is invalid TOML.
    """
    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(f"File not found: {json.dumps(str(path))}")

    try:
        # Read TOML as bytes
        with path.open("rb") as f:
            data = tomllib.load(f)  # Replace with 'toml.load(f)' if using the toml package
        return data

    except (OSError, tomllib.TOMLDecodeError):
        logger.exception(f"Failed to read TOML file: {json.dumps(str(file_path))}")
        raise


def load_cache(path: typing.Union[Path, str] = "cache.json") -> dict:
    """
    Loads a cache from the given path.

    Args:
    path (typing.Union[pathlib.Path, str], optional): The path of the cache file to load. Defaults to "cache.json".

    Returns:
    dict: The loaded cache.
    """
    logger.debug("Loading cache...")
    path = Path(path)
    if path.exists():
        try:
            logger.debug("Reading cache file...")
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                logger.debug("Read cache file.")
        except json.JSONDecodeError as e:
            logger.error("Failed to load cache from %s due to %s. Generating blank cache...", json.dumps(str(path)), e)
            data = {}
    else:
        logger.debug("Cache file %s does not exist. Generating blank cache...", json.dumps(str(path)))
        data = {}

    logger.debug("Cache loaded.")
    return data


def save_cache(data: dict, path: typing.Union[Path, str] = "cache.json") -> None:
    """
    Saves the given cache data to the given path.

    Args:
    data (dict): The cache data to save.
    path (typing.Union[pathlib.Path, str], optional): The path of the cache file to save. Defaults to "cache.json".
    """
    logger.debug("Saving cache...")
    path = Path(path)
    try:
        cache_dir = path.parent
        if not cache_dir.exists():
            cache_dir.mkdir(parents=True)
            logger.debug("Created cache directory %s.", json.dumps(str(cache_dir)))
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        logger.debug("Saved cache.")
    except Exception as e:
        logger.error("Failed to save cache to %s due to %s.", json.dumps(str(path)), e)
        raise


def find_files(root: str | Path, *, recursive: bool = True, include: list[str | Pattern] | None = None, ignore: list[str | Pattern] | None = None) -> Iterable[Path]:
    """
    Yield files in a directory with optional regex-based include and ignore filters.

    Args:
        root: Directory path to search.
        recursive: If True, search all subdirectories.
        include: List of regex strings or compiled patterns. Only files matching at least
            one pattern are included. If None, all files are included.
        ignore: List of regex strings or compiled patterns. Files matching any pattern
            are skipped.

    Yields:
        pathlib.Path objects for files that match the include/ignore criteria.

    Raises:
        FileNotFoundError: If `root` does not exist.
        ValueError: If `root` is not a directory.
    """
    root = Path(root)
    if not root.exists():
        raise FileNotFoundError(f"Root path does not exist: {root}")
    if not root.is_dir():
        raise ValueError(f"Expected a directory, got: {root}")

    # Compile regexes if needed
    include_patterns: list[Pattern] = [
        re.compile(p) if isinstance(p, str) else p for p in (include or [])
    ]
    ignore_patterns: list[Pattern] = [
        re.compile(p) if isinstance(p, str) else p for p in (ignore or [])
    ]

    iterator = root.rglob("*") if recursive else root.glob("*")
    for path in iterator:
        if not path.is_file():
            continue

        path_str = str(path)

        # Ignore if matches any ignore pattern
        if any(p.search(path_str) for p in ignore_patterns):
            continue

        # Include only if matches at least one include pattern (or include_patterns empty)
        if include_patterns and not any(p.search(path_str) for p in include_patterns):
            continue

        yield path


def get_file_data(file_path: Path | str) -> dict[str, int]:
    """
    Gets file stats safely across different Operating Systems.
    Returns times in nanoseconds.
    """
    path = Path(file_path)
    stat = path.stat()
    mtime = stat.st_mtime_ns
    try:
        # macOS/BSD
        ctime = getattr(stat, 'st_birthtime_ns', None)
        if ctime is None:
            # Windows (st_ctime is creation time on Windows)
            ctime = stat.st_ctime_ns
    except AttributeError:
        # Linux fallback (st_ctime is metadata change, not birth)
        ctime = mtime

    size = stat.st_size
    logger.debug(f"File stats for {path.name}: {mtime=}, {ctime=}, {size=}")
    return {"modified": mtime, "created": ctime, "size": size}


def generate_hash(file_path: str | Path, algorithm: str = "sha256") -> str:
    """
    Generate a hexadecimal hash for a file.

    Works with Python >=3.6. If running on Python >=3.11, uses
    `hashlib.file_digest` for optimal performance. Otherwise,
    reads the file in chunks to avoid memory issues with large files.

    Args:
        file_path: Path to the file (str or Path).
        algorithm: Hash algorithm name (e.g., 'sha256', 'md5').

    Returns:
        Hexadecimal string of the file hash.

    Raises:
        FileNotFoundError: if the file does not exist.
        ValueError: if the algorithm is not supported.
        OSError: if reading the file fails.
    """
    file_path = Path(file_path)
    if not file_path.is_file():
        raise FileNotFoundError(f"File does not exist: {file_path}")

    logger.debug(f"Generating hash for {json.dumps(str(file_path))} using {algorithm}...")
    try:
        with open(file_path, "rb") as f:
            # Python 3.11+ optimal path
            try:
                digest = hashlib.file_digest(f, algorithm)
                hexd = digest.hexdigest()
            except AttributeError:
                # Fallback for older Python: read in chunks
                h = hashlib.new(algorithm)
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(chunk)
                hexd = h.hexdigest()
    except Exception as e:
        logger.exception(f"Failed to generate hash for {file_path}: {e}")
        raise

    logger.debug(f"Generated hash {json.dumps(str(hexd))} for {file_path}")
    return hexd


def build_files_cache(cache: dict, config: dict) -> None:
    logger.debug("Building files cache...")

    changes_since_save = 0
    total_changes = 0

    for idx, file_path_obj in enumerate(get_media_files(config), start=1):
        file_path = pathlib.Path(file_path_obj).as_posix()
        try:
            modified_time, created_time, size = get_file_data(file_path)
        except Exception:
            logger.exception(f"Failed to stat file {file_path}; skipping.")
            continue

        # If file is not in cache, add it
        if file_path not in cache["files"]:
            try:
                file_hash = generate_hash(file_path)
            except Exception:
                logger.warning(f"Skipping file {file_path} due to hash/read error.")
                continue

            cache["files"][file_path] = {
                "modified_time": modified_time,
                "created_time": created_time,
                "size": size,
                "hash": file_hash
            }
            logger.debug(f"Added file to cache: {file_path}")
            changes_since_save += 1
            total_changes += 1

        else:
            old = cache["files"][file_path]
            if (modified_time != old.get("modified_time") or
                    created_time != old.get("created_time") or
                    size != old.get("size")):
                try:
                    file_hash = generate_hash(file_path)
                except Exception:
                    logger.warning(f"Skipping update of {file_path} due to hash/read error.")
                    continue

                cache["files"][file_path] = {
                    "modified_time": modified_time,
                    "created_time": created_time,
                    "size": size,
                    "hash": file_hash
                }
                logger.debug(f"Updated file in cache: {file_path}")
                changes_since_save += 1
                total_changes += 1
            else:
                logger.debug(f"File has not changed. Skipping: {file_path}")

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

    hash_groups = cache.get("hashes", {})

    # Step 1: Build mapping of final desired names
    rename_plan: dict[str, str] = {}
    for file_hash, entries in hash_groups.items():
        for i, entry in enumerate(entries, start=1):
            file_path = list(entry.keys())[0]
            ext = pathlib.Path(file_path).suffix.lower()
            if i == 1:
                new_name = f"{file_hash}{ext}"
            else:
                new_name = f"{file_hash} (copy {i}){ext}"
            rename_plan[file_path] = str(pathlib.Path(file_path).with_name(new_name))

    # Normalize planned destination paths to a set of POSIX strings for conflict checking
    planned_destinations = {pathlib.Path(p).as_posix() for p in rename_plan.values()}

    # Step 2: Resolve conflicts with temporary renames
    for old, new in list(rename_plan.items()):
        new_path = pathlib.Path(new)
        # If target exists on disk AND is not one of our intended destinations, move it aside
        if new_path.exists() and new_path.as_posix() not in planned_destinations:
            tmp_name = str(
                new_path.with_name(
                    f"{new_path.stem}_{int(time.time() * 1000)}{new_path.suffix}"
                )
            )
            logger.debug(f"Temporarily renaming '{new}' → '{tmp_name}'")
            try:
                new_path.rename(tmp_name)
            except OSError:
                shutil.move(str(new_path), tmp_name)

            # Update cache: remove old key, insert new (use POSIX keys)
            old_key = new_path.as_posix()
            new_key = pathlib.Path(tmp_name).as_posix()
            old_data = cache["files"].pop(old_key, None)
            if old_data is not None:
                cache["files"][new_key] = old_data
            else:
                logger.warning(
                    f"Tried to update cache for temporarily moved file {old_key}, but it was not in cache."
                )

    # Step 3: Apply renames according to final rename plan
    for old, new in rename_plan.items():
        old_path = pathlib.Path(old).resolve()
        new_path = pathlib.Path(new).resolve()

        if not old_path.exists():
            logger.warning(f"Skipping rename '{old}' → '{new}': source missing.")
            continue

        if old_path.as_posix() == new_path.as_posix():
            logger.debug(f"Skipping rename of '{old}'. Already correct.")
            continue

        if new_path.exists() and new_path.as_posix() not in planned_destinations:
            logger.warning(f"Cannot rename '{old}' → '{new}': destination exists.")
            continue

        logger.debug(f"Renaming '{old}' → '{new}'")
        try:
            old_path.rename(new_path)
        except OSError:
            # fallback across filesystems
            shutil.move(str(old_path), str(new_path))

        # Update cache safely (use POSIX keys)
        old_key = old_path.as_posix()
        new_key = new_path.as_posix()
        entry = cache["files"].pop(old_key, None)
        if entry is not None:
            cache["files"][new_key] = entry
        else:
            logger.warning(f"Cache entry for '{old_key}' missing when renaming to '{new_key}'.")

        logger.info("Rename applied.")

    save_cache(cache, config)


def delete_duplicate_files(cache: dict, config: dict) -> None:
    removed = 0

    for file_hash, entries in list(cache.get("hashes", {}).items()):
        if len(entries) <= 1:
            continue

        keep_file = list(entries[0].keys())[0]

        for entry in entries[1:]:
            file_path = list(entry.keys())[0]
            p = pathlib.Path(file_path)

            try:
                if p.exists():
                    logger.info(f"Sending duplicate to trash: {file_path}")
                    send2trash.send2trash(str(p))
                else:
                    logger.debug(f"Duplicate missing on disk (skip): {file_path}")
            except Exception:
                logger.exception(f"Failed to send {file_path} to trash.")

            cache["files"].pop(file_path, None)
            removed += 1

    logger.info(f"Removed {removed} duplicate files.")
    save_cache(cache, config)


def main(config) -> None:
    cache = load_cache(config)
    logger.debug(f"{cache=}")

    # build_files_cache(cache, config)
    # build_hashes_cache(cache, config)
    # delete_duplicate_files(cache, config)
    # build_hashes_cache(cache, config)
    # rename_files_to_hashes(cache, config)


def format_duration_long(duration_seconds: float) -> str:
    """
    Format duration in a human-friendly way, showing only the two largest non-zero units.
    For durations >= 1s, do not show microseconds or nanoseconds.
    For durations >= 1m, do not show milliseconds.
    """
    ns = int(duration_seconds * 1_000_000_000)
    units = [
        ("y", 365 * 24 * 60 * 60 * 1_000_000_000),
        ("mo", 30 * 24 * 60 * 60 * 1_000_000_000),
        ("d", 24 * 60 * 60 * 1_000_000_000),
        ("h", 60 * 60 * 1_000_000_000),
        ("m", 60 * 1_000_000_000),
        ("s", 1_000_000_000),
        ("ms", 1_000_000),
        ("us", 1_000),
        ("ns", 1),
    ]
    parts = []
    for name, factor in units:
        value, ns = divmod(ns, factor)
        if value:
            parts.append(f"{value}{name}")
        if len(parts) == 2:
            break
    if not parts:
        return "0s"
    return "".join(parts)


def enforce_max_log_count(dir_path: Path | str, max_count: int | None, script_name: str) -> None:
    """Keep only the N most recent logs for this script."""
    if max_count is None or max_count <= 0:
        return

    dir_path = Path(dir_path)

    # Get all logs for this script, sorted by name (which is our timestamp)
    # Newest will be at the end of the list
    files = sorted([f for f in dir_path.glob(f"*{script_name}*.log") if f.is_file()])

    # If we have more than the limit, calculate how many to delete
    if len(files) > max_count:
        to_delete = files[:-max_count]  # Everything except the last N files
        for f in to_delete:
            try:
                f.unlink()
                logger.debug(f"Deleted old log: {f.name}")
            except OSError as e:
                logger.error(f"Failed to delete {f.name}: {e}")


def setup_logging(
        logger_obj: logging.Logger,
        file_path: Path | str,
        script_name: str,
        max_log_files: int | None = None,
        console_logging_level: int = logging.DEBUG,
        file_logging_level: int = logging.DEBUG,
        message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s",
        date_format: str = "%Y-%m-%d %H:%M:%S"
) -> None:
    """
    Set up logging for a script.

    Args:
    logger_obj (logging.Logger): The logger object to configure.
    file_path (Path | str): The file path of the log file to write.
    max_log_files (int | None, optional): The maximum total size for all logs in the folder. Defaults to None.
    console_logging_level (int, optional): The logging level for console output. Defaults to logging.DEBUG.
    file_logging_level (int, optional): The logging level for file output. Defaults to logging.DEBUG.
    message_format (str, optional): The format string for log messages. Defaults to "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s".
    date_format (str, optional): The format string for log timestamps. Defaults to "%Y-%m-%d %H:%M:%S".
    """

    file_path = Path(file_path)
    dir_path = file_path.parent
    dir_path.mkdir(parents=True, exist_ok=True)

    logger_obj.handlers.clear()
    logger_obj.setLevel(file_logging_level)

    formatter = logging.Formatter(message_format, datefmt=date_format)

    # File Handler
    file_handler = logging.FileHandler(file_path, encoding="utf-8")
    file_handler.setLevel(file_logging_level)
    file_handler.setFormatter(formatter)
    logger_obj.addHandler(file_handler)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(formatter)
    logger_obj.addHandler(console_handler)

    if max_log_files is not None:
        enforce_max_log_count(dir_path, max_log_files, script_name)


def load_config(file_path: Path | str) -> dict:
    """
    Load configuration from a TOML file.

    Args:
    file_path (Path | str): The file path of the TOML file to read.

    Returns:
    dict: The contents of the TOML file as a dictionary.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {json.dumps(str(file_path))}")
    data = read_toml(file_path)
    return data


def bootstrap():
    """
    Handles environment setup, configuration loading,
    and logging before executing the main script logic.
    """
    exit_code = 0
    try:
        # Resolve paths and configuration
        script_path = Path(__file__)
        script_name = script_path.stem
        config_path = script_path.with_name(f"{script_name}_config.toml")

        # Load settings
        config = load_config(config_path)
        logger_config = config.get("logging", {})

        # Parse log levels and formats
        console_log_level = getattr(logging, logger_config.get("console_logging_level", "INFO").upper(), logging.INFO)
        file_log_level = getattr(logging, logger_config.get("file_logging_level", "INFO").upper(), logging.INFO)
        log_message_format = logger_config.get("log_message_format", "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s] - %(message)s")

        # Setup directories and filenames
        logs_folder = Path(logger_config.get("logs_folder_name", "logs"))
        logs_folder.mkdir(parents=True, exist_ok=True)

        pc_name = socket.gethostname()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = logs_folder / f"{timestamp}__{script_name}__{pc_name}.log"

        # Initialize logging
        setup_logging(
            logger_obj=logger,
            file_path=log_path,
            script_name=script_name,
            max_log_files=logger_config.get("max_log_files"),
            console_logging_level=console_log_level,
            file_logging_level=file_log_level,
            message_format=log_message_format
        )

        exit_behavior_config = config.get("exit_behavior", {})
        pause_before_exit = exit_behavior_config.get("always_pause", False)
        pause_before_exit_on_error = exit_behavior_config.get("pause_on_error", True)

        start_ns = time.perf_counter_ns()
        logger.info(f"Script: {json.dumps(script_name)} | Version: {__version__} | Host: {json.dumps(pc_name)}")

        main(config)

        end_ns = time.perf_counter_ns()
        duration_str = format_duration_long((end_ns - start_ns) / 1e9)
        logger.info(f"Execution completed in {duration_str}.")

    except KeyboardInterrupt:
        logger.warning("Operation interrupted by user.")
        exit_code = 130
    except Exception as e:  # pylint: disable=broad-exception-caught
        # Using 'err' or 'exc' is standard; logging the traceback handles the 'broad-except'
        logger.error(f"A fatal error has occurred: {e}")
        exit_code = 1
    finally:
        for handler in logger.handlers[:]:
            handler.close()
            logger.removeHandler(handler)

    if pause_before_exit or (pause_before_exit_on_error and exit_code != 0):
        input("Press Enter to exit...")

    return exit_code


if __name__ == "__main__":
    sys.exit(bootstrap())
