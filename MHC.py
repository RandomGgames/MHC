"""
A python script that hashes media and assists in removing duplicates.
"""

import bisect
import exiftool
import hashlib
import json
import logging
import os
import re
import send2trash
import shutil
import socket
import sys
import time
import tkinter
import tkinter.messagebox
import tomllib
import traceback
import typing
import win32com.client
import zipfile
from datetime import datetime
from hachoir.metadata import extractMetadata
from hachoir.parser import createParser
from pathlib import Path
from PIL import Image, ImageTk
from PIL.ExifTags import TAGS
from typing import Iterable, Pattern


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


def get_windows_details(file_path: str | Path, max_columns: int = 512) -> dict[str, str]:
    """
    Return Windows 'Details' tab properties for a file using the Shell Property System.

    - Windows only.
    - Requires pywin32: pip install pywin32
    - Keys are localized to the OS display language, matching File Explorer.
    - Values are formatted like Explorer shows them.

    Args:
        file_path: Target file.
        max_columns: Maximum number of property columns to probe.

    Returns:
        Dict of {property_label: value} for all non-empty properties.
    """
    p = Path(file_path)
    if not p.exists():
        raise FileNotFoundError(p)
    if os.name != "nt":
        raise OSError("get_windows_details is only supported on Windows")

    # Bind to Shell
    shell = win32com.client.Dispatch("Shell.Application")
    folder = shell.NameSpace(str(p.parent))
    if folder is None:
        raise RuntimeError(f"Shell.NameSpace failed for {p.parent}")
    item = folder.ParseName(p.name)
    if item is None:
        raise RuntimeError(f"Shell.ParseName failed for {p}")

    props: dict[str, str] = {}
    blanks = 0

    for i in range(max_columns):
        header = folder.GetDetailsOf(None, i)
        if not header:
            blanks += 1
            if blanks >= 25:
                break
            continue
        blanks = 0

        value = folder.GetDetailsOf(item, i)
        if isinstance(value, str):
            value = value.strip()
        if value:
            props[str(header).strip()] = str(value)

    return props


def load_cache(path: Path = Path("cache.json")) -> dict:
    """
    Loads a cache from the given path.

    Args:
    path (typing.Union[pathlib.Path, str], optional): The path of the cache file to load. Defaults to "cache.json".

    Returns:
    dict: The loaded cache.
    """
    path = Path(path)
    logger.debug(f"Loading cache file {json.dumps(str(path))}...")

    data = {}
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                logger.debug("Loaded cache.")
        except json.JSONDecodeError:
            logger.exception("Failed to decode cache file. Using empty cache.")
        except OSError:
            logger.exception("Failed to read cache file. Using empty cache.")
    else:
        logger.info("Cache file does not exist. Using empty cache.")

    return data


def ensure_dir(path: Path) -> None:
    """Checks if a directory exists (using pathlib) and creates it if it doesn't."""
    # Ensure path is a directory in case it's a file
    path = Path(path).resolve()
    if path.is_file():
        path = path.parent

    try:
        if not path.exists():
            path.mkdir(parents=True)
            logger.debug(f"Created folder: {json.dumps(str(path))}")
    except OSError as e:
        logger.error(f"Error creating directory {json.dumps(str(path))}", e)
        raise


def save_cache(data: dict, path: Path = Path("cache.json")) -> None:
    """
    Saves the given cache data to the specified path.

    Args:
        data (dict): The cache data to save.
        path (str | Path, optional): The path of the cache file to save. Defaults to "cache.json".
    """
    path = Path(path)
    logger.debug(f"Saving cache to {json.dumps(str(path))}...")

    try:
        ensure_dir(path)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)

        logger.debug(f"Saved cache with {len(data)} entries.")
    except Exception:
        logger.exception("Failed to save cache.")
        raise


def find_files(root: Path | str, *, recursive: bool = True, include: list[str | Pattern] | None = None, ignore: list[str | Pattern] | None = None) -> Iterable[Path]:
    """
    Yield files in a directory filtered by regex-based include and ignore patterns.

    Args:
        root: Directory path to search.
        recursive: If True, search all subdirectories.
        include: List of regex strings or compiled patterns. Only files matching
            at least one pattern are included. If None, all files are included.
        ignore: List of regex strings or compiled patterns. Files matching any
            pattern are skipped.

    Yields:
        pathlib.Path objects for files matching the include/ignore criteria.

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

        # Skip files matching any ignore pattern
        if any(p.search(path_str) for p in ignore_patterns):
            continue

        # Include only if it matches at least one include pattern (or no include pattern)
        if include_patterns and not any(p.search(path_str) for p in include_patterns):
            continue

        yield path


def get_file_data(file_path: Path | str) -> tuple[int, int, int]:
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
    return mtime, ctime, size


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

    # logger.debug(f"Generating hash for {json.dumps(str(file_path))} using {algorithm}...")
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

    # logger.debug(f"Generated hash {json.dumps(str(hexd))} for {file_path}")
    return hexd


def date_string_to_unix_nanos(date_str: str) -> int:
    """
    Cleans Unicode marks, parses the date, and returns a Unix timestamp in nanoseconds.
    """
    clean_str = re.sub(r'[^\x00-\x7f]', '', date_str).strip()
    dt = datetime.strptime(clean_str, "%m/%d/%Y %I:%M %p")
    timestamp_seconds = dt.timestamp()
    return int(timestamp_seconds * 1_000_000_000)


def build_files_cache(cache: dict, cache_file: Path, media_root: Path | str, media_extensions: list[str | Pattern], save_cache_every: int = 100) -> dict:
    """
    Scan media files and update the cache with metadata and hashes.

    Features:
        - Uses `find_files` with regex-based include filters
        - Purges non-existent files before scanning
        - Saves periodically every `checkpoint` changes
        - Saves at the end if there are unsaved changes
        - Handles hash/read errors gracefully

    Args:
        cache: The existing cache dictionary.
        cache_file: Path to the cache file for saving.
        media_root: Root directory to scan for media files.
        media_extensions: List of regex strings or compiled patterns for files to include.
        checkpoint: Number of changes before saving a partial cache.
    """
    logger.info("Building files cache...")

    cache.setdefault("files", {})
    cache.setdefault("hashes", {})

    # Purge entries for missing files
    cache_file = Path(cache_file)
    cache = purge_cache(cache)

    changes_since_save = 0
    total_changes = 0

    for file_path in find_files(media_root, recursive=True, include=media_extensions):
        full_file_path = Path(file_path).resolve()

        try:
            modified_time, created_time, size = get_file_data(full_file_path)
            # properties = get_windows_details(full_file_path)
            # date_taken = properties.get("Date taken", None)
            # if date_taken is not None:
            #     date_taken = date_string_to_unix_nanos(date_taken)
        except Exception:
            logger.exception(f"Failed to get file stats for {json.dumps(str(full_file_path.as_posix()))}. Skipping.")
            continue

        # Determine if we need to add/update the cache entry
        old_entry = cache["files"].get(full_file_path.as_posix())
        needs_update = (
            old_entry is None or
            old_entry.get("modified_time") != modified_time or
            old_entry.get("created_time") != created_time or
            old_entry.get("size") != size
        )

        if needs_update:
            try:
                file_hash = generate_hash(full_file_path)
            except Exception:
                logger.warning(f"Failed to generate hash for {json.dumps(str(full_file_path.as_posix()))}. Skipping.")
                continue

            cache["files"][full_file_path.as_posix()] = {
                "modified_time": modified_time,
                "created_time": created_time,
                "size": size,
                "hash": file_hash
            }

            if old_entry is None:
                logger.debug(f"Added {json.dumps(str(full_file_path.as_posix()))} to cache.")
            else:
                logger.debug(f"Updated {json.dumps(str(full_file_path.as_posix()))} in cache.")

            changes_since_save += 1
            total_changes += 1
        else:
            logger.debug(f"No changes detected for {json.dumps(str(full_file_path.as_posix()))}.")

        # Checkpoint save every `checkpoint` changes
        if changes_since_save >= save_cache_every:
            save_cache(cache, cache_file)
            changes_since_save = 0

    # Final save for any remaining changes
    if changes_since_save > 0:
        save_cache(cache, cache_file)

    logger.info(f"Finished building files cache. Total Files: {len(cache['files'])}. Total changes: {total_changes}")
    return cache


def build_hashes_cache(cache: dict, cache_file: Path) -> dict:
    """
    Build a hashes index from cached files and update only changed hashes.

    - Builds a fresh hashes cache from cache["files"]
    - Compares per-hash with existing cache["hashes"]
    - Updates only hashes that changed
    - Logs changes per hash
    - Saves cache once at the end
    """
    logger.info("Building hashes cache...")
    cache_file = Path(cache_file)

    old_hashes = cache.get("hashes", {})
    new_hashes = {}

    # Build new hashes cache from files
    for file_path, file_data in cache.get("files", {}).items():
        file_hash = file_data.get("hash")
        if not file_hash:
            logger.warning(f"No hash found for file {file_path}; skipping.")
            continue

        key = (
            file_data["created_time"],
            file_data["modified_time"],
            -file_data["size"],
        )

        new_hashes.setdefault(file_hash, []).append((key, file_path, file_data))

    # Sort and normalize format
    for file_hash, entries in new_hashes.items():
        entries.sort(key=lambda e: e[0])
        new_hashes[file_hash] = [{path: data} for _, path, data in entries]

    # Compare and update only changed hashes
    updated_hashes = old_hashes.copy()

    for file_hash, new_entries in new_hashes.items():
        old_entries = old_hashes.get(file_hash)

        if old_entries is None:
            updated_hashes[file_hash] = new_entries
            logger.debug(f"Added new hash {file_hash}.")
        elif old_entries != new_entries:
            updated_hashes[file_hash] = new_entries
            logger.debug(f"Updated hash {file_hash}.")
        else:
            logger.debug(f"No changes for hash {file_hash}.")

    # Optional: detect removed hashes
    removed_hashes = set(old_hashes) - set(new_hashes)
    for file_hash in removed_hashes:
        updated_hashes.pop(file_hash, None)
        logger.debug(f"Removed hash {file_hash} (no longer present).")

    cache["hashes"] = updated_hashes

    save_cache(cache, cache_file)
    logger.debug(f"Finished building hashes cache. Total Hashes: {len(cache['hashes'])}. Total changes: {len(updated_hashes)}")
    return cache


def purge_cache(cache: dict) -> dict:
    """
    Remove entries from the cache whose files no longer exist.

    Args:
        cache: The cache dictionary, expected to have a 'files' key.

    Returns:
        The cleaned cache dictionary.
    """
    if "files" in cache:
        for file_path_str in list(cache["files"].keys()):
            file_path = Path(file_path_str)
            if not file_path.exists():
                cache["files"].pop(file_path_str, None)
                logger.debug(f"Removed non-existing file from cache: {json.dumps(str(file_path.as_posix()))}")

    return cache


def main(config) -> None:
    # logger.debug(f"Config: {json.dumps(config, indent=4)}")

    cache_file = Path(str(config.get("cache_file")))
    if cache_file is None:
        logger.error("No cache file specified in config.")
        raise ValueError
    # logger.debug(f"Cache file: {json.dumps(str(cache_file))}")

    media_root = Path(str(config.get("media_dir")))
    if media_root is None:
        logger.error("No media root specified in config.")
        raise ValueError
    # logger.debug(f"Media root: {json.dumps(str(media_root))}")

    media_extensions = list(config.get("media_extensions"))
    if media_extensions is None:
        logger.error("No media extensions specified in config.")
        raise ValueError
    # logger.debug(f"Media extensions: {json.dumps(media_extensions)}")

    cache = load_cache(cache_file)
    # logger.debug(f"Cache: {json.dumps(cache, indent=4)}")

    cache = build_files_cache(cache, cache_file, media_root, media_extensions)
    cache = build_hashes_cache(cache, cache_file)
    # logger.debug(f"Cache: {json.dumps(cache, indent=4)}")


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
    Set up logging for a script safely.

    Immediate flush for Tkinter callbacks is handled automatically.
    """
    file_path = Path(file_path)
    dir_path = file_path.parent
    dir_path.mkdir(parents=True, exist_ok=True)

    logger_obj.handlers.clear()
    logger_obj.setLevel(min(console_logging_level, file_logging_level))

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
