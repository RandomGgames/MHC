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


__version__ = "1.0.0"  # Major.Minor.Patch


def read_toml(file_path: typing.Union[str, pathlib.Path]) -> dict:
    """
    Read configuration settings from the TOML file.
    """
    file_path = pathlib.Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    config = toml.load(file_path)
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
        cache_path = pathlib.Path(config["cache"])
        cache_dir = cache_path.parent
        cache_dir.mkdir(parents=True, exist_ok=True)

        temp_path = cache_path.with_suffix(".tmp")
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(cache, f, indent=4)

        # Atomic replace
        temp_path.replace(cache_path)
        logger.debug(f"Saved cache to {cache_path}.")
    except Exception as e:
        logger.error(f"Failed to save cache to {config.get('cache')} due to {e}.")
        raise


def get_media_files(config: dict) -> typing.Iterable[pathlib.Path]:
    media_dir_raw = config.get("media_dir")
    if not isinstance(media_dir_raw, (str, os.PathLike)):
        raise TypeError(f"config['media_dir'] must be a string path, got {type(media_dir_raw)}")

    media_dir = pathlib.Path(media_dir_raw).resolve()

    logger.debug(f"Searching for media files in '{media_dir}'...")

    exts = {ext.lower().lstrip(".") for ext in config.get("media_extensions", [])}
    ignore_phrases = [p.lower() for p in config.get("ignore_files_with", [])]

    if not media_dir.exists():
        logger.warning(f"Media directory does not exist: {media_dir}")
        return

    for file_path in media_dir.rglob("*"):
        if not file_path.is_file():
            continue

        suffix = file_path.suffix.lower().lstrip(".")
        if not suffix or suffix not in exts:
            continue

        if any(phrase in file_path.stem.lower() for phrase in ignore_phrases):
            continue

        logger.debug(f"Found media file: {file_path}")
        yield file_path


def get_file_data(file_path: typing.Union[str, pathlib.Path]) -> tuple[int, int, int]:
    file_path = pathlib.Path(file_path)
    st = file_path.stat()
    modified_time = getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9))
    # Prefer birthtime if available, otherwise use ctime (metadata change time)
    created_time = getattr(st, "st_birthtime_ns", getattr(st, "st_ctime_ns", int(st.st_ctime * 1e9)))
    size = st.st_size
    logger.debug(f"Got file data: {modified_time=}, {created_time=}, {size=}")
    return modified_time, created_time, size


def generate_hash(file_path: typing.Union[str, pathlib.Path], algorithm="sha256") -> str:
    file_path = pathlib.Path(file_path)
    logger.debug(f"Generating hash for {file_path}...")
    try:
        with open(file_path, "rb") as f:
            # Prefer file_digest if available (Python 3.11+)
            try:
                digest = hashlib.file_digest(f, algorithm)
                hexd = digest.hexdigest()
            except AttributeError:
                h = hashlib.new(algorithm)
                for chunk in iter(lambda: f.read(1024 * 1024), b""):
                    h.update(chunk)
                hexd = h.hexdigest()
    except Exception as e:
        # Bubble up with context so callers can decide to skip the file
        logger.exception(f"Failed to read/hash file {file_path}: {e}")
        raise
    logger.debug(f"Generated hash '{hexd}' for {file_path}")
    return hexd


def build_files_cache(cache: dict, config: dict) -> None:
    logger.debug("Building files cache...")

    changes_since_save = 0
    total_changes = 0

    for idx, file_path_obj in enumerate(get_media_files(config), start=1):
        file_path = str(file_path_obj)
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
        old_path = pathlib.Path(old)
        new_path = pathlib.Path(new)

        if not old_path.exists():
            logger.warning(f"Skipping rename '{old}' → '{new}': source missing.")
            continue

        if old == new:
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

        logger.debug("Rename applied and cache updated.")

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


def main() -> None:
    cache = load_cache(config)
    build_files_cache(cache, config)
    build_hashes_cache(cache, config)
    delete_duplicate_files(cache, config)
    build_hashes_cache(cache, config)
    rename_files_to_hashes(cache, config)


def format_duration_long(duration_seconds: float) -> str:
    """
    Format duration in a human-friendly way, showing only the two largest non-zero units.
    For durations >= 1s, do not show microseconds or nanoseconds.
    For durations >= 1m, do not show milliseconds.
    """
    ns = int(duration_seconds * 1_000_000_000)
    units = [
        ('y', 365 * 24 * 60 * 60 * 1_000_000_000),
        ('mo', 30 * 24 * 60 * 60 * 1_000_000_000),
        ('d', 24 * 60 * 60 * 1_000_000_000),
        ('h', 60 * 60 * 1_000_000_000),
        ('m', 60 * 1_000_000_000),
        ('s', 1_000_000_000),
        ('ms', 1_000_000),
        ('us', 1_000),
        ('ns', 1),
    ]
    parts = []
    for name, factor in units:
        value, ns = divmod(ns, factor)
        if value:
            parts.append(f"{value}{name}")
        # Stop after two largest non-zero units
        if len(parts) == 2:
            break
    if not parts:
        return "0s"
    return "".join(parts)


def setup_logging(
        logger: logging.Logger,
        log_file_path: typing.Union[str, pathlib.Path],
        number_of_logs_to_keep: typing.Union[int, None] = None,
        console_logging_level: int = logging.DEBUG,
        file_logging_level: int = logging.DEBUG,
        log_message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s] [%(name)s]: %(message)s",
        date_format: str = "%Y-%m-%d %H:%M:%S") -> None:
    log_file_path = pathlib.Path(log_file_path)
    log_dir = log_file_path.parent
    log_dir.mkdir(parents=True, exist_ok=True)

    # Limit # of logs in logs folder
    if number_of_logs_to_keep is not None:
        log_files = sorted([f for f in log_dir.glob("*.log")], key=lambda f: f.stat().st_mtime)
        if len(log_files) > number_of_logs_to_keep:
            for file in log_files[:-number_of_logs_to_keep]:
                file.unlink()

    # Clear old handlers to avoid duplication
    logger.handlers.clear()
    logger.setLevel(file_logging_level)

    formatter = logging.Formatter(log_message_format, datefmt=date_format)

    # File Handler
    file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler.setLevel(file_logging_level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_logging_level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)


if __name__ == "__main__":
    config_path = pathlib.Path("config.toml")
    if not config_path.exists():
        raise FileNotFoundError(f"Missing {config_path}")
    global config
    config = read_toml(config_path)

    console_logging_level = getattr(logging, config.get("logging", {}).get("console_logging_level", "INFO").upper(), logging.DEBUG)
    file_logging_level = getattr(logging, config.get("logging", {}).get("file_logging_level", "INFO").upper(), logging.DEBUG)
    logs_file_path = config.get("logging", {}).get("logs_file_path", "logs")
    use_logs_folder = config.get("logging", {}).get("use_logs_folder", True)
    number_of_logs_to_keep = config.get("logging", {}).get("number_of_logs_to_keep", 10)
    log_message_format = config.get("logging", {}).get(
        "log_message_format",
        "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s]: %(message)s"
    )

    script_name = pathlib.Path(__file__).stem
    pc_name = socket.gethostname()
    if use_logs_folder:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        log_dir = pathlib.Path(f"{logs_file_path}/{script_name}")
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file_name = f"{timestamp}_{script_name}_{pc_name}.log"
        log_file_path = log_dir / log_file_name
    else:
        log_file_path = pathlib.Path(f"{script_name}_{pc_name}.log")

    setup_logging(
        logger,
        log_file_path,
        console_logging_level=console_logging_level,
        file_logging_level=file_logging_level,
        number_of_logs_to_keep=number_of_logs_to_keep,
        log_message_format=log_message_format
    )

    error = 0
    try:
        start_time = time.perf_counter_ns()
        logger.info(f"Script: {script_name} | Version: {__version__} | Host: {pc_name}")

        main()
        end_time = time.perf_counter_ns()
        duration = end_time - start_time
        duration = format_duration_long(duration / 1e9)
        logger.info(f"Execution completed in {duration}.")
    except KeyboardInterrupt:
        logger.warning("Operation interrupted by user.")
        error = 130
    except Exception as e:
        logger.warning(f"A fatal error has occurred: {repr(e)}\n{traceback.format_exc()}")
        error = 1
    finally:
        for handler in logger.handlers:
            handler.close()
        logger.handlers.clear()
        sys.exit(error)
