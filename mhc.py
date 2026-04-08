"""
MHC: Media Hasher and Cacher

A python script which hashes images and videos, then based
on a keep priority, deletes duplicates.

{Summary of what the script does}

{How to use the script}
"""

import cv2
import hashlib
import imagehash
import io
import json
import logging
import logging.handlers
import mimetypes
import numpy as np
import os
import platform
import re
import send2trash
import socket
import sys
import tempfile
import time
import tomllib
import typing
import win32com.client
from dataclasses import dataclass, field, fields, asdict
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path
from PIL import Image
from pillow_heif import register_heif_opener
from typing import Iterable, Pattern

logger = logging.getLogger(__name__)

__version__ = "2.0.0"  # Major.Minor.Patch


logger = logging.getLogger(__name__)
log_buffer = logging.handlers.MemoryHandler(
    capacity=0,
    flushLevel=logging.CRITICAL,
    target=None,
)
logger.addHandler(log_buffer)
logger.setLevel(logging.DEBUG)

register_heif_opener()


@dataclass
class ScriptSettings:
    IMAGE_EXTENSIONS = {
        ".jpg",
        ".jpeg",
        ".png",
        ".bmp",
        ".webp",
        ".tiff",
        ".heic",
        ".heif"
    }
    VIDEO_EXTENSIONS = {
        ".mp4",
        ".avi",
        ".mov",
        ".mkv",
        ".webm"
    }


@dataclass
class LogSettings:
    mode: typing.Literal["per_run", "latest", "per_day", "single_file", "ConsoleOnly"] = "per_day"
    folder: Path = Path(r"Logs")
    console_level: int = logging.DEBUG
    file_level: int = logging.DEBUG
    date_format: str = "%Y-%m-%d %H:%M:%S"
    message_format: str = "%(asctime)s.%(msecs)03d %(levelname)s [%(funcName)s] - %(message)s"
    max_files: int | None = 10
    open_log_after_run: bool = False


@dataclass
class InternalSettings:
    use_config_file: bool = False


@dataclass
class RuntimeSettings:
    pause_on_error: bool = True
    always_pause: bool = False


@dataclass
class Config:
    script: ScriptSettings = field(default_factory=ScriptSettings)
    logs: LogSettings = field(default_factory=LogSettings)
    runtime: RuntimeSettings = field(default_factory=RuntimeSettings)


def hash_image_file(path, hash_size=16):
    """
    Compute a perceptual hash for a single image file.
    Supports HEIC and standard image formats.

    Returns:
        np.ndarray: binary hash array or None if failed
    """
    ext = os.path.splitext(path)[1].lower()

    try:
        img = Image.open(path).convert("L")
        h = imagehash.phash(img, hash_size=hash_size)
        return np.array(h.hash).astype(np.uint8)
    except Exception as e:
        print(f"Error hashing image {path}: {e}")
        return None


def hash_video_file(path, hash_size=16, num_frames=8):
    """
    Compute a perceptual hash for a video file by sampling frames.

    Returns:
        np.ndarray: median hash array or None if failed
    """
    ext = os.path.splitext(path)[1].lower()

    cap = cv2.VideoCapture(path)
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))
    if total_frames <= 0:
        cap.release()
        return None

    frame_idxs = np.linspace(0, total_frames - 1, num_frames, dtype=int)
    hashes = []

    for idx in frame_idxs:
        cap.set(cv2.CAP_PROP_POS_FRAMES, idx)
        ret, frame = cap.read()
        if not ret:
            continue

        # Convert to PIL image
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(frame).convert("L")
        h = imagehash.phash(pil_img, hash_size=hash_size)
        hashes.append(np.array(h.hash).astype(np.uint8))

    cap.release()
    if not hashes:
        return None

    # Median aggregation across frames
    return np.median(hashes, axis=0).astype(np.uint8)


def hash_distance(h1, h2):
    """
    Compute Hamming distance between two hashes (images or videos)
    """
    return np.sum(h1 != h2)


def is_similar(h1, h2, threshold=4):
    """
    Check if two hashes are similar based on a threshold.
    """
    return hash_distance(h1, h2) <= threshold


def main():
    """Code goes here"""


def read_json_file(file_path: Path) -> dict | list | None:
    """
    Safely reads and parses a JSON file.
    """
    if not file_path.exists():
        logger.warning("File not found: %s", json.dumps(str(file_path)))
        raise FileNotFoundError("File not found")

    try:
        data = json.loads(file_path.read_text(encoding='utf-8'))
        logger.info("Successfully read data from %s", json.dumps(str(file_path)))
        return data

    except json.JSONDecodeError as e:
        logger.error("Invalid JSON format in %s: %s", json.dumps(str(file_path)), e)
        return None

    except Exception as e:
        logger.error("Unexpected error reading %s: %s", json.dumps(str(file_path)), e)
        return None


def write_json_file(file_path: Path, data: dict | list) -> bool:
    """
    Writes data to a JSON file atomically.
    """
    file_path = Path(file_path).absolute()

    if not file_path.parent.exists():
        file_path.parent.mkdir(parents=True, exist_ok=True)
        logger.debug("Created %s", json.dumps(str(file_path.parent.as_posix())))

    temp_file_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', dir=str(file_path.parent), encoding='utf-8', suffix=".tmp", delete=False) as tf:
            # Get file path from tempfile object
            temp_file_path = Path(tf.name)
            json.dump(data, tf, indent=4)
            tf.flush()
            os.fsync(tf.fileno())

        # Atomic swap
        temp_file_path.replace(file_path)
        logger.info("Successfully saved to %s", json.dumps(str(file_path)))
        return True

    except (KeyboardInterrupt, SystemExit):
        logger.error("Write interrupted for %s. Cleaning up.", json.dumps(str(file_path)))
        if temp_file_path and temp_file_path.exists():
            temp_file_path.unlink()
        raise

    except Exception as e:
        logger.error("Failed to write to %s: %s", json.dumps(str(file_path)), e)
        if temp_file_path and temp_file_path.exists():
            temp_file_path.unlink()
        return False


def load_config(file_path: Path) -> Config:
    config = Config()
    needs_sync = False

    try:
        external_config = read_json_file(file_path)
        if not isinstance(external_config, dict):
            external_config = {}
            needs_sync = True
    except FileNotFoundError:
        external_config = {}
        needs_sync = True
    except Exception:
        raise

    # Merge logic
    for section in fields(config):
        section_name = section.name
        if section_name not in external_config:
            needs_sync = True
            continue

        section_instance = getattr(config, section_name)
        json_values = external_config[section_name]

        for f in fields(section_instance):
            if f.name in json_values:
                val = json_values[f.name]
                if f.type is Path and isinstance(val, str):
                    val = Path(val)
                setattr(section_instance, f.name, val)
            else:
                needs_sync = True

    # Check for keys in external config that aren't in internal config
    internal_field_names = {f.name for f in fields(config)}
    if any(k for k in external_config if k not in internal_field_names):
        needs_sync = True

    if needs_sync:
        def path_serializer(obj):
            if isinstance(obj, Path):
                return str(obj)
            raise TypeError(f"Type {type(obj)} not serializable")

        # We re-serialize the internal_config (which now has merged data)
        # This naturally prunes extra keys because they weren't in the dataclass!
        synced_config = json.loads(json.dumps(asdict(config), default=path_serializer))
        write_json_file(file_path, synced_config)

    return config


def enforce_max_log_count(dir_path: Path, max_count: int, script_name: str) -> None:
    """
    Enforce a maximum number of log files for this script.
    Deletes the oldest logs based on filename ordering.

    Rules:
    - Only affects files ending with `.log`
    - Only affects logs that contain the script_name
    - Sorting is done by filename (lexicographically)
    """
    if max_count <= 0:
        return
    if not dir_path.exists():
        return
    log_files = [
        f for f in dir_path.glob("*.log")
        if script_name in f.name
    ]
    if len(log_files) <= max_count:
        return
    log_files.sort(key=lambda p: p.name)
    to_delete = log_files[:-max_count]
    for file in to_delete:
        try:
            file.unlink()
            logger.debug("Removed %s", json.dumps(file.absolute().as_posix()))
        except Exception:
            # Avoid raising during bootstrap
            pass


def setup_logging(logger_obj: logging.Logger, log_settings: LogSettings) -> Path | None:
    """Set up file and console logging with flexible modes and rotation."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    day_stamp = datetime.now().strftime("%Y%m%d")
    script_name = Path(__file__).stem
    pc_name = socket.gethostname()

    log_path: Path | None = None

    if log_settings.mode != "ConsoleOnly":
        log_dir = (log_settings.folder if isinstance(log_settings.folder, Path) else Path(log_settings.folder))
        log_dir = log_dir.expanduser().resolve()
        if not log_dir.exists():
            log_dir.mkdir(parents=True, exist_ok=True)
            logger.debug("Created log folder: %s", log_dir.as_posix())

        match log_settings.mode:
            case "per_run":
                log_path = log_dir / f"{timestamp}__{script_name}__{pc_name}.log"
            case "latest":
                log_path = log_dir / f"latest_{script_name}__{pc_name}.log"
            case "per_day":
                log_path = log_dir / f"{day_stamp}__{script_name}__{pc_name}.log"
            case "single_file":
                log_path = log_dir / f"{script_name}__{pc_name}.log"
            case _:
                log_path = log_dir / f"{timestamp}__{script_name}__{pc_name}.log"

    logger_obj.handlers.clear()
    logger_obj.setLevel(logging.DEBUG)

    # Formatter
    formatter = logging.Formatter(
        log_settings.message_format,
        datefmt=log_settings.date_format,
    )

    # File handler
    file_handler: logging.Handler | None = None
    if log_path:
        match log_settings.mode:
            case "per_day":
                file_handler = TimedRotatingFileHandler(
                    filename=log_path,
                    when="midnight",
                    interval=1,
                    backupCount=log_settings.max_files or 0,
                    encoding="utf-8",
                )
            case "single_file" | "latest" | "per_run":
                file_mode = "a" if log_settings.mode == "single_file" else "w"
                file_handler = logging.FileHandler(
                    log_path,
                    mode=file_mode,
                    encoding="utf-8",
                )
    if file_handler:
        file_handler.setLevel(log_settings.file_level)
        file_handler.setFormatter(formatter)
        logger_obj.addHandler(file_handler)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_settings.console_level)
    console_handler.setFormatter(formatter)
    logger_obj.addHandler(console_handler)

    # Flush logs buffer from prior to logging initialization
    if "log_buffer" in globals():
        class _ForwardToLogger(logging.Handler):
            def emit(self, record):
                logger_obj.handle(record)

        forward_handler = _ForwardToLogger()
        log_buffer.setTarget(forward_handler)
        log_buffer.flush()
        log_buffer.close()

    # Enforce max log count (except per_day which rotates automatically)
    if log_settings.max_files and log_path and log_settings.mode not in ("per_day", "ConsoleOnly"):
        try:
            enforce_max_log_count(
                dir_path=log_path.parent,
                max_count=log_settings.max_files,
                script_name=script_name,
            )
        except Exception as e:
            logger_obj.debug("Log pruning skipped: %s", e)

    return log_path


def bootstrap():
    exit_code = 0
    log_path = None
    script_path = Path(__file__)

    logger.info("=" * 80)

    config = Config()
    config_path = script_path.with_name(f"{script_path.stem}_config.json")
    global_settings = InternalSettings()
    if global_settings.use_config_file:
        config = load_config(config_path)

    try:
        log_path = setup_logging(logger_obj=logger, log_settings=config.logs)
        logger.info("%-10s %s", "Version:", __version__)
        logger.info("%-10s %s on %s", "User/Host:", os.getlogin(), socket.gethostname())
        logger.info("%-10s %s %s (v%s)", "Platform:", platform.system(), platform.release(), platform.version())
        logger.info("%-10s Python %s", "Runtime:", sys.version.split()[0])
        logger.info("%-10s %s", "Directory:", Path.cwd().as_posix())
        logger.info("%-10s %s", "AppConfig:", config)

        main()

    except KeyboardInterrupt:
        logger.warning("Operation interrupted by user.")
        exit_code = 130
    except Exception as e:
        logger.exception("A fatal error has occurred: %s", e)
        exit_code = 1
    finally:
        for handler in logger.handlers[:]:
            handler.close()
            logger.removeHandler(handler)

    if config.logs.open_log_after_run and log_path and log_path.exists():
        try:
            match sys.platform:
                case plat if plat.startswith("win"):  # Windows
                    os.startfile(log_path)
                case "darwin":  # macOS
                    os.system(f'open "{log_path}"')
                case _:  # Linux / others
                    os.system(f'xdg-open "{log_path}"')
        except Exception as e:
            logger.warning("Failed to open log file: %s", e)

    if config.runtime.always_pause or (config.runtime.pause_on_error and exit_code != 0):
        input("Press Enter to exit...")

    return exit_code


if __name__ == "__main__":
    sys.exit(bootstrap())
