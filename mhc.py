"""
{Script Name}

{Summary of what the script does}

{How to use the script}
"""

import cv2
import datetime
import imagehash
import json
import logging
import logging.handlers
import os
import socket
import sys
import tempfile
import time
import typing
from dataclasses import dataclass, field
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path
from PIL import Image
from send2trash import send2trash


logger = logging.getLogger(__name__)

__version__ = "1.2.2"  # Major.Minor.Patch

log_buffer = logging.handlers.MemoryHandler(
    capacity=0,
    flushLevel=logging.CRITICAL,
    target=None,
)

logger.addHandler(log_buffer)
logger.setLevel(logging.DEBUG)


class FilterMode:
    OLDEST = "by_oldest"
    NEWEST = "by_newest"
    FOLDER_PRIORITY = "by_folder_priority"
    NONE = None  # Scanned order


@dataclass
class ScriptSettings:
    cache_file_path = Path(r"cache.json")
    media_dir_path = Path(r"media")
    remove_deleted_files_from_cache = False
    filter_mode = FilterMode.FOLDER_PRIORITY
    folder_priority = []

    save_cache_frequency = 100  # files scanned
    ignore_files_containing = ["$", "System"]


@dataclass
class LogSettings:
    mode: typing.Literal["per_run", "latest", "per_day", "single_file", "console_only"] = "per_run"
    folder: Path = Path("Logs")
    console_level: int = logging.DEBUG
    file_level: int = logging.DEBUG
    date_format: str = "%Y-%m-%dT%H:%M:%S"
    message_format: str = "%(asctime)s.%(msecs)03d [%(levelname)-8s] %(message)s"
    # message_format: str = "%(asctime)s.%(msecs)03d [%(levelname)-8s] %(module)s:%(funcName)s - %(message)s"
    max_files: int | None = 15
    open_log_after_run: bool = False


@dataclass
class RuntimeSettings:
    pause_on_error: bool = True
    always_pause: bool = False


@dataclass
class Config:
    script_settings: ScriptSettings = field(default_factory=ScriptSettings)
    log_settings: LogSettings = field(default_factory=LogSettings)
    runtime_settings: RuntimeSettings = field(default_factory=RuntimeSettings)


class HashGenerationError(Exception):
    """Custom exception for media hashing failures."""
    pass


def generate_fast_hash(file_path: Path, image_exts: set, video_exts: set):
    ext = file_path.suffix.lower()

    if ext in image_exts:
        return get_fast_image_hash(file_path)
    if ext in video_exts:
        return get_fast_video_hash(file_path)

    raise HashGenerationError(f"Unsupported file extension: {ext}")


def get_fast_image_hash(file_path: Path):
    try:
        with Image.open(file_path) as img:
            # Standardize to RGB first so getextrema() consistently
            # returns a 3-element tuple of (min, max) pairs.
            rgb_img = img.convert("RGB")
            extrema: tuple[tuple[int, int], tuple[int, int], tuple[int, int]] = rgb_img.getextrema()  # type: ignore

            # If min == max for all channels, the image is a solid uniform color
            if all(band_min == band_max for band_min, band_max in extrema):
                r = extrema[0][0]
                g = extrema[1][0]
                b = extrema[2][0]
                return f"solid_{r:02x}{g:02x}{b:02x}"

            # Helper to perform the hash for non-solid images
            def compute(image_obj):
                temp = image_obj.convert("L").resize((64, 64), Image.Resampling.BILINEAR)
                return imagehash.phash(temp, hash_size=16)

            h = compute(img)

            if str(h) == ('0' * 64) and img.mode in ("RGBA", "LA", "P"):
                bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
                bg.paste(img.convert("RGBA"), mask=img.convert("RGBA").getchannel('A'))
                h = compute(bg)

            return str(h)
    except Exception as e:
        logger.error("Failed to hash %s: %s", file_path, e)
        return None


def get_fast_video_hash(file_path: Path):
    cap = cv2.VideoCapture(str(file_path.resolve()))

    if not cap.isOpened():
        logger.error("OpenCV could not open %s", file_path)
        raise HashGenerationError("OpenCV open failure")

    try:
        success, frame = cap.read()
        if not success or frame is None:
            logger.error("Could not read frame from %s", file_path)
            raise HashGenerationError("Frame read failure")

        small_frame = cv2.resize(frame, (32, 32), interpolation=cv2.INTER_NEAREST)
        gray_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2GRAY)
        img = Image.fromarray(gray_frame)
        return str(imagehash.phash(img))
    except Exception as e:
        logger.error("Video hash failed for %s: %s", file_path, e)
        raise
    finally:
        cap.release()


def iter_files(root_dir: Path, valid_exts: set, ignore_list: list):
    """
    Scans root_dir recursively and yields absolute Path objects.

    - root_dir: The starting Path object.
    - valid_exts: A set of extensions including the dot (e.g., {'.txt', '.jpg'}).
    - ignore_list: A list of strings; if found in the path, the file is skipped.
    """
    root_dir = Path(root_dir).resolve()

    for path in root_dir.rglob("*"):
        if not path.is_file():
            continue

        if path.suffix.lower() not in valid_exts:
            continue

        resolved_path = path.resolve()
        path_str = resolved_path.as_posix()

        if any(ignore_str in path_str for ignore_str in ignore_list):
            continue

        yield resolved_path


def build_files_cache(cache: dict, config: Config) -> dict:
    logger.debug("Building files cache...")

    # 1. Configuration & Setup
    cache_path = config.script_settings.cache_file_path
    save_cache_frequency = config.script_settings.save_cache_frequency
    remove_deleted = config.script_settings.remove_deleted_files_from_cache

    root_dir = config.script_settings.media_dir_path
    ignore_list = config.script_settings.ignore_files_containing

    image_exts = {'.jpg', '.jpeg', '.png', '.bmp', '.webp', '.tiff'}
    video_exts = {'.mp4', '.mkv', '.mov', '.avi', '.gif', '.wmv'}
    valid_exts = image_exts | video_exts

    # Ensure cache structure exists
    if "files" not in cache:
        cache["files"] = {}

    cache_updates = 0
    seen_paths = set()
    files_scanned = 0
    last_status_message_time = time.time()

    # 2. Scan and Update Loop
    for file_path in iter_files(root_dir, valid_exts, ignore_list):
        files_scanned += 1

        # Periodic status logging
        if time.time() - last_status_message_time >= 1.5:
            logger.debug("Scanned %s files...", files_scanned)
            last_status_message_time = time.time()

        try:
            # Consistent path string for cache keys
            path_str = file_path.as_posix()
            seen_paths.add(path_str)

            file_stats = file_path.stat()
            size, mtime = file_stats.st_size, file_stats.st_mtime

            # Check if cache is already up to date
            if path_str in cache["files"]:
                entry = cache["files"][path_str]
                if entry.get("size") == size and entry.get("mtime") == mtime:
                    continue

            # Generate hash for new or changed files
            file_hash = generate_fast_hash(file_path, image_exts=image_exts, video_exts=video_exts)
            if file_hash:
                cache["files"][path_str] = {
                    "hash": file_hash,
                    "size": size,
                    "mtime": mtime,
                }
                logger.debug("Hashed: %s", path_str)
                cache_updates += 1

        except Exception as e:
            logger.error("Failed to process %s: %s", file_path, e)
            continue

        # Incremental save
        if cache_updates >= save_cache_frequency > 0:
            write_json_file(cache_path, cache)
            cache_updates = 0

    if remove_deleted:
        cached_keys = list(cache["files"].keys())
        for path_key in cached_keys:
            if not Path(path_key).exists():
                logger.debug("Removing stale entry from cache: %s", path_key)
                del cache["files"][path_key]
                cache_updates += 1

    # Final save if there are pending changes
    if cache_updates > 0:
        write_json_file(cache_path, cache)

    return cache


def build_hashes_cache(cache: dict, config: Config) -> dict:
    logger.debug("Building hashes cache...")

    cache["hashes"] = {}
    cache_path = config.script_settings.cache_file_path

    for file_path, file_data in cache.get("files", {}).items():
        if not Path(file_path).exists():
            continue

        file_hash = file_data["hash"]

        # Initialize the hash bucket if it doesn't exist
        cache["hashes"].setdefault(file_hash, {})

        # Assign the file data to the path key under that hash
        cache["hashes"][file_hash][file_path] = {
            "hash": file_hash,
            "size": file_data["size"],
            "mtime": file_data["mtime"]
        }
    write_json_file(cache_path, cache)
    return cache


def sort_hashes_cache(cache: dict, config: Config) -> dict:
    logger.debug("Sorting hashes using strategy: %s...", config.script_settings.filter_mode)

    cache_path = config.script_settings.cache_file_path
    filter_mode = config.script_settings.filter_mode
    # Ensure paths are standardized for comparison
    folder_priority = [str(Path(p).as_posix()) for p in config.script_settings.folder_priority]

    def get_folder_score(path_str: str) -> int:
        path_posix = str(Path(path_str).as_posix())
        for i, priority_path in enumerate(folder_priority):
            if priority_path in path_posix:
                return i
        return len(folder_priority)

    hashes_data = cache.get("hashes", {})

    match filter_mode:
        case FilterMode.OLDEST:
            for file_hash, file_dicts in hashes_data.items():
                items = sorted(file_dicts.items(), key=lambda x: x[1]['mtime'])
                hashes_data[file_hash] = dict(items)

        case FilterMode.NEWEST:
            for file_hash, file_dicts in hashes_data.items():
                items = sorted(file_dicts.items(), key=lambda x: x[1]['mtime'], reverse=True)
                hashes_data[file_hash] = dict(items)

        case FilterMode.FOLDER_PRIORITY:
            # --- Validation Step ---
            all_media_folders = set()
            for file_dicts in hashes_data.values():
                for file_path in file_dicts.keys():
                    all_media_folders.add(str(Path(file_path).parent.as_posix()))

            missing_paths = sorted([p for p in all_media_folders if p not in folder_priority])

            if missing_paths:
                raise ValueError(f"Config folder_priority is missing folders:\n{json.dumps(missing_paths, indent=4)}")

            # --- Sorting Step ---
            for file_hash, file_group in hashes_data.items():
                # Corrected: Indented inside the loop and removed reverse=True
                items = sorted(file_group.items(), key=lambda x: get_folder_score(x[0]))
                hashes_data[file_hash] = dict(items)

        case FilterMode.NONE | None:
            pass

        case _:
            raise ValueError(f"Unknown filter mode: {json.dumps(str(filter_mode))}")

    write_json_file(cache_path, cache)
    return cache


def build_duplicates_list(cache: dict, config: Config) -> dict:
    logger.debug("Building duplicates cache...")
    cache_path = config.script_settings.cache_file_path

    cache["duplicates"] = []
    for hash, images_data in cache["hashes"].items():
        if len(images_data) <= 1:
            continue

        image_paths = list(images_data.keys())
        # Log the file we are keeping
        logger.info("Keeping %s", image_paths[0])
        # Loop through the rest, log them individually, and cache them
        for image_path in image_paths[1:]:
            logger.info("    - Discarding %s", image_path)
            cache["duplicates"].append(image_path)

    logger.debug("Found %s duplicates.", len(cache["duplicates"]))
    write_json_file(cache_path, cache)
    return cache


def remove_duplicates(cache):
    logger.debug("Removing duplicates...")
    for image_path in cache["duplicates"]:
        image_path = Path(image_path)
        send2trash(image_path)
        logger.info("Sent %s to recycle bin", image_path)


def main(config: Config):
    cache_path = config.script_settings.cache_file_path

    try:
        cache = read_json_file(cache_path)
    except FileNotFoundError:
        logger.debug("Initialized new blank cache...")
        cache = {}
    if not isinstance(cache, dict):
        raise TypeError("Cache should be formatted as a dictionary.")
    cache.setdefault("files", {})
    cache.setdefault("hashes", {})

    cache = build_files_cache(cache=cache, config=config)
    cache = build_hashes_cache(cache=cache, config=config)
    cache = sort_hashes_cache(cache=cache, config=config)
    cache = build_duplicates_list(cache=cache, config=config)
    remove_duplicates(cache=cache)


def read_json_file(file_path: Path) -> dict | list | None:
    """
    Safely reads and parses a JSON file.
    """
    if not file_path.exists():
        logger.warning("File not found: %s", file_path)
        raise FileNotFoundError("File not found")

    try:
        data = json.loads(file_path.read_text(encoding='utf-8'))
        logger.info("Successfully read data from %s", file_path)
        return data

    except json.JSONDecodeError as e:
        logger.error("Invalid JSON format in %s: %s", file_path, e)
        return None

    except Exception as e:
        logger.error("Unexpected error reading %s: %s", file_path, e)
        return None


def write_json_file(file_path: Path, data: dict | list) -> bool:
    """
    Writes data to a JSON file atomically.
    """
    file_path = Path(file_path).absolute()

    if not file_path.parent.exists():
        file_path.parent.mkdir(parents=True, exist_ok=True)
        logger.debug("Created %s", file_path.parent)

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
        logger.info("Successfully saved to %s", file_path)
        return True

    except (KeyboardInterrupt, SystemExit):
        logger.error("Write interrupted for %s. Cleaning up.", file_path)
        if temp_file_path and temp_file_path.exists():
            temp_file_path.unlink()
        raise

    except Exception as e:
        logger.error("Failed to write to %s: %s", file_path, e)
        if temp_file_path and temp_file_path.exists():
            temp_file_path.unlink()
        return False


def enforce_max_log_count(dir_path: Path, max_count: int, script_name: str) -> None:
    """
    Enforce a maximum number of log files for this script.

    Rules:
    - Only affects files ending with `.log`
    - Only affects logs that contain the script name
    - Sorting is performed lexicographically by filename
    """
    if max_count <= 0:
        return

    if not dir_path.exists():
        return

    log_files = [f for f in dir_path.glob("*.log") if script_name in f.name]
    if len(log_files) <= max_count:
        return
    log_files.sort(key=lambda p: p.name)
    to_delete = log_files[:-max_count]
    for file in to_delete:
        try:
            file.unlink()
            logger.debug("Removed old log %s", file)
        except OSError as e:
            logger.debug("Failed removing old log %s: %s", file, e)


def build_log_path(log_settings: LogSettings) -> Path | None:
    """
    Builds the final log file path based on logging mode.
    """
    if log_settings.mode == "console_only":
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    day_stamp = datetime.now().strftime("%Y%m%d")

    script_name = Path(__file__).stem
    pc_name = socket.gethostname()

    log_dir = Path(log_settings.folder).expanduser().resolve()

    match log_settings.mode:
        case "per_run":
            filename = f"{timestamp}__{script_name}__{pc_name}.log"
        case "latest":
            filename = f"latest_{script_name}__{pc_name}.log"
        case "per_day":
            filename = f"{day_stamp}__{script_name}__{pc_name}.log"
        case "single_file":
            filename = f"{script_name}__{pc_name}.log"
        case _:
            filename = f"{timestamp}__{script_name}__{pc_name}.log"

    return log_dir / filename


class JsonArgsFilter(logging.Filter):
    """
    Automatically formats log arguments:
    - Keeps numeric types intact (so %d / %.6f still work)
    - Applies JSON-style formatting to Path and str (adds quotes)
    - Safely serializes other objects
    """

    def filter(self, record: logging.LogRecord) -> bool:
        if not record.args:
            return True

        raw_args = list(record.args) if isinstance(record.args, tuple) else [record.args]
        processed_args = []

        for val in raw_args:
            match val:
                case Path():
                    processed_args.append(json.dumps(val.as_posix()))

                case str():
                    processed_args.append(json.dumps(val))

                case int() | float() | bool():
                    processed_args.append(val)

                case None:
                    processed_args.append(val)

                case _:
                    processed_args.append(json.dumps(val, default=str))

        record.args = tuple(processed_args)
        return True


def setup_logging(logger_obj: logging.Logger, log_settings: LogSettings) -> Path | None:
    """
    Set up console and file logging.
    """
    logger_obj.handlers.clear()
    logger_obj.setLevel(logging.DEBUG)
    logger_obj.propagate = False

    # Attach the automatic JSON formatting filter
    logger_obj.addFilter(JsonArgsFilter())

    log_path = build_log_path(log_settings)

    formatter = logging.Formatter(
        log_settings.message_format,
        datefmt=log_settings.date_format,
    )

    if log_path:
        try:
            log_path.parent.mkdir(parents=True, exist_ok=True)

        except OSError as e:
            raise RuntimeError(f"Failed creating log directory {log_path.parent}") from e

        file_handler: logging.Handler

        match log_settings.mode:
            case "per_day":
                file_handler = TimedRotatingFileHandler(filename=log_path, when="midnight", interval=1, backupCount=log_settings.max_files or 0, encoding="utf-8")
            case "single_file":
                file_handler = logging.FileHandler(log_path, mode="a", encoding="utf-8")
            case _:
                file_handler = logging.FileHandler(log_path, mode="w", encoding="utf-8")

        file_handler.setLevel(log_settings.file_level)
        file_handler.setFormatter(formatter)
        logger_obj.addHandler(file_handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(log_settings.console_level)
    console_handler.setFormatter(formatter)

    logger_obj.addHandler(console_handler)

    write_banner(logger_obj)

    if log_buffer:
        class _ForwardToLogger(logging.Handler):
            def emit(self, record):
                logger_obj.handle(record)

        forward_handler = _ForwardToLogger()
        log_buffer.setTarget(forward_handler)
        log_buffer.flush()
        log_buffer.close()

    if (log_settings.max_files and log_path and log_settings.mode not in ("per_day", "console_only")):
        enforce_max_log_count(dir_path=log_path.parent, max_count=log_settings.max_files, script_name=Path(__file__).stem)

    return log_path


def write_banner(logger_obj: logging.Logger):
    """
    Writes a clean session banner without log prefixes.
    """
    separator = "-" * 80

    banner = (
        f"{separator}\n"
        f"SCRIPT     | {json.dumps(Path(__file__).resolve().as_posix())}\n"
        f"VERSION    | {__version__}\n"
        f"START TIME | {datetime.now().isoformat(timespec='milliseconds')}\n"
        f"USER       | {os.getlogin()}\n"
        f"HOST       | {socket.gethostname()}\n"
        f"RUNTIME    | Python {sys.version.split()[0]}\n"
        f"{separator}"
    )

    original_formatters = {}

    class RawFormatter(logging.Formatter):
        """
        Formatter that outputs only the log message with no prefixes.
        """

        def format(self, record):
            return record.getMessage()

    try:
        for handler in logger_obj.handlers:
            original_formatters[handler] = handler.formatter
            handler.setFormatter(RawFormatter())

        logger_obj.info(banner)

    finally:
        for handler, formatter in original_formatters.items():
            handler.setFormatter(formatter)


def bootstrap():
    exit_code = 0
    log_path: Path | None = None
    config = Config()

    try:
        log_path = setup_logging(logger_obj=logger, log_settings=config.log_settings)
        main(config)

    except KeyboardInterrupt:
        logger.warning("Operation interrupted by user.")
        exit_code = 130

    except Exception as e:
        logger.exception("A fatal error has occurred: %s", e)
        exit_code = 1

    if (config.log_settings.open_log_after_run and log_path and log_path.exists()):
        try:
            match sys.platform:
                case plat if plat.startswith("win"):
                    os.startfile(log_path)
                case "darwin":
                    os.system(f'open "{log_path}"')
                case _:
                    os.system(f'xdg-open "{log_path}"')

        except Exception as e:
            logger.warning("Failed to open log file: %s", e)

    if (config.runtime_settings.always_pause or (config.runtime_settings.pause_on_error and exit_code != 0)):
        input("Press Enter to exit...")

    return exit_code


if __name__ == "__main__":
    sys.exit(bootstrap())
