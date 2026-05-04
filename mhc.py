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
import platform
import socket
import sys
import tempfile
import time
import typing
from dataclasses import dataclass, field, fields, asdict
from datetime import datetime
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path
from PIL import Image


logger = logging.getLogger(__name__)

__version__ = "1.2.0"  # Major.Minor.Patch

logger = logging.getLogger(__name__)
log_buffer = logging.handlers.MemoryHandler(
    capacity=0,
    flushLevel=logging.CRITICAL,
    target=None,
)
logger.addHandler(log_buffer)
logger.setLevel(logging.DEBUG)

cache = {}


class FilterMode:
    OLDEST = "by_oldest"
    NEWEST = "by_newest"
    FOLDER_PRIORITY = "by_folder_priority"
    NONE = None  # Scanned order


@dataclass
class ScriptSettings:
    media_dir_path = Path(r"media")
    filter_mode = FilterMode.FOLDER_PRIORITY
    folder_priority = [
        Path(r"")
    ]
    remove_deleted_files_from_cache = True
    save_cache_frequency = 100  # files scanned
    cache_file_path = Path(r"cache.json")
    ignore_files_containing = ["$", "System"]


@dataclass
class LogSettings:
    mode: typing.Literal["per_run", "latest", "per_day", "single_file", "ConsoleOnly"] = "latest"
    folder: Path = Path(r"Logs")
    console_level: int = logging.DEBUG
    file_level: int = logging.DEBUG
    date_format: str = "%Y-%m-%dT%H:%M:%S"
    message_format: str = "%(asctime)s.%(msecs)03d [%(levelname)-8s] %(module)s:%(funcName)s - %(message)s"
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
            # Helper to perform the hash
            def compute(image_obj):
                # Using BILINEAR to prevent "all zeros" from aliasing artifacts
                temp = image_obj.convert("L").resize((32, 32), Image.Resampling.BILINEAR)
                return imagehash.phash(temp)

            h = compute(img)

            # If hash is all zeros and image has transparency, retry with a white background
            if str(h) == '0000000000000000' and img.mode in ("RGBA", "LA", "P"):
                # Create white background
                bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
                # Composite original image over white
                bg.paste(img.convert("RGBA"), mask=img.convert("RGBA").getchannel('A'))
                h = compute(bg)

            return str(h)

    except Exception as e:
        logger.error("Image hash failed for %s: %s", json.dumps(str(file_path.as_posix())), e)
        raise


def get_fast_video_hash(file_path: Path):
    cap = cv2.VideoCapture(str(file_path.resolve()))

    if not cap.isOpened():
        logger.error("OpenCV could not open %s", json.dumps(str(file_path.as_posix())))
        raise HashGenerationError("OpenCV open failure")

    try:
        success, frame = cap.read()
        if not success or frame is None:
            logger.error("Could not read frame from %s", json.dumps(str(file_path.as_posix())))
            raise HashGenerationError("Frame read failure")

        small_frame = cv2.resize(frame, (32, 32), interpolation=cv2.INTER_NEAREST)
        gray_frame = cv2.cvtColor(small_frame, cv2.COLOR_BGR2GRAY)
        img = Image.fromarray(gray_frame)
        return str(imagehash.phash(img))
    except Exception as e:
        logger.error("Video hash failed for %s: %s", json.dumps(str(file_path.as_posix())), e)
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
    cache_path = config.script.cache_file_path
    save_cache_frequency = config.script.save_cache_frequency
    remove_deleted = config.script.remove_deleted_files_from_cache

    root_dir = config.script.media_dir_path
    ignore_list = config.script.ignore_files_containing

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
            if path_key not in seen_paths:
                logger.debug("Removing stale entry from cache: %s", path_key)
                del cache["files"][path_key]
                cache_updates += 1

    # Final save if there are pending changes
    if cache_updates > 0:
        write_json_file(cache_path, cache)

    return cache


def build_hashes_cache(cache: dict, config: Config) -> dict:
    cache["hashes"] = {}

    cache_path = config.script.cache_file_path

    for file_path, file_data in cache.get("files", {}).items():
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
    logger.debug("Sorting hashes using strategy: %s", config.script.filter_mode)

    cache_path = config.script.cache_file_path
    filter_mode = config.script.filter_mode
    # Ensure paths are standardized for comparison
    folder_priority = [str(Path(p).as_posix()) for p in config.script.folder_priority]

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


def main():
    config = Config()
    cache_path = config.script.cache_file_path

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


def save_config(file_path: Path, config_data: dict | list) -> bool:
    """Alias for write_json_file, specifically for configuration files."""
    return write_json_file(file_path, config_data)


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


def setup_logging(logger_obj: logging.Logger, log_settings: LogSettings, config: Config) -> Path | None:
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
            logger_obj.debug("Created log folder: %s", log_dir.as_posix())

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

    # Write the banner after handlers are attached
    write_banner(logger_obj, config)

    # Flush logs buffer from prior to logging initialization
    if "log_buffer" in globals() and 'log_buffer' in globals():
        class _ForwardToLogger(logging.Handler):
            def emit(self, record):
                logger_obj.handle(record)

        forward_handler = _ForwardToLogger()
        log_buffer.setTarget(forward_handler)
        log_buffer.flush()
        log_buffer.close()

    # Enforce max log count
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


def write_banner(logger_obj: logging.Logger, config: Config):
    """Writes a clean, unformatted session header directly to the output streams."""
    script_name = Path(__file__).name
    separator = "-" * 80
    header_text = (
        f"{separator}\n"
        f"SCRIPT          | {json.dumps(str(script_name))}\n"
        f"VERSION         | {__version__}\n"
        f"USER/HOST       | {os.getlogin()} on {socket.gethostname()}\n"
        f"EXECUTION START | {datetime.now().isoformat(sep='T', timespec='milliseconds')}\n"
        f"DIRECTORY       | {json.dumps(str(Path.cwd().as_posix()))}\n"
        f"PLATFORM        | {platform.system()} {platform.release()}\n"
        f"RUNTIME         | Python {sys.version.split()[0]}\n"
        f"{separator}\n"
    )

    for handler in logger_obj.handlers:
        if isinstance(handler, logging.StreamHandler):
            try:
                handler.stream.write(header_text)
                handler.flush()
            except Exception:
                continue


def bootstrap():
    exit_code = 0
    log_path = None
    script_path = Path(__file__)

    # 1. Initialize Config First
    config = Config()
    config_path = script_path.with_name(f"{script_path.stem}_config.json")
    if InternalSettings().use_config_file:
        config = load_config(config_path)

    try:
        # 2. Setup Logging (This now handles the banner)
        log_path = setup_logging(logger_obj=logger, log_settings=config.logs, config=config)

        # 3. Only log dynamic app data here
        logger.info("Configuration loaded: %s", config)

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
