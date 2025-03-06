import os
import logging
import time
import datetime
import zipfile
import hashlib
import re
import shutil
import tempfile
from typing import List, Dict, Tuple, Optional, Set, Any

logger = logging.getLogger("ComicCompressor")


def setup_logging(log_file=None, console_level=logging.INFO):
    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)

    # Configure console output
    console = logging.StreamHandler()
    console.setLevel(console_level)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console.setFormatter(formatter)
    root_logger.addHandler(console)

    # Configure file output if log_file is provided
    if log_file:
        log_dir = os.path.dirname(log_file)
        # Only try to create directory if there's actually a directory component
        if log_dir:
            os.makedirs(log_dir, exist_ok=True)

        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)


def format_size(size_bytes: int) -> str:
    """将字节大小格式化为人类可读的格式"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.2f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.2f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"


def format_time(seconds: float) -> str:
    """将秒数格式化为人类可读的时间格式"""
    if seconds < 60:
        return f"{seconds:.1f} 秒"
    elif seconds < 3600:
        minutes = seconds / 60
        return f"{int(minutes)} 分 {int(seconds % 60)} 秒"
    else:
        hours = seconds / 3600
        minutes = (seconds % 3600) / 60
        return f"{int(hours)} 小时 {int(minutes)} 分"


def is_valid_zip(file_path: str) -> bool:
    """检查ZIP文件是否有效"""
    try:
        with zipfile.ZipFile(file_path, 'r') as zipf:
            # 测试ZIP文件完整性
            result = zipf.testzip()
            if result is not None:
                logger.warning(f"ZIP文件损坏，第一个坏文件是 {result}")
                return False
            return True
    except zipfile.BadZipFile:
        logger.warning(f"无效的ZIP文件: {file_path}")
        return False
    except Exception as e:
        logger.warning(f"检查ZIP文件时出错: {e}")
        return False


def calculate_md5(file_path: str, chunk_size: int = 8192) -> str:
    """计算文件的MD5哈希值"""
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(chunk_size), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()
    except Exception as e:
        logger.error(f"计算MD5时出错: {e}")
        return ""


def find_images_recursively(directory: str, image_extensions: Set[str]) -> Tuple[List[str], int]:
    """
    递归查找目录中的所有图片文件
    返回: (图片文件路径列表, 总大小)
    """
    images = []
    total_size = 0

    try:
        for root, _, files in os.walk(directory):
            for file in files:
                if os.path.splitext(file.lower())[1] in image_extensions:
                    file_path = os.path.join(root, file)
                    images.append(file_path)
                    total_size += os.path.getsize(file_path)
    except Exception as e:
        logger.error(f"查找图片时出错: {e}")

    return images, total_size


def create_backup(directory: str) -> Optional[str]:
    """创建目录的备份"""
    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_dir = f"{directory}_backup_{timestamp}"
        shutil.copytree(directory, backup_dir)
        logger.info(f"已创建备份: {backup_dir}")
        return backup_dir
    except Exception as e:
        logger.error(f"创建备份时出错: {e}")
        return None


def restore_from_backup(backup_dir: str, target_dir: str) -> bool:
    """从备份恢复目录"""
    try:
        if os.path.exists(target_dir):
            shutil.rmtree(target_dir)
        shutil.copytree(backup_dir, target_dir)
        logger.info(f"已从备份恢复: {target_dir}")
        return True
    except Exception as e:
        logger.error(f"从备份恢复时出错: {e}")
        return False


def is_path_too_long(path: str, max_length: int = 260) -> bool:
    """检查路径是否过长（Windows路径长度限制）"""
    return len(os.path.abspath(path)) > max_length


def sanitize_filename(filename: str) -> str:
    """清理文件名，移除不允许的字符"""
    # 替换Windows文件系统不允许的字符
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', filename)
    # 移除结尾的空格和点
    sanitized = sanitized.rstrip('. ')
    return sanitized


def ensure_directory_exists(directory: str) -> bool:
    """确保目录存在，如果不存在则创建"""
    try:
        os.makedirs(directory, exist_ok=True)
        return True
    except Exception as e:
        logger.error(f"创建目录失败: {e}")
        return False


def get_free_space(directory: str) -> int:
    """获取目录所在磁盘的可用空间（字节）"""
    try:
        if os.name == 'nt':  # Windows
            import ctypes
            free_bytes = ctypes.c_ulonglong(0)
            ctypes.windll.kernel32.GetDiskFreeSpaceExW(
                ctypes.c_wchar_p(directory), None, None, ctypes.pointer(free_bytes)
            )
            return free_bytes.value
        else:  # Unix/Linux/MacOS
            st = os.statvfs(directory)
            return st.f_bavail * st.f_frsize
    except Exception as e:
        logger.error(f"获取可用空间时出错: {e}")
        return 0


def is_enough_space(directory: str, required_bytes: int, safety_factor: float = 1.2) -> bool:
    """检查目录是否有足够的可用空间"""
    free_space = get_free_space(directory)
    # 添加安全系数，预留额外空间
    required_with_safety = required_bytes * safety_factor
    return free_space >= required_with_safety