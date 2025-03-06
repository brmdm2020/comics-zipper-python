import os
import logging
import time
from pathlib import Path
from typing import List, Tuple, Dict, Set, Optional, Callable

logger = logging.getLogger("ComicCompressor")


class FileSystemScanner:
    """用于扫描文件系统并识别需要压缩的目录的类"""

    def __init__(self, max_depth: int = 10):
        self.max_depth = max_depth
        # 图片文件扩展名集合
        self.image_extensions = {
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp',
            '.tiff', '.tif', '.ico', '.jfif', '.heic'
        }
        # 排除的目录名
        self.excluded_dirs = {'.git', '__pycache__', '.svn', 'node_modules', '.vscode'}
        # 已发现的章节缓存，用于避免重复扫描
        self.chapter_cache = {}

    def is_image_file(self, file_path: str) -> bool:
        """检查文件是否为图片文件"""
        return os.path.splitext(file_path.lower())[1] in self.image_extensions

    def is_chapter_directory(self, dir_path: str) -> bool:
        """检查目录是否为漫画章节目录（包含图片的最底层目录）"""
        # 如果已经在缓存中，直接返回结果
        if dir_path in self.chapter_cache:
            return self.chapter_cache[dir_path]

        # 检查目录是否直接包含图片
        has_images = False
        has_subdirs = False

        try:
            for item in os.scandir(dir_path):
                if item.is_file() and self.is_image_file(item.name):
                    has_images = True
                elif item.is_dir() and item.name not in self.excluded_dirs:
                    has_subdirs = True

                    # 如果已经找到图片和子目录，不需要继续扫描
                    if has_images:
                        break
        except (PermissionError, FileNotFoundError) as e:
            logger.warning(f"无法访问目录 {dir_path}: {e}")
            self.chapter_cache[dir_path] = False
            return False

        # 如果目录直接包含图片，但不包含子目录，则认为是章节目录
        result = has_images and not has_subdirs

        # 如果目录包含子目录，需要检查是否这些子目录包含图片
        if has_images and has_subdirs:
            # 章节目录可能包含子目录（如"pages"），但仍然是最底层的章节目录
            result = True

        # 缓存结果
        self.chapter_cache[dir_path] = result
        return result

    def scan_for_comic_directories(self,
                                   root_path: str,
                                   progress_callback: Optional[Callable[[float, str], None]] = None) -> List[
        Tuple[str, str, str]]:
        """
        扫描漫画目录，寻找需要压缩的章节目录
        返回格式: [(漫画标题目录, 章节目录, 相对目录名)]
        """
        comic_chapters = []
        total_dirs = 0
        processed_dirs = 0

        # 首先计算总目录数以便更新进度
        for root, dirs, _ in os.walk(root_path):
            total_dirs += len(dirs)

            # 从扫描中排除不需要的目录
            dirs[:] = [d for d in dirs if d not in self.excluded_dirs]

            # 限制递归深度
            if root.count(os.sep) - root_path.count(os.sep) >= self.max_depth:
                dirs[:] = []

        logger.info(f"扫描到 {total_dirs} 个目录")

        # 开始实际扫描
        for root, dirs, _ in os.walk(root_path):
            # 排除不需要的目录
            dirs[:] = [d for d in dirs if d not in self.excluded_dirs]

            # 限制递归深度
            if root.count(os.sep) - root_path.count(os.sep) >= self.max_depth:
                dirs[:] = []
                continue

            for dir_name in dirs:
                dir_path = os.path.join(root, dir_name)
                processed_dirs += 1

                # 更新进度
                if progress_callback and total_dirs > 0:
                    progress = processed_dirs / total_dirs
                    progress_callback(progress, f"扫描: {dir_path}")

                # 检查是否为章节目录
                if self.is_chapter_directory(dir_path):
                    # 找到章节目录，确定漫画标题目录
                    # 假设章节目录的父目录是漫画标题目录
                    comic_title_dir = os.path.dirname(dir_path)
                    comic_chapters.append((comic_title_dir, dir_path, dir_name))

        return comic_chapters

    def prepare_compression_tasks(self,
                                  chapters: List[Tuple[str, str, str]],
                                  rename_pattern: bool = False) -> List[Tuple[str, str]]:
        """
        准备压缩任务列表
        返回: [(源目录, 目标ZIP路径)]
        """
        tasks = []

        for comic_title_dir, chapter_dir, chapter_name in chapters:
            # 构建目标ZIP文件路径
            target_zip = os.path.join(comic_title_dir, f"{chapter_name}.zip")
            tasks.append((chapter_dir, target_zip))

        return tasks


class FileSystemWatcher:
    """用于监视文件系统更改的类，可用于增量处理"""

    def __init__(self):
        self.last_modified_times = {}

    def snapshot_directory(self, directory: str) -> Dict[str, float]:
        """记录目录中所有文件的最后修改时间"""
        snapshot = {}
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                try:
                    snapshot[file_path] = os.path.getmtime(file_path)
                except (FileNotFoundError, PermissionError):
                    pass
        return snapshot

    def get_changed_files(self, directory: str) -> Set[str]:
        """获取自上次快照以来更改的文件"""
        current_snapshot = self.snapshot_directory(directory)
        changed_files = set()

        # 检查新增和修改的文件
        for file_path, mtime in current_snapshot.items():
            if file_path not in self.last_modified_times or mtime > self.last_modified_times[file_path]:
                changed_files.add(file_path)

        # 更新快照
        self.last_modified_times = current_snapshot
        return changed_files