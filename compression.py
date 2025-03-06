import os
import zipfile
import shutil
import hashlib
import time
import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import re

logger = logging.getLogger("ComicCompressor")


class CompressionTask:
    """表示单个压缩任务的类"""

    def __init__(self, source_path, target_path, preserve_timestamp=True,
                 compression_level=zipfile.ZIP_DEFLATED, rename_pattern=None):
        self.source_path = source_path
        self.target_path = target_path
        self.preserve_timestamp = preserve_timestamp
        self.compression_level = compression_level
        self.rename_pattern = rename_pattern
        self.status = "pending"
        self.error = None
        self.start_time = None
        self.end_time = None
        self.image_count = 0
        self.original_size = 0
        self.compressed_size = 0
        self.md5 = None

    def to_dict(self):
        """将任务信息转换为字典格式"""
        return {
            "source_path": self.source_path,
            "target_path": self.target_path,
            "status": self.status,
            "error": str(self.error) if self.error else None,
            "start_time": self.start_time,
            "end_time": self.end_time,
            "duration": (self.end_time - self.start_time) if self.end_time and self.start_time else None,
            "image_count": self.image_count,
            "original_size": self.original_size,
            "compressed_size": self.compressed_size,
            "compression_ratio": self.original_size / self.compressed_size if self.compressed_size else 0,
            "md5": self.md5
        }


class CompressionManager:
    """管理压缩任务的类"""

    def __init__(self, max_workers=None, update_callback=None):
        self.max_workers = max_workers or os.cpu_count()
        self.update_callback = update_callback
        self.tasks = []
        self.completed_tasks = 0
        self.total_tasks = 0
        self.running = False
        self.paused = False
        self.lock = Lock()
        self.executor = None
        self.futures = []

        # 图片文件扩展名检测
        self.image_extensions = {
            '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp',
            '.tiff', '.tif', '.ico', '.jfif', '.heic'
        }

        # 用于重命名的正则表达式
        self.number_only_pattern = re.compile(r'^(\d+)\.zip$')

    def is_image_file(self, filename):
        """检查文件是否为图片文件"""
        return os.path.splitext(filename.lower())[1] in self.image_extensions

    def count_images_in_directory(self, directory):
        """计算目录中图片文件的数量和总大小"""
        count = 0
        total_size = 0
        try:
            for root, _, files in os.walk(directory):
                for file in files:
                    if self.is_image_file(file):
                        count += 1
                        file_path = os.path.join(root, file)
                        total_size += os.path.getsize(file_path)
        except Exception as e:
            logger.error(f"计算图片时出错: {e}")
        return count, total_size

    def calculate_md5(self, file_path):
        """计算文件的MD5哈希值"""
        hash_md5 = hashlib.md5()
        with open(file_path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
        return hash_md5.hexdigest()

    def compress_directory(self, task):
        """压缩目录到ZIP文件"""
        task.start_time = time.time()
        task.status = "running"

        try:
            # 计算图片数量和总大小
            task.image_count, task.original_size = self.count_images_in_directory(task.source_path)

            # 获取源目录的时间戳
            src_mtime = os.path.getmtime(task.source_path)

            # 创建临时目标路径，避免直接覆盖
            temp_target_path = f"{task.target_path}.temp"

            # 创建ZIP文件
            with zipfile.ZipFile(temp_target_path, 'w', task.compression_level) as zipf:
                # 直接添加图片文件，不保留目录结构
                for root, _, files in os.walk(task.source_path):
                    for file in files:
                        if self.is_image_file(file):
                            file_path = os.path.join(root, file)
                            # 将文件添加到ZIP的根目录下
                            zipf.write(file_path, arcname=file)

            # 检查是否需要重命名
            final_target_path = task.target_path
            if task.rename_pattern:
                base_name = os.path.basename(task.target_path)
                match = self.number_only_pattern.match(base_name)
                if match:
                    # 如果文件名只包含数字，则重命名为"第X章.zip"
                    number = match.group(1)
                    new_name = f"第{number}章.zip"
                    final_target_path = os.path.join(os.path.dirname(task.target_path), new_name)

            # 将临时文件重命名为最终文件
            if os.path.exists(final_target_path):
                os.remove(final_target_path)
            os.rename(temp_target_path, final_target_path)

            # 更新任务目标路径（如果发生了重命名）
            task.target_path = final_target_path

            # 保留原始时间戳
            if task.preserve_timestamp:
                os.utime(final_target_path, (src_mtime, src_mtime))

            # 计算MD5和压缩后大小
            task.md5 = self.calculate_md5(final_target_path)
            task.compressed_size = os.path.getsize(final_target_path)

            # 移除原目录
            shutil.rmtree(task.source_path)

            task.status = "completed"

        except Exception as e:
            task.status = "failed"
            task.error = e
            logger.error(f"压缩失败: {e}")

        task.end_time = time.time()
        return task

    def add_task(self, source_path, target_path, preserve_timestamp=True,
                 compression_level=zipfile.ZIP_DEFLATED, rename_pattern=None):
        """添加压缩任务"""
        task = CompressionTask(
            source_path=source_path,
            target_path=target_path,
            preserve_timestamp=preserve_timestamp,
            compression_level=compression_level,
            rename_pattern=rename_pattern
        )
        self.tasks.append(task)
        return task

    def start(self):
        """启动所有压缩任务"""
        if self.running:
            return

        self.running = True
        self.paused = False
        self.completed_tasks = 0
        self.total_tasks = len(self.tasks)

        if self.total_tasks == 0:
            return

        self.executor = ThreadPoolExecutor(max_workers=self.max_workers)
        self.futures = []

        # 提交所有任务
        for task in self.tasks:
            if task.status == "pending":
                future = self.executor.submit(self.compress_directory, task)
                future.add_done_callback(self._task_completed)
                self.futures.append(future)

    def _task_completed(self, future):
        """任务完成回调"""
        with self.lock:
            self.completed_tasks += 1
            progress = self.completed_tasks / self.total_tasks

            # 如果有回调函数，通知进度更新
            if self.update_callback:
                task = future.result()
                self.update_callback(progress, task)

    def pause(self):
        """暂停所有任务"""
        self.paused = True
        if self.executor:
            self.executor.shutdown(wait=False, cancel_futures=True)
            self.executor = None

    def resume(self):
        """恢复暂停的任务"""
        if not self.paused:
            return

        # 重新启动未完成的任务
        self.executor = ThreadPoolExecutor(max_workers=self.max_workers)
        self.futures = []

        for task in self.tasks:
            if task.status == "pending":
                future = self.executor.submit(self.compress_directory, task)
                future.add_done_callback(self._task_completed)
                self.futures.append(future)

        self.paused = False

    def cancel(self):
        """取消所有任务"""
        self.running = False
        if self.executor:
            self.executor.shutdown(wait=False, cancel_futures=True)
            self.executor = None

        # 重置所有未完成任务的状态
        for task in self.tasks:
            if task.status == "running":
                task.status = "cancelled"

    def get_progress(self):
        """获取当前进度"""
        if self.total_tasks == 0:
            return 0
        return self.completed_tasks / self.total_tasks

    def get_stats(self):
        """获取统计信息"""
        completed = [task for task in self.tasks if task.status == "completed"]
        failed = [task for task in self.tasks if task.status == "failed"]

        total_images = sum(task.image_count for task in completed)
        total_original_size = sum(task.original_size for task in completed)
        total_compressed_size = sum(task.compressed_size for task in completed)

        return {
            "total_tasks": self.total_tasks,
            "completed_tasks": self.completed_tasks,
            "failed_tasks": len(failed),
            "total_images": total_images,
            "total_original_size": total_original_size,
            "total_compressed_size": total_compressed_size,
            "compression_ratio": total_compressed_size / total_original_size if total_original_size else 0
        }