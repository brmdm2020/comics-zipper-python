import sys
import os
import time
import threading
import psutil
from typing import List, Dict, Tuple, Optional
import logging
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QLabel, QProgressBar,
                             QFileDialog, QComboBox, QCheckBox, QSpinBox,
                             QTabWidget, QTextEdit, QTreeView, QHeaderView,
                             QMessageBox, QFrame, QSplitter, QTreeWidget,
                             QTreeWidgetItem, QRadioButton, QGroupBox,
                             QLineEdit, QStatusBar, QStyle)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QTimer, QModelIndex
from PyQt5.QtGui import QIcon, QFont, QPixmap, QColor, QPalette, QStandardItemModel, QStandardItem

# 导入其他模块
from compression import CompressionManager, CompressionTask
from filesystem import FileSystemScanner, FileSystemWatcher
from report import ReportGenerator
import zipfile

logger = logging.getLogger("ComicCompressor")


class SystemMonitor(QThread):
    """系统资源监控线程"""
    update_signal = pyqtSignal(dict)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.running = False
        self.process = psutil.Process(os.getpid())

    def run(self):
        self.running = True
        while self.running:
            try:
                # 获取CPU和内存使用率
                cpu_percent = psutil.cpu_percent(interval=None)
                memory_percent = psutil.virtual_memory().percent
                process_cpu_percent = self.process.cpu_percent() / psutil.cpu_count()
                process_memory = self.process.memory_info().rss / (1024 * 1024)  # MB

                stats = {
                    'system_cpu': cpu_percent,
                    'system_memory': memory_percent,
                    'process_cpu': process_cpu_percent,
                    'process_memory': process_memory
                }

                self.update_signal.emit(stats)
                time.sleep(1)
            except Exception as e:
                logger.error(f"系统监控错误: {e}")
                time.sleep(2)

    def stop(self):
        self.running = False


class DirectoryStructureModel(QStandardItemModel):
    """用于预览目录结构的模型"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHorizontalHeaderLabels(["目录结构"])
        self.scanner = FileSystemScanner()

    def load_directory(self, root_path):
        """加载目录结构"""
        self.clear()
        self.setHorizontalHeaderLabels(["目录结构"])

        root_item = self.invisibleRootItem()
        self._load_directory_recursive(root_path, root_item, 0)

    def _load_directory_recursive(self, path, parent_item, depth):
        """递归加载目录结构"""
        if depth > 3:  # 限制深度以避免UI卡顿
            return

        try:
            dir_name = os.path.basename(path)
            item = QStandardItem(dir_name)
            item.setData(path, Qt.UserRole)
            parent_item.appendRow(item)

            # 标记章节目录
            if self.scanner.is_chapter_directory(path):
                item.setBackground(QColor(200, 255, 200))  # 浅绿色
                item.setText(f"{dir_name} (将被压缩)")
                return  # 不再展开章节目录

            # 加载子目录
            for entry in os.scandir(path):
                if entry.is_dir() and entry.name not in self.scanner.excluded_dirs:
                    self._load_directory_recursive(entry.path, item, depth + 1)
        except (PermissionError, FileNotFoundError) as e:
            logger.warning(f"无法访问目录 {path}: {e}")


class CompressionWorker(QThread):
    """压缩工作线程"""
    progress_signal = pyqtSignal(float, CompressionTask)
    scanning_signal = pyqtSignal(float, str)
    completed_signal = pyqtSignal()
    error_signal = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.root_path = None
        self.compression_level = zipfile.ZIP_DEFLATED
        self.preserve_timestamp = True
        self.rename_pattern = False
        self.max_workers = os.cpu_count()
        self.running = False
        self.manager = None
        self.paused = False
        self.total_images = 0
        self.processed_images = 0
        self.start_time = None
        self.end_time = None

    def configure(self, root_path, compression_level, preserve_timestamp,
                  rename_pattern, max_workers):
        """配置压缩任务"""
        self.root_path = root_path
        self.compression_level = compression_level
        self.preserve_timestamp = preserve_timestamp
        self.rename_pattern = rename_pattern
        self.max_workers = max_workers

    def run(self):
        """执行压缩任务"""
        self.running = True
        self.paused = False
        self.start_time = time.time()

        try:
            # 初始化管理器
            self.manager = CompressionManager(
                max_workers=self.max_workers,
                update_callback=self._on_task_update
            )

            # 扫描文件系统
            scanner = FileSystemScanner()
            chapters = scanner.scan_for_comic_directories(
                self.root_path,
                progress_callback=self._on_scanning_progress
            )

            if not chapters:
                self.error_signal.emit("未找到可压缩的目录")
                return

            # 准备任务
            tasks = scanner.prepare_compression_tasks(chapters, self.rename_pattern)

            # 计算总图片数
            self.total_images = 0
            for source_dir, _ in tasks:
                image_count, _ = self.manager.count_images_in_directory(source_dir)
                self.total_images += image_count

            # 添加任务
            for source_dir, target_zip in tasks:
                self.manager.add_task(
                    source_dir,
                    target_zip,
                    self.preserve_timestamp,
                    self.compression_level,
                    self.rename_pattern
                )

            # 开始压缩
            self.manager.start()

            # 等待完成
            while self.running and self.manager.get_progress() < 1.0:
                if self.paused:
                    time.sleep(0.5)
                    continue

                time.sleep(0.1)

            self.end_time = time.time()

            if self.running:
                self.completed_signal.emit()

        except Exception as e:
            logger.error(f"压缩过程出错: {e}")
            self.error_signal.emit(f"压缩过程出错: {e}")

        self.running = False

    def _on_task_update(self, progress, task):
        """任务更新回调"""
        self.processed_images += task.image_count
        self.progress_signal.emit(progress, task)

    def _on_scanning_progress(self, progress, status):
        """扫描进度回调"""
        self.scanning_signal.emit(progress, status)

    def pause(self):
        """暂停任务"""
        self.paused = True
        if self.manager:
            self.manager.pause()

    def resume(self):
        """恢复任务"""
        self.paused = False
        if self.manager:
            self.manager.resume()

    def cancel(self):
        """取消任务"""
        self.running = False
        if self.manager:
            self.manager.cancel()


class MainWindow(QMainWindow):
    """主窗口类"""

    def __init__(self):
        super().__init__()

        # 设置窗口属性
        self.setWindowTitle("漫画文件夹批量压缩工具")
        self.setMinimumSize(900, 700)

        # 初始化变量
        self.root_path = None
        self.report_generator = ReportGenerator()
        self.worker = None
        self.system_monitor = SystemMonitor()
        self.system_monitor.update_signal.connect(self.update_system_stats)
        self.system_monitor.start()

        # 创建UI
        self.setup_ui()

    def setup_ui(self):
        """创建用户界面"""
        # 创建中央窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 创建主布局
        main_layout = QVBoxLayout(central_widget)

        # 创建选项卡窗口部件
        tabs = QTabWidget()
        main_layout.addWidget(tabs)

        # 创建选项卡
        task_tab = QWidget()
        preview_tab = QWidget()
        logs_tab = QWidget()
        stats_tab = QWidget()

        tabs.addTab(task_tab, "任务")
        tabs.addTab(preview_tab, "目录预览")
        tabs.addTab(logs_tab, "日志")
        tabs.addTab(stats_tab, "统计")

        # 设置任务选项卡
        self.setup_task_tab(task_tab)

        # 设置预览选项卡
        self.setup_preview_tab(preview_tab)

        # 设置日志选项卡
        self.setup_logs_tab(logs_tab)

        # 设置统计选项卡
        self.setup_stats_tab(stats_tab)

        # 创建状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        # 添加系统资源标签到状态栏
        self.cpu_label = QLabel("CPU: 0%")
        self.memory_label = QLabel("内存: 0 MB")
        self.statusBar.addPermanentWidget(self.cpu_label)
        self.statusBar.addPermanentWidget(self.memory_label)

        # 设置初始状态
        self.statusBar.showMessage("准备就绪")

    def setup_task_tab(self, tab):
        """设置任务选项卡"""
        layout = QVBoxLayout(tab)

        # 目录选择部分
        dir_group = QGroupBox("目录选择")
        dir_layout = QHBoxLayout(dir_group)

        self.dir_path_edit = QLineEdit()
        self.dir_path_edit.setReadOnly(True)
        self.dir_path_edit.setPlaceholderText("请选择一个包含漫画文件夹的目录")

        browse_button = QPushButton("浏览...")
        browse_button.clicked.connect(self.browse_directory)

        dir_layout.addWidget(self.dir_path_edit, 3)
        dir_layout.addWidget(browse_button, 1)

        layout.addWidget(dir_group)

        # 压缩选项部分
        options_group = QGroupBox("压缩选项")
        options_layout = QVBoxLayout(options_group)

        # 压缩级别选择
        level_layout = QHBoxLayout()
        level_label = QLabel("压缩级别:")
        self.level_combo = QComboBox()
        self.level_combo.addItem("存储 (不压缩)", zipfile.ZIP_STORED)
        self.level_combo.addItem("快速 (低压缩率)", zipfile.ZIP_DEFLATED)
        self.level_combo.addItem("适中 (平衡)", zipfile.ZIP_DEFLATED)
        self.level_combo.addItem("最佳 (高压缩率)", zipfile.ZIP_DEFLATED)
        self.level_combo.setCurrentIndex(2)  # 默认选择"适中"

        level_layout.addWidget(level_label)
        level_layout.addWidget(self.level_combo)
        level_layout.addStretch()

        # 线程数选择
        threads_layout = QHBoxLayout()
        threads_label = QLabel("并行任务数:")
        self.threads_spin = QSpinBox()
        self.threads_spin.setMinimum(1)
        self.threads_spin.setMaximum(os.cpu_count() * 2)
        self.threads_spin.setValue(os.cpu_count())

        threads_layout.addWidget(threads_label)
        threads_layout.addWidget(self.threads_spin)
        threads_layout.addStretch()

        # 其他选项
        self.timestamp_check = QCheckBox("保留原始时间戳")
        self.timestamp_check.setChecked(True)

        self.rename_check = QCheckBox("自动重命名数字章节 (例如: 1.zip → 第1章.zip)")
        self.rename_check.setChecked(True)

        options_layout.addLayout(level_layout)
        options_layout.addLayout(threads_layout)
        options_layout.addWidget(self.timestamp_check)
        options_layout.addWidget(self.rename_check)

        layout.addWidget(options_group)

        # 进度信息部分
        progress_group = QGroupBox("处理进度")
        progress_layout = QVBoxLayout(progress_group)

        # 总体进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setFormat("%p% - %v/%m 任务完成")

        # 当前任务标签
        self.current_task_label = QLabel("等待开始...")

        # 状态信息
        info_layout = QHBoxLayout()

        self.images_label = QLabel("图片: 0/0")
        self.time_label = QLabel("耗时: 00:00:00")
        self.eta_label = QLabel("剩余: --:--:--")

        info_layout.addWidget(self.images_label)
        info_layout.addWidget(self.time_label)
        info_layout.addWidget(self.eta_label)

        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.current_task_label)
        progress_layout.addLayout(info_layout)

        layout.addWidget(progress_group)

        # 控制按钮
        buttons_layout = QHBoxLayout()

        self.start_button = QPushButton("开始处理")
        self.start_button.clicked.connect(self.start_compression)

        self.pause_button = QPushButton("暂停")
        self.pause_button.clicked.connect(self.toggle_pause)
        self.pause_button.setEnabled(False)

        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self.cancel_compression)
        self.cancel_button.setEnabled(False)

        buttons_layout.addWidget(self.start_button)
        buttons_layout.addWidget(self.pause_button)
        buttons_layout.addWidget(self.cancel_button)

        layout.addLayout(buttons_layout)
        layout.addStretch()

    def setup_preview_tab(self, tab):
        """设置预览选项卡"""
        layout = QVBoxLayout(tab)

        # 预览控制
        controls_layout = QHBoxLayout()

        refresh_button = QPushButton("刷新预览")
        refresh_button.clicked.connect(self.refresh_preview)

        controls_layout.addWidget(refresh_button)
        controls_layout.addStretch()

        # 目录树视图
        self.tree_model = DirectoryStructureModel()
        self.tree_view = QTreeView()
        self.tree_view.setModel(self.tree_model)
        self.tree_view.setAnimated(True)
        self.tree_view.setHeaderHidden(False)
        self.tree_view.header().setSectionResizeMode(QHeaderView.Stretch)

        layout.addLayout(controls_layout)
        layout.addWidget(self.tree_view)

    def setup_logs_tab(self, tab):
        """设置日志选项卡"""
        layout = QVBoxLayout(tab)

        # 日志文本区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        # 添加自定义处理程序以将日志消息重定向到文本区域
        logger.addHandler(LogTextHandler(self.log_text))

        layout.addWidget(self.log_text)

        # 清除按钮
        clear_button = QPushButton("清除日志")
        clear_button.clicked.connect(self.log_text.clear)

        layout.addWidget(clear_button)

    def setup_stats_tab(self, tab):
        """设置统计选项卡"""
        layout = QVBoxLayout(tab)

        # 统计信息
        self.stats_text = QTextEdit()
        self.stats_text.setReadOnly(True)

        layout.addWidget(self.stats_text)

        # 导出报告按钮
        export_button = QPushButton("导出Excel报告")
        export_button.clicked.connect(self.export_report)

        layout.addWidget(export_button)

    def browse_directory(self):
        """浏览文件夹对话框"""
        dir_path = QFileDialog.getExistingDirectory(
            self, "选择漫画所在目录", "", QFileDialog.ShowDirsOnly
        )

        if dir_path:
            self.root_path = dir_path
            self.dir_path_edit.setText(dir_path)
            self.refresh_preview()

    def refresh_preview(self):
        """刷新目录预览"""
        if not self.root_path:
            return

        try:
            self.tree_model.load_directory(self.root_path)
            self.tree_view.expandToDepth(1)
        except Exception as e:
            logger.error(f"加载目录预览失败: {e}")
            QMessageBox.warning(self, "错误", f"加载目录预览失败: {e}")

    def start_compression(self):
        """开始压缩任务"""
        if not self.root_path:
            QMessageBox.warning(self, "错误", "请先选择一个目录")
            return

        # 获取选项
        compression_level = self.level_combo.currentData()
        preserve_timestamp = self.timestamp_check.isChecked()
        rename_pattern = self.rename_check.isChecked()
        max_workers = self.threads_spin.value()

        # 创建工作线程
        self.worker = CompressionWorker()
        self.worker.configure(
            self.root_path,
            compression_level,
            preserve_timestamp,
            rename_pattern,
            max_workers
        )

        # 连接信号
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.scanning_signal.connect(self.update_scanning_progress)
        self.worker.completed_signal.connect(self.on_compression_completed)
        self.worker.error_signal.connect(self.on_compression_error)

        # 更新UI状态
        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.cancel_button.setEnabled(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(100)
        self.statusBar.showMessage("正在扫描文件系统...")

        # 开始任务
        self.worker.start()

        # 启动计时器
        self.start_time = time.time()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)  # 每秒更新一次

    def toggle_pause(self):
        """暂停/恢复任务"""
        if not self.worker:
            return

        if self.worker.paused:
            self.worker.resume()
            self.pause_button.setText("暂停")
            self.statusBar.showMessage("任务已恢复")
        else:
            self.worker.pause()
            self.pause_button.setText("恢复")
            self.statusBar.showMessage("任务已暂停")

    def cancel_compression(self):
        """取消任务"""
        if not self.worker:
            return

        reply = QMessageBox.question(
            self, "确认取消",
            "确定要取消当前任务吗？已处理的文件将保持压缩状态。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.worker.cancel()
            self.statusBar.showMessage("任务已取消")
            self.reset_ui_state()

    def update_scanning_progress(self, progress, status):
        """更新扫描进度"""
        self.progress_bar.setValue(int(progress * 100))
        self.current_task_label.setText(status)

    def update_progress(self, progress, task):
        """更新压缩进度"""
        self.progress_bar.setValue(int(progress * 100))

        # 更新当前任务标签
        if task.status == "completed":
            status_text = "完成"
        elif task.status == "failed":
            status_text = f"失败: {task.error}"
        else:
            status_text = "处理中"

        self.current_task_label.setText(f"{task.source_path} - {status_text}")

        # 更新图片计数
        if hasattr(self.worker, 'total_images') and hasattr(self.worker, 'processed_images'):
            self.images_label.setText(f"图片: {self.worker.processed_images}/{self.worker.total_images}")

        # 更新统计信息
        if task.status == "completed":
            self.update_stats(task)

    def update_time(self):
        """更新计时器"""
        if not hasattr(self, 'start_time') or not self.worker or not self.worker.running:
            return

        elapsed = time.time() - self.start_time
        elapsed_str = str(timedelta(seconds=int(elapsed)))
        self.time_label.setText(f"耗时: {elapsed_str}")

        # 估算剩余时间
        if hasattr(self.worker, 'total_images') and self.worker.total_images > 0 and self.worker.processed_images > 0:
            progress = self.worker.processed_images / self.worker.total_images
            if progress > 0:
                total_time = elapsed / progress
                remaining = total_time - elapsed
                if remaining > 0:
                    remaining_str = str(timedelta(seconds=int(remaining)))
                    self.eta_label.setText(f"剩余: {remaining_str}")
                else:
                    self.eta_label.setText("剩余: 完成中...")

    def update_system_stats(self, stats):
        """更新系统资源统计信息"""
        self.cpu_label.setText(f"CPU: {stats['process_cpu']:.1f}%")
        self.memory_label.setText(f"内存: {stats['process_memory']:.1f} MB")

    def update_stats(self, task=None):
        """更新统计信息"""
        if task:
            # 将任务信息添加到报告生成器
            self.report_generator.add_task_result(task)

        # 如果没有任务，清空统计信息
        if not hasattr(self, 'worker') or not self.worker:
            self.stats_text.clear()
            return

        # 获取统计数据
        if hasattr(self.worker, 'manager') and self.worker.manager:
            stats = self.worker.manager.get_stats()

            # 生成统计文本
            stats_text = "处理统计:\n\n"
            stats_text += f"总任务数: {stats['total_tasks']}\n"
            stats_text += f"已完成: {stats['completed_tasks']}\n"
            stats_text += f"失败: {stats['failed_tasks']}\n"
            stats_text += f"总图片数: {stats['total_images']}\n"

            if stats['total_original_size'] > 0:
                original_mb = stats['total_original_size'] / (1024 * 1024)
                compressed_mb = stats['total_compressed_size'] / (1024 * 1024)
                savings_mb = original_mb - compressed_mb
                ratio = stats['compression_ratio'] * 100

                stats_text += f"\n压缩前总大小: {original_mb:.2f} MB\n"
                stats_text += f"压缩后总大小: {compressed_mb:.2f} MB\n"
                stats_text += f"节省空间: {savings_mb:.2f} MB ({100 - ratio:.1f}%)\n"

            self.stats_text.setText(stats_text)

    def export_report(self):
        """导出Excel报告"""
        if not self.root_path or not hasattr(self, 'worker') or not self.worker:
            QMessageBox.warning(self, "错误", "请先运行压缩任务")
            return

        try:
            # 生成报告文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            report_path = os.path.join(self.root_path, f"漫画压缩报告_{timestamp}.xlsx")

            # 生成报告
            self.report_generator.generate_report(report_path)

            QMessageBox.information(
                self, "报告已生成",
                f"报告已保存至:\n{report_path}"
            )

        except Exception as e:
            logger.error(f"生成报告失败: {e}")
            QMessageBox.warning(self, "错误", f"生成报告失败: {e}")

    def on_compression_completed(self):
        """压缩任务完成回调"""
        self.statusBar.showMessage("任务已完成")
        logger.info("所有任务已完成")
        self.reset_ui_state()

        # 更新最终统计
        self.update_stats()

        QMessageBox.information(self, "完成", "所有压缩任务已完成")

        # 提示导出报告
        reply = QMessageBox.question(
            self, "导出报告",
            "是否要导出Excel报告？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes
        )

        if reply == QMessageBox.Yes:
            self.export_report()

    def on_compression_error(self, error_message):
        """压缩任务错误回调"""
        self.statusBar.showMessage(f"错误: {error_message}")
        logger.error(error_message)
        self.reset_ui_state()

        QMessageBox.critical(self, "错误", error_message)

    def reset_ui_state(self):
        """重置UI状态"""
        if hasattr(self, 'timer') and self.timer.isActive():
            self.timer.stop()

        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.pause_button.setText("暂停")
        self.cancel_button.setEnabled(False)

    def closeEvent(self, event):
        """窗口关闭事件"""
        # 停止系统监控线程
        if hasattr(self, 'system_monitor'):
            self.system_monitor.stop()
            self.system_monitor.wait()

        # 停止工作线程
        if hasattr(self, 'worker') and self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "确认退出",
                "任务正在进行中，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.Yes:
                self.worker.cancel()
                self.worker.wait()
            else:
                event.ignore()
                return

        event.accept()


class LogTextHandler(logging.Handler):
    """将日志消息发送到QTextEdit的处理程序"""

    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    def emit(self, record):
        """发送日志记录到文本编辑器"""
        msg = self.format(record)

        # 不同级别的日志使用不同颜色
        color = "black"
        if record.levelno >= logging.ERROR:
            color = "red"
        elif record.levelno >= logging.WARNING:
            color = "orange"
        elif record.levelno >= logging.INFO:
            color = "blue"

        # 添加带颜色的文本
        self.text_edit.append(f'<span style="color:{color}">{msg}</span>')