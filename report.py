import os
import pandas as pd
import logging
from datetime import datetime

logger = logging.getLogger("ComicCompressor")


class ReportGenerator:
    """用于生成Excel报告的类"""

    def __init__(self):
        self.task_results = []
        self.summary = {
            "total_comics": set(),
            "total_chapters": 0,
            "total_images": 0,
            "total_original_size": 0,
            "total_compressed_size": 0,
            "start_time": None,
            "end_time": None,
            "failed_tasks": 0
        }

    def add_task_result(self, task):
        """添加一个任务的结果"""
        # 记录开始时间
        if not self.summary["start_time"] or (task.start_time and task.start_time < self.summary["start_time"]):
            self.summary["start_time"] = task.start_time

        # 记录结束时间
        if not self.summary["end_time"] or (task.end_time and task.end_time > self.summary["end_time"]):
            self.summary["end_time"] = task.end_time

        # 更新统计信息
        if task.status == "completed":
            comic_title = os.path.basename(os.path.dirname(task.source_path))
            self.summary["total_comics"].add(comic_title)
            self.summary["total_chapters"] += 1
            self.summary["total_images"] += task.image_count
            self.summary["total_original_size"] += task.original_size
            self.summary["total_compressed_size"] += task.compressed_size
        elif task.status == "failed":
            self.summary["failed_tasks"] += 1

        # 添加到任务结果列表
        task_data = {
            "漫画标题": os.path.basename(os.path.dirname(task.source_path)),
            "原章节名称": os.path.basename(task.source_path),
            "压缩文件名": os.path.basename(task.target_path),
            "图片总数": task.image_count,
            "压缩前大小(MB)": task.original_size / (1024 * 1024) if task.original_size else 0,
            "压缩后大小(MB)": task.compressed_size / (1024 * 1024) if task.compressed_size else 0,
            "压缩比例": task.compressed_size / task.original_size if task.original_size and task.compressed_size else 0,
            "压缩时间": datetime.fromtimestamp(task.end_time).strftime("%Y-%m-%d %H:%M:%S") if task.end_time else "",
            "耗时(秒)": task.end_time - task.start_time if task.end_time and task.start_time else 0,
            "MD5校验码": task.md5 or "",
            "原始路径": task.source_path,
            "状态": task.status,
            "错误信息": str(task.error) if task.error else ""
        }
        self.task_results.append(task_data)

    def generate_report(self, output_path):
        """生成Excel报告"""
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 章节明细表
                if self.task_results:
                    df_details = pd.DataFrame(self.task_results)
                    df_details.sort_values(by=["漫画标题", "原章节名称"], inplace=True)
                    df_details.to_excel(writer, sheet_name="章节明细", index=False)

                    # 设置列宽
                    worksheet = writer.sheets["章节明细"]
                    for i, col in enumerate(df_details.columns):
                        max_len = max(
                            df_details[col].astype(str).map(len).max(),
                            len(col)
                        ) + 2
                        worksheet.column_dimensions[chr(65 + i)].width = min(max_len, 50)

                # 汇总表
                summary_data = {
                    "统计项": [
                        "总漫画系列数",
                        "总章节数",
                        "成功章节数",
                        "失败章节数",
                        "总图片文件数",
                        "总原始大小(MB)",
                        "总压缩后大小(MB)",
                        "节省空间(MB)",
                        "平均压缩率",
                        "处理开始时间",
                        "处理结束时间",
                        "总处理耗时(分钟)"
                    ],
                    "数值": [
                        len(self.summary["total_comics"]),
                        self.summary["total_chapters"] + self.summary["failed_tasks"],
                        self.summary["total_chapters"],
                        self.summary["failed_tasks"],
                        self.summary["total_images"],
                        self.summary["total_original_size"] / (1024 * 1024),
                        self.summary["total_compressed_size"] / (1024 * 1024),
                        (self.summary["total_original_size"] - self.summary["total_compressed_size"]) / (1024 * 1024),
                        self.summary["total_compressed_size"] / self.summary["total_original_size"] if self.summary[
                            "total_original_size"] else 0,
                        datetime.fromtimestamp(self.summary["start_time"]).strftime("%Y-%m-%d %H:%M:%S") if
                        self.summary["start_time"] else "",
                        datetime.fromtimestamp(self.summary["end_time"]).strftime("%Y-%m-%d %H:%M:%S") if self.summary[
                            "end_time"] else "",
                        (self.summary["end_time"] - self.summary["start_time"]) / 60 if self.summary["end_time"] and
                                                                                        self.summary[
                                                                                            "start_time"] else 0
                    ]
                }
                df_summary = pd.DataFrame(summary_data)
                df_summary.to_excel(writer, sheet_name="汇总统计", index=False)

                # 设置列宽
                worksheet = writer.sheets["汇总统计"]
                worksheet.column_dimensions['A'].width = 20
                worksheet.column_dimensions['B'].width = 25

                # 漫画标题统计表
                if self.task_results:
                    # 按漫画标题分组统计
                    comic_stats = {}
                    for task in self.task_results:
                        title = task["漫画标题"]
                        if title not in comic_stats:
                            comic_stats[title] = {
                                "章节数": 0,
                                "图片总数": 0,
                                "原始大小(MB)": 0,
                                "压缩后大小(MB)": 0
                            }

                        if task["状态"] == "completed":
                            comic_stats[title]["章节数"] += 1
                            comic_stats[title]["图片总数"] += task["图片总数"]
                            comic_stats[title]["原始大小(MB)"] += task["压缩前大小(MB)"]
                            comic_stats[title]["压缩后大小(MB)"] += task["压缩后大小(MB)"]

                    # 创建漫画标题统计数据框
                    comic_stats_data = []
                    for title, stats in comic_stats.items():
                        comic_stats_data.append({
                            "漫画标题": title,
                            "章节数": stats["章节数"],
                            "图片总数": stats["图片总数"],
                            "原始大小(MB)": stats["原始大小(MB)"],
                            "压缩后大小(MB)": stats["压缩后大小(MB)"],
                            "节省空间(MB)": stats["原始大小(MB)"] - stats["压缩后大小(MB)"],
                            "压缩比例": stats["压缩后大小(MB)"] / stats["原始大小(MB)"] if stats["原始大小(MB)"] else 0
                        })

                    if comic_stats_data:
                        df_comics = pd.DataFrame(comic_stats_data)
                        df_comics.sort_values(by=["漫画标题"], inplace=True)
                        df_comics.to_excel(writer, sheet_name="漫画标题统计", index=False)

                        # 设置列宽
                        worksheet = writer.sheets["漫画标题统计"]
                        for i, col in enumerate(df_comics.columns):
                            max_len = max(
                                df_comics[col].astype(str).map(len).max(),
                                len(col)
                            ) + 2
                            worksheet.column_dimensions[chr(65 + i)].width = min(max_len, 30)

                # 文件类型分布表
                if self.task_results:
                    # 统计图片类型
                    file_types = {}
                    for task in self.task_results:
                        if task["状态"] != "completed":
                            continue

                        source_path = task["原始路径"]
                        try:
                            # 统计目录中各类型图片的数量
                            for root, _, files in os.walk(source_path):
                                for file in files:
                                    ext = os.path.splitext(file.lower())[1]
                                    if ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff', '.tif']:
                                        file_types[ext] = file_types.get(ext, 0) + 1
                        except:
                            # 如果目录已被删除，无法统计
                            pass

                    if file_types:
                        file_types_data = [{"文件类型": ext, "数量": count} for ext, count in file_types.items()]
                        df_file_types = pd.DataFrame(file_types_data)
                        df_file_types.sort_values(by=["数量"], ascending=False, inplace=True)
                        df_file_types.to_excel(writer, sheet_name="文件类型分布", index=False)

            logger.info(f"报告已生成: {output_path}")
            return True

        except Exception as e:
            logger.error(f"生成报告失败: {e}")
            return False