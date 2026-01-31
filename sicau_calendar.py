# -*- coding: utf-8 -*-
"""
校历桌面应用 - 支持自定义导入校历和课表
支持系统托盘、开机自启动、课程表、上课提醒
适用于各类高校教师和学生使用

开发者：肖诗顺
本程序为开源程序，供教师和学生免费使用

注意：用户可通过导入向导自定义本校校历和课表数据
"""

import sys
import os
import json
import winreg
import winsound
import re
from datetime import datetime, date, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTableWidget, QTableWidgetItem, QHeaderView, QFrame,
    QCalendarWidget, QGroupBox, QScrollArea, QSystemTrayIcon,
    QMenu, QAction, QMessageBox, QCheckBox, QDialog, QPushButton,
    QDialogButtonBox, QTabWidget, QGridLayout, QFileDialog,
    QLineEdit, QComboBox, QSpinBox, QDateEdit, QTextEdit,
    QWizard, QWizardPage, QListWidget, QListWidgetItem, QSplitter,
    QFormLayout, QRadioButton, QButtonGroup
)
from PyQt5.QtCore import Qt, QDate, QTimer, QSettings
from PyQt5.QtGui import QFont, QColor, QTextCharFormat, QIcon

# 应用信息
APP_NAME = "校历助手"
APP_KEY = "CalendarAssistant"
APP_VERSION = "2025-2026"
APP_DEVELOPER = "肖诗顺"
APP_LICENSE = "开源软件，供教师和学生免费使用"
NEW_VERSION_NOTICE = "注意：当学校公布新的校历后，本程序会发布新的学年版本"

# 获取数据存储路径
def get_data_dir():
    """获取数据存储目录"""
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(base_dir, "data")
    if not os.path.exists(data_dir):
        os.makedirs(data_dir)
    return data_dir

def get_data_file():
    """获取数据文件路径"""
    return os.path.join(get_data_dir(), "calendar_data.json")

# ==================== 默认示例数据 ====================
DEFAULT_DATA = {
    "school_name": "示例学校",
    "academic_year": "2025-2026",
    "semesters": {
        "fall": {
            "name": "秋季学期",
            "start_date": "2025-09-08",
            "end_date": "2026-01-18"
        },
        "spring": {
            "name": "春季学期", 
            "start_date": "2026-03-02",
            "end_date": "2026-07-12"
        }
    },
    "class_times": {
        "1": ["08:00", "08:50"],
        "2": ["08:55", "09:45"],
        "3": ["10:05", "10:55"],
        "4": ["11:00", "11:50"],
        "5": ["14:00", "14:50"],
        "6": ["14:55", "15:45"],
        "7": ["16:05", "16:55"],
        "8": ["17:00", "17:50"],
        "9": ["19:00", "19:50"],
        "10": ["19:55", "20:45"]
    },
    "important_dates": [
        {"date": "2025-09-05", "event": "开学上班", "category": "开学"},
        {"date": "2025-09-08", "event": "正式行课", "category": "上课"},
        {"date": "2025-10-01", "event": "国庆节", "category": "节日"},
        {"date": "2026-01-01", "event": "元旦", "category": "节日"},
        {"date": "2026-01-19", "event": "寒假开始", "category": "假期"},
        {"date": "2026-02-17", "event": "春节", "category": "节日"},
        {"date": "2026-03-02", "event": "正式行课", "category": "上课"},
        {"date": "2026-05-01", "event": "劳动节", "category": "节日"},
    ],
    "courses": [
        {
            "name": "示例课程",
            "weeks": [1, 2, 3, 4, 5],
            "weekday": 1,
            "sections": [1, 2],
            "location": "教学楼101",
            "teacher": "张老师",
            "type": "必修"
        }
    ]
}

CATEGORY_COLORS = {
    "假期": "#4CAF50",
    "开学": "#2196F3",
    "注册": "#9C27B0",
    "上课": "#FF9800",
    "节日": "#E91E63",
    "实践": "#00BCD4",
    "考试": "#F44336",
}

# ==================== 数据管理类 ====================
class DataManager:
    """数据管理器 - 负责加载、保存和管理校历数据"""
    
    def __init__(self):
        self.data = None
        self.load_data()
    
    def load_data(self):
        """加载数据，如果不存在则使用默认数据"""
        data_file = get_data_file()
        if os.path.exists(data_file):
            try:
                with open(data_file, 'r', encoding='utf-8') as f:
                    self.data = json.load(f)
            except:
                self.data = DEFAULT_DATA.copy()
        else:
            self.data = DEFAULT_DATA.copy()
    
    def save_data(self):
        """保存数据到文件"""
        data_file = get_data_file()
        with open(data_file, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)
    
    def reset_to_default(self):
        """重置为默认数据"""
        self.data = DEFAULT_DATA.copy()
        self.save_data()
    
    def get_school_name(self):
        return self.data.get("school_name", "我的学校")
    
    def get_academic_year(self):
        return self.data.get("academic_year", "2025-2026")
    
    def get_semester_dates(self, semester):
        """获取学期开始和结束日期"""
        sem_data = self.data.get("semesters", {}).get(semester, {})
        start = sem_data.get("start_date")
        end = sem_data.get("end_date")
        if start and end:
            return (
                datetime.strptime(start, "%Y-%m-%d").date(),
                datetime.strptime(end, "%Y-%m-%d").date()
            )
        return None, None
    
    def get_class_times(self):
        """获取节次时间表"""
        times = self.data.get("class_times", {})
        return {int(k): tuple(v) for k, v in times.items()}
    
    def get_important_dates(self):
        return self.data.get("important_dates", [])
    
    def get_courses(self):
        return self.data.get("courses", [])
    
    def set_school_info(self, name, year):
        self.data["school_name"] = name
        self.data["academic_year"] = year
        self.save_data()
    
    def set_semester(self, semester, name, start_date, end_date):
        if "semesters" not in self.data:
            self.data["semesters"] = {}
        self.data["semesters"][semester] = {
            "name": name,
            "start_date": start_date,
            "end_date": end_date
        }
        self.save_data()
    
    def set_important_dates(self, dates):
        self.data["important_dates"] = dates
        self.save_data()
    
    def add_important_date(self, date_str, event, category):
        if "important_dates" not in self.data:
            self.data["important_dates"] = []
        self.data["important_dates"].append({
            "date": date_str,
            "event": event,
            "category": category
        })
        self.save_data()
    
    def set_courses(self, courses):
        self.data["courses"] = courses
        self.save_data()
    
    def add_course(self, course):
        if "courses" not in self.data:
            self.data["courses"] = []
        self.data["courses"].append(course)
        self.save_data()

# 全局数据管理器
data_manager = DataManager()

# ==================== 工具函数 ====================
def get_week_number(target_date):
    """计算给定日期是第几周"""
    if isinstance(target_date, datetime):
        target_date = target_date.date()
    
    # 检查秋季学期
    fall_start, fall_end = data_manager.get_semester_dates("fall")
    if fall_start and fall_end and fall_start <= target_date <= fall_end:
        days_diff = (target_date - fall_start).days
        week_num = days_diff // 7 + 1
        return ("秋季学期", week_num)
    
    # 检查春季学期
    spring_start, spring_end = data_manager.get_semester_dates("spring")
    if spring_start and spring_end and spring_start <= target_date <= spring_end:
        days_diff = (target_date - spring_start).days
        week_num = days_diff // 7 + 1
        return ("春季学期", week_num)
    
    return (None, None)

def get_weekday_name(target_date):
    """获取星期几的中文名称"""
    weekdays = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
    return weekdays[target_date.weekday()]

def get_courses_on_date(target_date):
    """获取指定日期的课程"""
    semester, week_num = get_week_number(target_date)
    if not semester:
        return []
    
    weekday = target_date.weekday() + 1
    courses = []
    for course in data_manager.get_courses():
        if week_num in course.get("weeks", []) and weekday == course.get("weekday"):
            courses.append(course)
    
    return sorted(courses, key=lambda x: x.get("sections", [0])[0])

def get_app_path():
    if getattr(sys, 'frozen', False):
        return sys.executable
    return os.path.abspath(__file__)

def is_autostart_enabled():
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_READ
        )
        try:
            winreg.QueryValueEx(key, APP_KEY)
            return True
        except WindowsError:
            return False
        finally:
            winreg.CloseKey(key)
    except WindowsError:
        return False

def set_autostart(enable):
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        0, winreg.KEY_SET_VALUE | winreg.KEY_READ
    )
    try:
        if enable:
            app_path = get_app_path()
            winreg.SetValueEx(key, APP_KEY, 0, winreg.REG_SZ, f'"{app_path}"')
        else:
            try:
                winreg.DeleteValue(key, APP_KEY)
            except WindowsError:
                pass
    finally:
        winreg.CloseKey(key)

# ==================== 导入向导 ====================
class ImportWizard(QWizard):
    """数据导入向导"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("导入校历和课表数据")
        self.setMinimumSize(700, 550)
        self.setWizardStyle(QWizard.ModernStyle)
        
        self.addPage(WelcomePage())
        self.addPage(SchoolInfoPage())
        self.addPage(SemesterPage())
        self.addPage(ImportDatesPage())
        self.addPage(ImportCoursesPage())
        self.addPage(FinishPage())

class WelcomePage(QWizardPage):
    """欢迎页面"""
    def __init__(self):
        super().__init__()
        self.setTitle("欢迎使用数据导入向导")
        self.setSubTitle("本向导将帮助您导入校历和课表数据")
        
        layout = QVBoxLayout(self)
        
        info = QLabel(
            "您可以通过以下方式导入数据：\n\n"
            "1. 手动输入学校信息和学期日期\n"
            "2. 手动添加重要日期（开学、放假、考试等）\n"
            "3. 手动添加课程表\n"
            "4. 从Excel文件导入课表（.xlsx格式）\n\n"
            "如果您不导入数据，系统将使用默认的示例数据。\n"
            "您可以随时在设置中重新导入数据。"
        )
        info.setWordWrap(True)
        info.setFont(QFont("Microsoft YaHei", 11))
        layout.addWidget(info)
        
        layout.addStretch()

class SchoolInfoPage(QWizardPage):
    """学校信息页面"""
    def __init__(self):
        super().__init__()
        self.setTitle("学校信息")
        self.setSubTitle("请输入您的学校名称和当前学年")
        
        layout = QFormLayout(self)
        layout.setSpacing(15)
        
        self.school_name = QLineEdit()
        self.school_name.setText(data_manager.get_school_name())
        self.school_name.setFont(QFont("Microsoft YaHei", 11))
        layout.addRow("学校名称:", self.school_name)
        
        self.academic_year = QLineEdit()
        self.academic_year.setText(data_manager.get_academic_year())
        self.academic_year.setPlaceholderText("例如: 2025-2026")
        self.academic_year.setFont(QFont("Microsoft YaHei", 11))
        layout.addRow("学年:", self.academic_year)
        
        self.registerField("school_name", self.school_name)
        self.registerField("academic_year", self.academic_year)
    
    def validatePage(self):
        data_manager.set_school_info(
            self.school_name.text(),
            self.academic_year.text()
        )
        return True

class SemesterPage(QWizardPage):
    """学期设置页面"""
    def __init__(self):
        super().__init__()
        self.setTitle("学期设置")
        self.setSubTitle("请设置学期的开始和结束日期（正式行课第一天为第一周周一）")
        
        layout = QVBoxLayout(self)
        
        # 秋季学期
        fall_group = QGroupBox("秋季学期")
        fall_layout = QFormLayout()
        
        self.fall_start = QDateEdit()
        self.fall_start.setCalendarPopup(True)
        self.fall_start.setDisplayFormat("yyyy-MM-dd")
        fall_start, _ = data_manager.get_semester_dates("fall")
        if fall_start:
            self.fall_start.setDate(QDate(fall_start.year, fall_start.month, fall_start.day))
        fall_layout.addRow("开始日期:", self.fall_start)
        
        self.fall_end = QDateEdit()
        self.fall_end.setCalendarPopup(True)
        self.fall_end.setDisplayFormat("yyyy-MM-dd")
        _, fall_end = data_manager.get_semester_dates("fall")
        if fall_end:
            self.fall_end.setDate(QDate(fall_end.year, fall_end.month, fall_end.day))
        fall_layout.addRow("结束日期:", self.fall_end)
        
        fall_group.setLayout(fall_layout)
        layout.addWidget(fall_group)
        
        # 春季学期
        spring_group = QGroupBox("春季学期")
        spring_layout = QFormLayout()
        
        self.spring_start = QDateEdit()
        self.spring_start.setCalendarPopup(True)
        self.spring_start.setDisplayFormat("yyyy-MM-dd")
        spring_start, _ = data_manager.get_semester_dates("spring")
        if spring_start:
            self.spring_start.setDate(QDate(spring_start.year, spring_start.month, spring_start.day))
        spring_layout.addRow("开始日期:", self.spring_start)
        
        self.spring_end = QDateEdit()
        self.spring_end.setCalendarPopup(True)
        self.spring_end.setDisplayFormat("yyyy-MM-dd")
        _, spring_end = data_manager.get_semester_dates("spring")
        if spring_end:
            self.spring_end.setDate(QDate(spring_end.year, spring_end.month, spring_end.day))
        spring_layout.addRow("结束日期:", self.spring_end)
        
        spring_group.setLayout(spring_layout)
        layout.addWidget(spring_group)
        
        layout.addStretch()
    
    def validatePage(self):
        data_manager.set_semester(
            "fall", "秋季学期",
            self.fall_start.date().toString("yyyy-MM-dd"),
            self.fall_end.date().toString("yyyy-MM-dd")
        )
        data_manager.set_semester(
            "spring", "春季学期",
            self.spring_start.date().toString("yyyy-MM-dd"),
            self.spring_end.date().toString("yyyy-MM-dd")
        )
        return True

class ImportDatesPage(QWizardPage):
    """导入重要日期页面"""
    def __init__(self):
        super().__init__()
        self.setTitle("重要日期")
        self.setSubTitle("添加开学、放假、考试等重要日期")
        
        layout = QVBoxLayout(self)
        
        # 日期列表
        self.dates_list = QListWidget()
        self.dates_list.setFont(QFont("Microsoft YaHei", 10))
        self.refresh_dates_list()
        layout.addWidget(self.dates_list)
        
        # 添加日期区域
        add_group = QGroupBox("添加新日期")
        add_layout = QHBoxLayout()
        
        self.new_date = QDateEdit()
        self.new_date.setCalendarPopup(True)
        self.new_date.setDisplayFormat("yyyy-MM-dd")
        self.new_date.setDate(QDate.currentDate())
        add_layout.addWidget(QLabel("日期:"))
        add_layout.addWidget(self.new_date)
        
        self.new_event = QLineEdit()
        self.new_event.setPlaceholderText("事件名称")
        add_layout.addWidget(QLabel("事件:"))
        add_layout.addWidget(self.new_event)
        
        self.new_category = QComboBox()
        self.new_category.addItems(["开学", "假期", "节日", "考试", "注册", "实践", "上课"])
        add_layout.addWidget(QLabel("类别:"))
        add_layout.addWidget(self.new_category)
        
        add_btn = QPushButton("添加")
        add_btn.clicked.connect(self.add_date)
        add_layout.addWidget(add_btn)
        
        add_group.setLayout(add_layout)
        layout.addWidget(add_group)
        
        # 删除按钮
        del_btn = QPushButton("删除选中")
        del_btn.clicked.connect(self.delete_date)
        layout.addWidget(del_btn)
    
    def refresh_dates_list(self):
        self.dates_list.clear()
        for item in data_manager.get_important_dates():
            text = f"{item['date']} - {item['event']} ({item['category']})"
            self.dates_list.addItem(text)
    
    def add_date(self):
        if not self.new_event.text():
            return
        data_manager.add_important_date(
            self.new_date.date().toString("yyyy-MM-dd"),
            self.new_event.text(),
            self.new_category.currentText()
        )
        self.new_event.clear()
        self.refresh_dates_list()
    
    def delete_date(self):
        row = self.dates_list.currentRow()
        if row >= 0:
            dates = data_manager.get_important_dates()
            del dates[row]
            data_manager.set_important_dates(dates)
            self.refresh_dates_list()

class ImportCoursesPage(QWizardPage):
    """导入课程页面"""
    def __init__(self):
        super().__init__()
        self.setTitle("课程表")
        self.setSubTitle("添加您的课程或从Excel导入")
        
        layout = QVBoxLayout(self)
        
        # 导入按钮
        import_layout = QHBoxLayout()
        
        excel_btn = QPushButton("从Excel导入")
        excel_btn.clicked.connect(self.import_from_excel)
        import_layout.addWidget(excel_btn)
        
        clear_btn = QPushButton("清空课程")
        clear_btn.clicked.connect(self.clear_courses)
        import_layout.addWidget(clear_btn)
        
        import_layout.addStretch()
        layout.addLayout(import_layout)
        
        # 课程列表
        self.courses_list = QListWidget()
        self.courses_list.setFont(QFont("Microsoft YaHei", 10))
        self.refresh_courses_list()
        layout.addWidget(self.courses_list)
        
        # 添加课程区域
        add_group = QGroupBox("手动添加课程")
        add_layout = QGridLayout()
        
        add_layout.addWidget(QLabel("课程名:"), 0, 0)
        self.course_name = QLineEdit()
        add_layout.addWidget(self.course_name, 0, 1)
        
        add_layout.addWidget(QLabel("教室:"), 0, 2)
        self.course_location = QLineEdit()
        add_layout.addWidget(self.course_location, 0, 3)
        
        add_layout.addWidget(QLabel("教师:"), 1, 0)
        self.course_teacher = QLineEdit()
        add_layout.addWidget(self.course_teacher, 1, 1)
        
        add_layout.addWidget(QLabel("星期:"), 1, 2)
        self.course_weekday = QComboBox()
        self.course_weekday.addItems(["周一", "周二", "周三", "周四", "周五", "周六", "周日"])
        add_layout.addWidget(self.course_weekday, 1, 3)
        
        add_layout.addWidget(QLabel("节次(如1-2):"), 2, 0)
        self.course_sections = QLineEdit()
        self.course_sections.setPlaceholderText("1-2")
        add_layout.addWidget(self.course_sections, 2, 1)
        
        add_layout.addWidget(QLabel("周次(如1-16):"), 2, 2)
        self.course_weeks = QLineEdit()
        self.course_weeks.setPlaceholderText("1-16")
        add_layout.addWidget(self.course_weeks, 2, 3)
        
        add_btn = QPushButton("添加课程")
        add_btn.clicked.connect(self.add_course)
        add_layout.addWidget(add_btn, 3, 0, 1, 4)
        
        add_group.setLayout(add_layout)
        layout.addWidget(add_group)
        
        # 删除按钮
        del_btn = QPushButton("删除选中课程")
        del_btn.clicked.connect(self.delete_course)
        layout.addWidget(del_btn)
    
    def refresh_courses_list(self):
        self.courses_list.clear()
        weekdays = ["", "周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        for course in data_manager.get_courses():
            sections = course.get("sections", [])
            weeks = course.get("weeks", [])
            sec_str = f"{sections[0]}-{sections[-1]}节" if sections else ""
            week_str = f"第{weeks[0]}-{weeks[-1]}周" if weeks else ""
            weekday = weekdays[course.get("weekday", 0)]
            text = f"{course.get('name', '')} | {weekday} {sec_str} | {week_str} | {course.get('location', '')}"
            self.courses_list.addItem(text)
    
    def parse_range(self, text):
        """解析范围字符串如 '1-16' 或 '1,3,5'"""
        result = []
        parts = text.replace(" ", "").split(",")
        for part in parts:
            if "-" in part:
                start, end = part.split("-")
                result.extend(range(int(start), int(end) + 1))
            else:
                result.append(int(part))
        return result
    
    def add_course(self):
        if not self.course_name.text():
            return
        
        try:
            sections = self.parse_range(self.course_sections.text())
            weeks = self.parse_range(self.course_weeks.text())
        except:
            QMessageBox.warning(self, "错误", "节次或周次格式不正确")
            return
        
        course = {
            "name": self.course_name.text(),
            "location": self.course_location.text(),
            "teacher": self.course_teacher.text(),
            "weekday": self.course_weekday.currentIndex() + 1,
            "sections": sections,
            "weeks": weeks,
            "type": "课程"
        }
        data_manager.add_course(course)
        self.refresh_courses_list()
        
        # 清空输入
        self.course_name.clear()
        self.course_location.clear()
        self.course_teacher.clear()
        self.course_sections.clear()
        self.course_weeks.clear()
    
    def delete_course(self):
        row = self.courses_list.currentRow()
        if row >= 0:
            courses = data_manager.get_courses()
            del courses[row]
            data_manager.set_courses(courses)
            self.refresh_courses_list()
    
    def clear_courses(self):
        reply = QMessageBox.question(self, "确认", "确定清空所有课程？",
                                    QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            data_manager.set_courses([])
            self.refresh_courses_list()
    
    def import_from_excel(self):
        """从Excel导入课程"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择Excel文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if not file_path:
            return
        
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            
            imported = 0
            # 假设格式: 课程名, 教师, 教室, 星期, 节次, 周次
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue
                
                try:
                    weekday_map = {"周一": 1, "周二": 2, "周三": 3, "周四": 4, 
                                  "周五": 5, "周六": 6, "周日": 7,
                                  "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7}
                    
                    weekday = weekday_map.get(str(row[3]).strip(), 1)
                    sections = self.parse_range(str(row[4])) if row[4] else [1, 2]
                    weeks = self.parse_range(str(row[5])) if row[5] else list(range(1, 17))
                    
                    course = {
                        "name": str(row[0]),
                        "teacher": str(row[1]) if row[1] else "",
                        "location": str(row[2]) if row[2] else "",
                        "weekday": weekday,
                        "sections": sections,
                        "weeks": weeks,
                        "type": "导入"
                    }
                    data_manager.add_course(course)
                    imported += 1
                except Exception as e:
                    continue
            
            self.refresh_courses_list()
            QMessageBox.information(self, "导入完成", f"成功导入 {imported} 门课程")
            
        except ImportError:
            QMessageBox.warning(self, "错误", "请先安装openpyxl库:\npip install openpyxl")
        except Exception as e:
            QMessageBox.warning(self, "导入失败", f"导入失败: {str(e)}")

class FinishPage(QWizardPage):
    """完成页面"""
    def __init__(self):
        super().__init__()
        self.setTitle("完成")
        self.setSubTitle("数据导入完成")
        
        layout = QVBoxLayout(self)
        
        info = QLabel(
            "数据已保存！\n\n"
            "您可以随时通过以下方式修改数据：\n"
            "- 点击主界面的「数据管理」按钮\n"
            "- 在设置中选择「重新导入数据」\n\n"
            "点击「完成」开始使用应用。"
        )
        info.setWordWrap(True)
        info.setFont(QFont("Microsoft YaHei", 11))
        layout.addWidget(info)
        
        layout.addStretch()

# ==================== 设置对话框 ====================
class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setFixedSize(450, 350)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        title = QLabel("应用设置")
        title.setFont(QFont("Microsoft YaHei", 14, QFont.Bold))
        layout.addWidget(title)
        
        self.autostart_checkbox = QCheckBox("开机自动启动")
        self.autostart_checkbox.setFont(QFont("Microsoft YaHei", 11))
        self.autostart_checkbox.setChecked(is_autostart_enabled())
        layout.addWidget(self.autostart_checkbox)
        
        self.minimize_to_tray_checkbox = QCheckBox("关闭时最小化到系统托盘")
        self.minimize_to_tray_checkbox.setFont(QFont("Microsoft YaHei", 11))
        settings = QSettings(APP_KEY, APP_NAME)
        self.minimize_to_tray_checkbox.setChecked(
            settings.value("minimize_to_tray", True, type=bool)
        )
        layout.addWidget(self.minimize_to_tray_checkbox)
        
        self.alarm_checkbox = QCheckBox("上课前30分钟闹钟提醒")
        self.alarm_checkbox.setFont(QFont("Microsoft YaHei", 11))
        self.alarm_checkbox.setChecked(
            settings.value("alarm_enabled", True, type=bool)
        )
        layout.addWidget(self.alarm_checkbox)
        
        self.day_before_checkbox = QCheckBox("上课前一天弹窗提醒")
        self.day_before_checkbox.setFont(QFont("Microsoft YaHei", 11))
        self.day_before_checkbox.setChecked(
            settings.value("day_before_reminder", True, type=bool)
        )
        layout.addWidget(self.day_before_checkbox)
        
        layout.addSpacing(10)
        
        # 数据管理按钮
        data_btn = QPushButton("重新导入数据...")
        data_btn.clicked.connect(self.open_import_wizard)
        layout.addWidget(data_btn)
        
        reset_btn = QPushButton("重置为默认数据")
        reset_btn.clicked.connect(self.reset_data)
        layout.addWidget(reset_btn)
        
        # 版本和开发者信息
        info_label = QLabel(
            f"{APP_VERSION}学年版本\\n\\n"
            f"开发者：{APP_DEVELOPER}\\n"
            f"{APP_LICENSE}\\n\\n"
            f"{NEW_VERSION_NOTICE}"
        )
        info_label.setFont(QFont("Microsoft YaHei", 10))
        info_label.setStyleSheet("color: #666;")
        layout.addWidget(info_label)
        
        layout.addStretch()
        
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.save_settings)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def save_settings(self):
        set_autostart(self.autostart_checkbox.isChecked())
        
        settings = QSettings(APP_KEY, APP_NAME)
        settings.setValue("minimize_to_tray", self.minimize_to_tray_checkbox.isChecked())
        settings.setValue("alarm_enabled", self.alarm_checkbox.isChecked())
        settings.setValue("day_before_reminder", self.day_before_checkbox.isChecked())
        
        self.accept()
    
    def open_import_wizard(self):
        wizard = ImportWizard(self)
        if wizard.exec_() == QWizard.Accepted:
            QMessageBox.information(self, "提示", "数据已更新，请重启应用以应用更改。")
    
    def reset_data(self):
        reply = QMessageBox.question(self, "确认", 
                                    "确定重置为默认数据？\n这将删除您导入的所有数据。",
                                    QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            data_manager.reset_to_default()
            QMessageBox.information(self, "提示", "已重置为默认数据，请重启应用。")

# ==================== 主窗口 ====================
class CalendarApp(QMainWindow):
    def __init__(self):
        super().__init__()
        school = data_manager.get_school_name()
        year = data_manager.get_academic_year()
        self.setWindowTitle(f"{school}校历 {year}学年 - 2025-2026版")
        self.setMinimumSize(1100, 750)
        
        self.settings = QSettings(APP_KEY, APP_NAME)
        self.reminded_classes = set()
        self.reminded_day_before = set()
        
        self.setup_tray_icon()
        self.setup_ui()
        self.setup_timer()
        self.setup_alarm_timer()
        
        # 首次运行显示导入向导
        if not self.settings.value("first_run_done", False, type=bool):
            self.show_first_run_dialog()
        
        self.check_day_before_reminder()
    
    def show_first_run_dialog(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("欢迎使用校历助手")
        msg.setText("是否现在导入您的校历和课表数据？\n\n"
                   f"当前为{APP_VERSION}学年版本\n\n"
                   f"{NEW_VERSION_NOTICE}\n\n"
                   "您也可以稍后在设置中导入。")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.Yes)
        
        if msg.exec_() == QMessageBox.Yes:
            wizard = ImportWizard(self)
            wizard.exec_()
        
        # 询问开机自启动
        msg2 = QMessageBox(self)
        msg2.setWindowTitle("开机自启动")
        msg2.setText("是否设置开机自动启动？\n\n"
                   f"开发者：{APP_DEVELOPER}\n"
                   f"{APP_LICENSE}")
        msg2.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        if msg2.exec_() == QMessageBox.Yes:
            set_autostart(True)
        
        self.settings.setValue("first_run_done", True)
    
    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        icon = self.style().standardIcon(self.style().SP_ComputerIcon)
        self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip(f"{data_manager.get_school_name()}校历")
        
        tray_menu = QMenu()
        
        show_action = QAction("显示主窗口", self)
        show_action.triggered.connect(self.show_window)
        tray_menu.addAction(show_action)
        
        self.week_action = QAction("", self)
        self.week_action.setEnabled(False)
        tray_menu.addAction(self.week_action)
        self.update_tray_week_info()
        
        self.today_course_action = QAction("", self)
        self.today_course_action.setEnabled(False)
        tray_menu.addAction(self.today_course_action)
        self.update_today_course_info()
        
        tray_menu.addSeparator()
        
        data_action = QAction("数据管理", self)
        data_action.triggered.connect(self.open_data_manager)
        tray_menu.addAction(data_action)
        
        settings_action = QAction("设置", self)
        settings_action.triggered.connect(self.show_settings)
        tray_menu.addAction(settings_action)
        
        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        tray_menu.addAction(about_action)
        
        tray_menu.addSeparator()
        
        quit_action = QAction("退出", self)
        quit_action.triggered.connect(self.quit_app)
        tray_menu.addAction(quit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()
    
    def update_tray_week_info(self):
        today = date.today()
        semester, week_num = get_week_number(today)
        if semester and week_num:
            self.week_action.setText(f"当前: {semester} 第{week_num}周")
        else:
            self.week_action.setText("当前: 假期")
    
    def update_today_course_info(self):
        courses = get_courses_on_date(date.today())
        if courses:
            self.today_course_action.setText(f"今日课程: {len(courses)}节")
        else:
            self.today_course_action.setText("今日无课")
    
    def tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()
    
    def show_window(self):
        self.showNormal()
        self.activateWindow()
        self.raise_()
    
    def show_settings(self):
        dialog = SettingsDialog(self)
        dialog.exec_()
        self.refresh_display()
    
    def show_about(self):
        QMessageBox.about(self, "关于校历助手", 
                         f"校历助手 - {APP_VERSION}学年版本\\n\\n"
                         f"开发者：{APP_DEVELOPER}\\n"
                         f"{APP_LICENSE}\\n\\n"
                         f"{NEW_VERSION_NOTICE}")
    
    def open_data_manager(self):
        wizard = ImportWizard(self)
        if wizard.exec_() == QWizard.Accepted:
            self.refresh_display()
    
    def refresh_display(self):
        """刷新显示"""
        school = data_manager.get_school_name()
        year = data_manager.get_academic_year()
        self.setWindowTitle(f"{school}校历 {year}学年 - 2025-2026版")
        self.update_current_date()
        self.update_today_courses_display()
        self.populate_week_table()
        self.populate_events_table()
        self.highlight_important_dates()
        self.highlight_course_dates()
    
    def quit_app(self):
        self.tray_icon.hide()
        QApplication.quit()
    
    def closeEvent(self, event):
        minimize_to_tray = self.settings.value("minimize_to_tray", True, type=bool)
        
        if minimize_to_tray:
            event.ignore()
            self.hide()
            self.tray_icon.showMessage(
                "校历助手",
                "程序已最小化到系统托盘，双击图标可恢复窗口",
                QSystemTrayIcon.Information,
                2000
            )
        else:
            self.quit_app()
    
    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setSpacing(20)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 左侧面板
        left_panel = QVBoxLayout()
        
        self.title_label = QLabel(data_manager.get_school_name())
        self.title_label.setFont(QFont("Microsoft YaHei", 24, QFont.Bold))
        self.title_label.setStyleSheet("color: #1565C0;")
        self.title_label.setAlignment(Qt.AlignCenter)
        left_panel.addWidget(self.title_label)
        
        self.subtitle_label = QLabel(f"{data_manager.get_academic_year()}学年校历")
        self.subtitle_label.setFont(QFont("Microsoft YaHei", 14))
        self.subtitle_label.setStyleSheet("color: #666;")
        self.subtitle_label.setAlignment(Qt.AlignCenter)
        left_panel.addWidget(self.subtitle_label)
        
        left_panel.addSpacing(10)
        
        self.date_label = QLabel()
        self.date_label.setFont(QFont("Microsoft YaHei", 12))
        self.date_label.setAlignment(Qt.AlignCenter)
        self.date_label.setStyleSheet("""
            QLabel {
                background-color: #E3F2FD;
                padding: 15px;
                border-radius: 8px;
                color: #1565C0;
            }
        """)
        left_panel.addWidget(self.date_label)
        self.update_current_date()
        
        left_panel.addSpacing(10)
        
        self.calendar = QCalendarWidget()
        self.calendar.setGridVisible(True)
        self.calendar.setFont(QFont("Microsoft YaHei", 10))
        self.calendar.setStyleSheet("""
            QCalendarWidget {
                background-color: white;
                border: 1px solid #ddd;
                border-radius: 8px;
            }
            QCalendarWidget QToolButton {
                color: #333;
                background-color: #f5f5f5;
                border: none;
                border-radius: 4px;
                padding: 5px;
            }
        """)
        self.highlight_important_dates()
        self.highlight_course_dates()
        self.calendar.clicked.connect(self.on_date_clicked)
        left_panel.addWidget(self.calendar)
        
        bottom_layout = QHBoxLayout()
        
        legend_group = QGroupBox("图例")
        legend_group.setFont(QFont("Microsoft YaHei", 9))
        legend_layout = QHBoxLayout()
        legend_items = [("假期", "#4CAF50"), ("节日", "#E91E63"), ("课程", "#2196F3")]
        for name, color in legend_items:
            legend_item = QLabel(f"● {name}")
            legend_item.setStyleSheet(f"color: {color}; font-weight: bold;")
            legend_layout.addWidget(legend_item)
        legend_group.setLayout(legend_layout)
        bottom_layout.addWidget(legend_group)
        
        data_btn = QPushButton("数据管理")
        data_btn.setFont(QFont("Microsoft YaHei", 10))
        data_btn.clicked.connect(self.open_data_manager)
        bottom_layout.addWidget(data_btn)
        
        settings_btn = QPushButton("设置")
        settings_btn.setFont(QFont("Microsoft YaHei", 10))
        settings_btn.clicked.connect(self.show_settings)
        bottom_layout.addWidget(settings_btn)
        
        left_panel.addLayout(bottom_layout)
        
        main_layout.addLayout(left_panel, 1)
        
        # 右侧面板
        right_panel = QVBoxLayout()
        
        self.tab_widget = QTabWidget()
        self.tab_widget.setFont(QFont("Microsoft YaHei", 10))
        
        # Tab 1: 今日课程
        today_tab = QWidget()
        today_layout = QVBoxLayout(today_tab)
        self.today_course_label = QLabel()
        self.today_course_label.setFont(QFont("Microsoft YaHei", 11))
        self.today_course_label.setWordWrap(True)
        self.today_course_label.setStyleSheet("""
            QLabel {
                background-color: white;
                padding: 15px;
                border-radius: 8px;
                border: 1px solid #ddd;
            }
        """)
        today_layout.addWidget(self.today_course_label)
        self.update_today_courses_display()
        self.tab_widget.addTab(today_tab, "今日课程")
        
        # Tab 2: 本周课表
        week_tab = QWidget()
        week_layout = QVBoxLayout(week_tab)
        self.week_table = QTableWidget()
        self.week_table.setColumnCount(8)
        self.week_table.setHorizontalHeaderLabels(["节次", "周一", "周二", "周三", "周四", "周五", "周六", "周日"])
        self.week_table.setRowCount(10)
        self.week_table.verticalHeader().setVisible(False)
        self.week_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.week_table.setFont(QFont("Microsoft YaHei", 9))
        self.week_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 1px solid #ddd;
                gridline-color: #eee;
            }
            QHeaderView::section {
                background-color: #1565C0;
                color: white;
                padding: 8px;
                border: none;
            }
        """)
        self.week_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.populate_week_table()
        week_layout.addWidget(self.week_table)
        self.tab_widget.addTab(week_tab, "本周课表")
        
        # Tab 3: 重要日期
        events_tab = QWidget()
        events_layout = QVBoxLayout(events_tab)
        self.events_table = QTableWidget()
        self.events_table.setColumnCount(4)
        self.events_table.setHorizontalHeaderLabels(["日期", "周次", "事件", "类别"])
        self.events_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.events_table.setFont(QFont("Microsoft YaHei", 10))
        self.events_table.setAlternatingRowColors(True)
        self.events_table.setStyleSheet("""
            QTableWidget {
                background-color: white;
                border: 1px solid #ddd;
            }
            QHeaderView::section {
                background-color: #1565C0;
                color: white;
                padding: 8px;
                border: none;
            }
        """)
        self.events_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.populate_events_table()
        events_layout.addWidget(self.events_table)
        self.tab_widget.addTab(events_tab, "重要日期")
        
        right_panel.addWidget(self.tab_widget)
        
        self.selected_date_label = QLabel("点击日历查看当日详情")
        self.selected_date_label.setFont(QFont("Microsoft YaHei", 11))
        self.selected_date_label.setWordWrap(True)
        self.selected_date_label.setMinimumHeight(120)
        self.selected_date_label.setStyleSheet("""
            QLabel {
                background-color: #FFF8E1;
                padding: 15px;
                border-radius: 8px;
                border-left: 4px solid #FFC107;
            }
        """)
        right_panel.addWidget(self.selected_date_label)
        
        main_layout.addLayout(right_panel, 1)
        
        self.setStyleSheet("""
            QMainWindow { background-color: #f5f5f5; }
            QGroupBox {
                border: 1px solid #ddd;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: white;
            }
            QTabWidget::pane {
                border: 1px solid #ddd;
                border-radius: 8px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #f5f5f5;
                padding: 8px 16px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #1565C0;
                color: white;
            }
        """)
    
    def setup_timer(self):
        timer = QTimer(self)
        timer.timeout.connect(self.update_current_date)
        timer.timeout.connect(self.update_tray_week_info)
        timer.timeout.connect(self.update_today_course_info)
        timer.timeout.connect(self.update_today_courses_display)
        timer.start(60000)
    
    def setup_alarm_timer(self):
        self.alarm_timer = QTimer(self)
        self.alarm_timer.timeout.connect(self.check_class_alarm)
        self.alarm_timer.start(60000)
        QTimer.singleShot(1000, self.check_class_alarm)
    
    def check_class_alarm(self):
        if not self.settings.value("alarm_enabled", True, type=bool):
            return
        
        now = datetime.now()
        today = now.date()
        current_time = now.time()
        
        courses = get_courses_on_date(today)
        class_times = data_manager.get_class_times()
        
        for course in courses:
            first_section = course.get("sections", [1])[0]
            if first_section not in class_times:
                continue
            
            class_start_str = class_times[first_section][0]
            class_start = datetime.strptime(class_start_str, "%H:%M").time()
            
            class_datetime = datetime.combine(today, class_start)
            remind_datetime = class_datetime - timedelta(minutes=30)
            remind_time = remind_datetime.time()
            
            course_id = f"{today}_{course.get('name')}_{first_section}"
            
            current_minutes = current_time.hour * 60 + current_time.minute
            remind_minutes = remind_time.hour * 60 + remind_time.minute
            
            if abs(current_minutes - remind_minutes) <= 1 and course_id not in self.reminded_classes:
                self.reminded_classes.add(course_id)
                self.show_class_alarm(course, class_start_str)
    
    def show_class_alarm(self, course, start_time):
        try:
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)
        except:
            pass
        
        sections = course.get("sections", [])
        sections_str = f"第{sections[0]}-{sections[-1]}节" if sections else ""
        
        self.tray_icon.showMessage(
            "上课提醒",
            f"30分钟后上课！\n\n"
            f"课程: {course.get('name')}\n"
            f"时间: {start_time} ({sections_str})\n"
            f"地点: {course.get('location')}",
            QSystemTrayIcon.Warning,
            10000
        )
        
        self.show_window()
        msg = QMessageBox(self)
        msg.setWindowTitle("上课提醒")
        msg.setIcon(QMessageBox.Warning)
        msg.setText(f"30分钟后上课！\n\n"
                   f"课程: {course.get('name')}\n"
                   f"时间: {start_time} ({sections_str})\n"
                   f"地点: {course.get('location')}\n"
                   f"教师: {course.get('teacher')}")
        msg.exec_()
    
    def check_day_before_reminder(self):
        if not self.settings.value("day_before_reminder", True, type=bool):
            return
        
        tomorrow = date.today() + timedelta(days=1)
        courses = get_courses_on_date(tomorrow)
        
        if not courses:
            return
        
        reminder_id = f"day_before_{tomorrow}"
        if reminder_id in self.reminded_day_before:
            return
        
        last_reminder = self.settings.value("last_day_before_reminder", "")
        if last_reminder == str(tomorrow):
            return
        
        self.reminded_day_before.add(reminder_id)
        self.settings.setValue("last_day_before_reminder", str(tomorrow))
        
        class_times = data_manager.get_class_times()
        course_list = ""
        for c in courses:
            sections = c.get("sections", [])
            sections_str = f"第{sections[0]}-{sections[-1]}节" if sections else ""
            start_time = class_times.get(sections[0], ["?"])[0] if sections else "?"
            course_list += f"\n  - {c.get('name')} ({start_time}, {sections_str})"
        
        semester, week_num = get_week_number(tomorrow)
        weekday = get_weekday_name(tomorrow)
        
        self.tray_icon.showMessage(
            "明日课程提醒",
            f"明天 ({tomorrow.strftime('%m月%d日')} {weekday}) 有 {len(courses)} 节课"
            f"{course_list}",
            QSystemTrayIcon.Information,
            15000
        )
    
    def update_current_date(self):
        now = datetime.now()
        today_date = date.today()
        weekday = get_weekday_name(today_date)
        today_str = now.strftime("%Y年%m月%d日") + f" {weekday}"
        
        semester, week_num = get_week_number(today_date)
        
        text = f"今天: {today_str}\n"
        
        if semester and week_num:
            text += f"当前: {semester} 第{week_num}周\n"
        else:
            spring_start, _ = data_manager.get_semester_dates("spring")
            if spring_start:
                days_to_spring = (spring_start - today_date).days
                if days_to_spring > 0:
                    text += f"距离春季学期开学还有 {days_to_spring} 天\n"
        
        courses = get_courses_on_date(today_date)
        text += f"今日课程: {len(courses)}节" if courses else "今日无课"
        
        self.date_label.setText(text)
    
    def update_today_courses_display(self):
        today = date.today()
        courses = get_courses_on_date(today)
        semester, week_num = get_week_number(today)
        class_times = data_manager.get_class_times()
        
        if not semester:
            spring_start, _ = data_manager.get_semester_dates("spring")
            if spring_start:
                days = (spring_start - today).days
                text = f"<b style='color:#E91E63;'>当前为假期</b><br><br>"
                text += f"<span>距离开学还有 <b>{days}</b> 天</span>"
            else:
                text = "<b>当前为假期</b>"
            self.today_course_label.setText(text)
            return
        
        if not courses:
            text = f"<b>第{week_num}周 {get_weekday_name(today)}</b><br><br>"
            text += "<span style='color:#666;'>今日没有课程安排</span>"
            self.today_course_label.setText(text)
            return
        
        text = f"<b>今日课程 (第{week_num}周 {get_weekday_name(today)})</b><br><br>"
        
        for course in courses:
            sections = course.get("sections", [])
            start_time = class_times.get(sections[0], ["?", "?"])[0] if sections else "?"
            end_time = class_times.get(sections[-1], ["?", "?"])[1] if sections else "?"
            sections_str = f"第{sections[0]}-{sections[-1]}节" if sections else ""
            
            text += f"<div style='margin-bottom:10px; padding:10px; background:#E3F2FD; border-radius:5px;'>"
            text += f"<b style='color:#1565C0;'>{course.get('name')}</b><br>"
            text += f"<span style='color:#666;'>时间: {start_time}-{end_time} ({sections_str})</span><br>"
            text += f"<span style='color:#666;'>地点: {course.get('location')}</span>"
            text += "</div>"
        
        self.today_course_label.setText(text)
    
    def populate_week_table(self):
        today = date.today()
        semester, week_num = get_week_number(today)
        class_times = data_manager.get_class_times()
        
        for row in range(10):
            time_str = class_times.get(row + 1, ["?", "?"])[0]
            item = QTableWidgetItem(f"{row+1}节\n{time_str}")
            item.setTextAlignment(Qt.AlignCenter)
            self.week_table.setItem(row, 0, item)
        
        for row in range(10):
            for col in range(1, 8):
                self.week_table.setItem(row, col, QTableWidgetItem(""))
        
        display_week = week_num if semester else 1
        
        for course in data_manager.get_courses():
            if display_week not in course.get("weeks", []):
                continue
            
            weekday = course.get("weekday", 1)
            for section in course.get("sections", []):
                if section <= 10:
                    name = course.get("name", "")[:6]
                    loc = course.get("location", "")
                    item = QTableWidgetItem(f"{name}\n{loc}")
                    item.setTextAlignment(Qt.AlignCenter)
                    item.setBackground(QColor("#E3F2FD"))
                    item.setToolTip(f"{course.get('name')}\n{loc}\n{course.get('teacher')}")
                    self.week_table.setItem(section - 1, weekday, item)
        
        self.week_table.resizeRowsToContents()
    
    def highlight_important_dates(self):
        # 清除旧的高亮
        self.calendar.setDateTextFormat(QDate(), QTextCharFormat())
        
        highlight_format = QTextCharFormat()
        highlight_format.setBackground(QColor("#FFEB3B"))
        highlight_format.setForeground(QColor("#333"))
        
        for item in data_manager.get_important_dates():
            try:
                event_date = datetime.strptime(item["date"], "%Y-%m-%d")
                qdate = QDate(event_date.year, event_date.month, event_date.day)
                self.calendar.setDateTextFormat(qdate, highlight_format)
            except:
                continue
    
    def highlight_course_dates(self):
        course_format = QTextCharFormat()
        course_format.setBackground(QColor("#BBDEFB"))
        course_format.setForeground(QColor("#1565C0"))
        
        spring_start, spring_end = data_manager.get_semester_dates("spring")
        fall_start, fall_end = data_manager.get_semester_dates("fall")
        
        for start, end in [(fall_start, fall_end), (spring_start, spring_end)]:
            if not start or not end:
                continue
            current = start
            while current <= end:
                courses = get_courses_on_date(current)
                if courses:
                    qdate = QDate(current.year, current.month, current.day)
                    existing = self.calendar.dateTextFormat(qdate)
                    if existing.background().color() != QColor("#FFEB3B"):
                        self.calendar.setDateTextFormat(qdate, course_format)
                current += timedelta(days=1)
    
    def populate_events_table(self):
        dates = data_manager.get_important_dates()
        sorted_events = sorted(dates, key=lambda x: x.get("date", ""))
        
        self.events_table.setRowCount(len(sorted_events))
        
        for row, item in enumerate(sorted_events):
            try:
                event_date = datetime.strptime(item["date"], "%Y-%m-%d").date()
            except:
                continue
            
            date_item = QTableWidgetItem(item["date"])
            date_item.setTextAlignment(Qt.AlignCenter)
            self.events_table.setItem(row, 0, date_item)
            
            semester, week_num = get_week_number(event_date)
            week_text = f"第{week_num}周" if semester and week_num else "-"
            week_item = QTableWidgetItem(week_text)
            week_item.setTextAlignment(Qt.AlignCenter)
            self.events_table.setItem(row, 1, week_item)
            
            event_item = QTableWidgetItem(item.get("event", ""))
            self.events_table.setItem(row, 2, event_item)
            
            category = item.get("category", "")
            category_item = QTableWidgetItem(category)
            category_item.setTextAlignment(Qt.AlignCenter)
            color = CATEGORY_COLORS.get(category, "#333")
            category_item.setForeground(QColor(color))
            self.events_table.setItem(row, 3, category_item)
    
    def on_date_clicked(self, qdate):
        selected = date(qdate.year(), qdate.month(), qdate.day())
        selected_str = selected.strftime("%Y-%m-%d")
        weekday = get_weekday_name(selected)
        
        semester, week_num = get_week_number(selected)
        week_info = f" | {semester} 第{week_num}周 {weekday}" if semester and week_num else f" {weekday}"
        
        text = f"<b>{selected.strftime('%Y年%m月%d日')}{week_info}</b><br>"
        text += "━" * 25 + "<br>"
        
        events_on_date = [item for item in data_manager.get_important_dates() 
                         if item.get("date") == selected_str]
        if events_on_date:
            for event in events_on_date:
                color = CATEGORY_COLORS.get(event.get("category"), "#333")
                text += f"<span style='color:{color};'>● {event.get('event')} ({event.get('category')})</span><br>"
        
        courses = get_courses_on_date(selected)
        class_times = data_manager.get_class_times()
        if courses:
            text += "<br><b>课程安排:</b><br>"
            for course in courses:
                sections = course.get("sections", [])
                start_time = class_times.get(sections[0], ["?"])[0] if sections else "?"
                sections_str = f"第{sections[0]}-{sections[-1]}节" if sections else ""
                text += f"<span style='color:#1565C0;'>● {course.get('name')}</span><br>"
                text += f"  <span style='color:#666;'>{start_time} {sections_str} | {course.get('location')}</span><br>"
        
        if not events_on_date and not courses:
            text += "<span style='color:#666;'>无安排</span>" if semester else "<span style='color:#666;'>假期</span>"
        
        self.selected_date_label.setText(text)
        
        style = """
            QLabel {
                background-color: %s;
                padding: 15px;
                border-radius: 8px;
                border-left: 4px solid %s;
            }
        """
        if events_on_date or courses:
            self.selected_date_label.setStyleSheet(style % ("#E8F5E9", "#4CAF50"))
        else:
            self.selected_date_label.setStyleSheet(style % ("#FFF8E1", "#FFC107"))

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setQuitOnLastWindowClosed(False)
    
    font = QFont("Microsoft YaHei", 9)
    app.setFont(font)
    
    window = CalendarApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
