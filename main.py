import sys
import os
import datetime
import json
import logging
import random
import requests
import xml.etree.ElementTree as ET
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QCalendarWidget,
    QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit,
    QPushButton, QCheckBox, QMessageBox, QDialog,
    QDialogButtonBox, QGroupBox, QComboBox, QMenu, QSpinBox, QLabel
)
from PySide6.QtCore import QDate, Qt, QRect, QPoint, QSize
from PySide6.QtGui import QPainter, QColor, QFont, QPalette, QPixmap, QScreen

# 로그 설정 (필요한 경우)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='debug_log.txt',
    filemode='w'
)
logging.debug("프로그램 시작")

# ----- Capture Settings Dialog -----
class CaptureSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("캡쳐 설정")
        layout = QVBoxLayout(self)
        
        # 오프셋 설정
        offset_layout = QHBoxLayout()
        offset_layout.addWidget(QLabel("X 오프셋:"))
        self.xOffsetSpin = QSpinBox()
        self.xOffsetSpin.setRange(-500, 500)
        self.xOffsetSpin.setValue(0)
        offset_layout.addWidget(self.xOffsetSpin)
        offset_layout.addWidget(QLabel("Y 오프셋:"))
        self.yOffsetSpin = QSpinBox()
        self.yOffsetSpin.setRange(-500, 500)
        self.yOffsetSpin.setValue(0)
        offset_layout.addWidget(self.yOffsetSpin)
        layout.addLayout(offset_layout)
        
        # 크기 설정
        size_layout = QHBoxLayout()
        size_layout.addWidget(QLabel("너비:"))
        self.widthSpin = QSpinBox()
        self.widthSpin.setRange(100, 3000)
        self.widthSpin.setValue(1200)
        size_layout.addWidget(self.widthSpin)
        size_layout.addWidget(QLabel("높이:"))
        self.heightSpin = QSpinBox()
        self.heightSpin.setRange(100, 3000)
        self.heightSpin.setValue(900)
        size_layout.addWidget(self.heightSpin)
        layout.addLayout(size_layout)
        
        # 확인/취소 버튼
        btn_layout = QHBoxLayout()
        ok_btn = QPushButton("적용")
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("취소")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(ok_btn)
        btn_layout.addWidget(cancel_btn)
        layout.addLayout(btn_layout)
    
    def getSettings(self):
        return (self.xOffsetSpin.value(), self.yOffsetSpin.value(),
                self.widthSpin.value(), self.heightSpin.value())

# ----- Shift Change Dialog -----
class ShiftChangeDialog(QDialog):
    def __init__(self, entries, parent=None):
        super().__init__(parent)
        self.setWindowTitle("근무조 변경")
        self.layout = QVBoxLayout(self)
        self.checkboxes = []
        for i, entry in enumerate(entries):
            cb = QCheckBox(f"{entry['shift']} {entry['name']}")
            cb.setChecked(True)
            self.checkboxes.append(cb)
            self.layout.addWidget(cb)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.layout.addWidget(buttons)
    def getSelectedIndices(self):
        selected = []
        for i, cb in enumerate(self.checkboxes):
            if cb.isChecked():
                selected.append(i)
        return selected

# ----- Attendance Dialog (결근/지각 상태 변경) -----
class AttendanceDialog(QDialog):
    def __init__(self, entries, parent=None):
        super().__init__(parent)
        self.setWindowTitle("출석 상태 변경")
        self.layout = QVBoxLayout(self)
        self.entries = entries
        self.widgets = []
        for entry in entries:
            hbox = QHBoxLayout()
            text = f"{entry['shift']} {entry['name']}"
            absent_cb = QCheckBox("결근")
            absent_cb.setChecked(entry.get("absent", False))
            tardy_cb = QCheckBox("지각")
            tardy_cb.setChecked(entry.get("tardy", False))
            text_edit = QLineEdit(text)
            text_edit.setReadOnly(True)
            hbox.addWidget(text_edit)
            hbox.addWidget(absent_cb)
            hbox.addWidget(tardy_cb)
            self.layout.addLayout(hbox)
            self.widgets.append((absent_cb, tardy_cb))
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.layout.addWidget(buttons)
    def accept(self):
        for i, (absent_cb, tardy_cb) in enumerate(self.widgets):
            self.entries[i]["absent"] = absent_cb.isChecked()
            self.entries[i]["tardy"] = tardy_cb.isChecked()
        super().accept()

# ----- Add Worker Dialog (해당일자 근무자 추가) -----
class AddWorkerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("해당일자 근무자 추가")
        layout = QFormLayout(self)
        self.name_edit = QLineEdit()
        self.shift_combo = QComboBox()
        self.shift_combo.addItems(["오전", "오후"])
        layout.addRow("이름:", self.name_edit)
        layout.addRow("근무조:", self.shift_combo)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    def getValues(self):
        name = self.name_edit.text().strip()
        shift_text = self.shift_combo.currentText()
        shift = "AM" if shift_text == "오전" else "PM"
        return name, shift

# ----- Delete Worker Dialog (해당일자 근무자 삭제) -----
class DeleteWorkerDialog(QDialog):
    def __init__(self, entries, parent=None):
        super().__init__(parent)
        self.setWindowTitle("해당일자 근무자 삭제")
        self.layout = QVBoxLayout(self)
        self.checkboxes = []
        for entry in entries:
            cb = QCheckBox(f"{entry['shift']} {entry['name']}")
            cb.setChecked(False)
            self.checkboxes.append(cb)
            self.layout.addWidget(cb)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        self.layout.addWidget(buttons)
    def getSelectedIndices(self):
        selected = []
        for i, cb in enumerate(self.checkboxes):
            if cb.isChecked():
                selected.append(i)
        return selected

# ----- Custom Day-Of-Week Header Widget -----
class DayOfWeekHeader(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.days = ["일", "월", "화", "수", "목", "금", "토"]
        self.setMinimumHeight(30)
    def resizeEvent(self, event):
        self.update()
        super().resizeEvent(event)
    def paintEvent(self, event):
        painter = QPainter(self)
        width = self.width() / 7
        height = self.height()
        font = QFont("Arial", int(height * 0.5))
        painter.setFont(font)
        for i, day in enumerate(self.days):
            rect = QRect(int(i * width), 0, int(width), height)
            painter.drawText(rect, Qt.AlignCenter, day)
        painter.end()

# ----- Schedule Manager -----
class ScheduleManager:
    def __init__(self, excel_file="schedule.xlsx"):
        self.excel_file = excel_file
        self.schedule = {}
        self.ensure_excel_file()
        self.load_schedule()
    
    def ensure_excel_file(self):
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame(columns=["date", "name", "shift", "absent", "tardy"])
            df.to_excel(self.excel_file, index=False)
            logging.debug("엑셀 파일 생성")
        else:
            try:
                df = pd.read_excel(self.excel_file)
                if not {"date", "name", "shift", "absent", "tardy"}.issubset(set(df.columns)):
                    df = pd.DataFrame(columns=["date", "name", "shift", "absent", "tardy"])
                    df.to_excel(self.excel_file, index=False)
                    logging.debug("엑셀 파일 재생성")
            except Exception as e:
                logging.error("엑셀 파일 확인 오류: %s", e)
                df = pd.DataFrame(columns=["date", "name", "shift", "absent", "tardy"])
                df.to_excel(self.excel_file, index=False)
    
    def load_schedule(self):
        try:
            df = pd.read_excel(self.excel_file, converters={"date": lambda x: x})
            logging.debug("엑셀 로드 성공: %s", self.excel_file)
            if "date" not in df.columns:
                self.schedule = {}
                logging.error("엑셀에 'date' 컬럼 없음")
                return
            for idx, row in df.iterrows():
                try:
                    date_val = pd.to_datetime(row['date'])
                    date_str = date_val.strftime("%Y-%m-%d")
                except Exception as e:
                    date_str = str(row['date']).strip().split(" ")[0]
                date = QDate.fromString(date_str, "yyyy-MM-dd")
                if date.isValid():
                    key = date.toString(Qt.ISODate)
                    if key not in self.schedule:
                        self.schedule[key] = []
                    absent = bool(row['absent']) if "absent" in df.columns else False
                    tardy = bool(row['tardy']) if "tardy" in df.columns else False
                    self.schedule[key].append({
                        "name": str(row['name']),
                        "shift": row['shift'],
                        "absent": absent,
                        "tardy": tardy
                    })
                else:
                    logging.error("QDate 변환 실패: %s", date_str)
            logging.debug("최종 스케줄: %s", self.schedule)
        except Exception as e:
            logging.error("엑셀 읽기 오류: %s", e)
            self.schedule = {}
    
    def save_schedule(self):
        data = []
        for date_str, entries in self.schedule.items():
            for entry in entries:
                data.append({
                    "date": date_str,
                    "name": entry["name"],
                    "shift": entry["shift"],
                    "absent": entry.get("absent", False),
                    "tardy": entry.get("tardy", False)
                })
        df = pd.DataFrame(data, columns=["date", "name", "shift", "absent", "tardy"])
        df.to_excel(self.excel_file, index=False)
        logging.debug("스케줄 저장: %s", self.schedule)
    
    def export_to_excel(self):
        data = []
        for date_str, entries in self.schedule.items():
            for entry in entries:
                data.append({
                    "date": date_str,
                    "name": entry["name"],
                    "shift": entry["shift"],
                    "absent": entry.get("absent", False),
                    "tardy": entry.get("tardy", False)
                })
        df = pd.DataFrame(data, columns=["date", "name", "shift", "absent", "tardy"])
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M")
        filename = f"{timestamp}.xlsx"
        df.to_excel(filename, index=False)
        logging.debug("엑셀 출력: %s", filename)
        return filename
    
    def add_schedule(self, name, shifts, weekday_info, start_date, end_date):
        current = start_date
        while current <= end_date:
            dow = current.dayOfWeek()
            if dow in weekday_info:
                day_cb, biweekly_cb, biweekly_combo = weekday_info[dow]
                if day_cb.isChecked():
                    key = current.toString(Qt.ISODate)
                    if key not in self.schedule:
                        self.schedule[key] = []
                    if biweekly_cb.isChecked():
                        week_offset = (current.toJulianDay() - start_date.toJulianDay()) // 7
                        base_shift = biweekly_combo.currentText()
                        shift_to_use = base_shift if week_offset % 2 == 0 else ("PM" if base_shift == "AM" else "AM")
                        self.schedule[key].append({"name": str(name), "shift": shift_to_use, "absent": False, "tardy": False})
                    else:
                        for shift in shifts:
                            self.schedule[key].append({"name": str(name), "shift": shift, "absent": False, "tardy": False})
            current = current.addDays(1)
        self.save_schedule()
    
    def delete_schedule(self, name, start_date, end_date):
        current = start_date
        while current <= end_date:
            key = current.toString(Qt.ISODate)
            if key in self.schedule:
                self.schedule[key] = [entry for entry in self.schedule[key] if entry["name"] != str(name)]
                if not self.schedule[key]:
                    del self.schedule[key]
            current = current.addDays(1)
        self.save_schedule()
    
    def toggle_shift(self, date, indices):
        key = date.toString(Qt.ISODate)
        if key in self.schedule:
            for idx in indices:
                entry = self.schedule[key][idx]
                entry["shift"] = "AM" if entry["shift"] == "PM" else "PM"
            self.save_schedule()

# ----- Custom Calendar Widget -----
class CustomCalendar(QCalendarWidget):
    def __init__(self, schedule_manager, holiday_info, name_color_map, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.schedule_manager = schedule_manager
        self.holiday_info = holiday_info
        self.name_color_map = name_color_map
        self.setGridVisible(True)
        self.setHorizontalHeaderFormat(QCalendarWidget.NoHorizontalHeader)
        self.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
    
    def paintCell(self, painter, rect, date):
        painter.save()
        super().paintCell(painter, rect, date)
        
        # 중앙 숫자 덮개
        bg_color = self.palette().color(QPalette.Base)
        center_rect = QRect(rect.center().x() - 8, rect.center().y() - 6, 16, 12)
        painter.fillRect(center_rect, bg_color)
        
        # 날짜 숫자 (좌측 상단)
        date_color = Qt.black
        date_str = date.toString("yyyy-MM-dd")
        if date.dayOfWeek() in (6, 7) or date_str in self.holiday_info:
            date_color = QColor("red")
        date_font = QFont("Arial", 14, QFont.Bold)
        painter.setFont(date_font)
        painter.setPen(date_color)
        painter.drawText(rect.adjusted(5, 5, -5, -5), Qt.AlignTop | Qt.AlignLeft, str(date.day()))
        
        # 공휴일명칭 (오른쪽 아래)
        if date_str in self.holiday_info:
            holiday_name = self.holiday_info[date_str]
            holiday_font = QFont("Arial", 8)
            painter.setFont(holiday_font)
            painter.setPen(QColor("red"))
            painter.drawText(rect.adjusted(0, 0, -2, -2), Qt.AlignBottom | Qt.AlignRight, holiday_name)
        
        # 스케줄 항목 출력
        key = date.toString(Qt.ISODate)
        if key in self.schedule_manager.schedule:
            entries = sorted(self.schedule_manager.schedule[key], key=lambda x: 0 if x["shift"]=="AM" else 1)
            entry_font_size = 10
            entry_font = QFont("Arial", entry_font_size, QFont.Bold)
            painter.setFont(entry_font)
            line_height = entry_font_size + 2
            entry_y = rect.y() + 30
            for entry in entries:
                shift = entry["shift"]
                name = entry["name"]
                if entry.get("absent", False):
                    painter.setPen(QColor("gray"))
                    text = f"{shift} {name}"
                    painter.drawText(rect.x() + 5, entry_y + entry_font_size, text)
                    fm = painter.fontMetrics()
                    text_width = fm.horizontalAdvance(text)
                    y_mid = entry_y + entry_font_size / 2
                    painter.drawLine(rect.x() + 5, y_mid, rect.x() + 5 + text_width, y_mid)
                elif entry.get("tardy", False):
                    tardy_color = QColor("darkorange")
                    painter.setPen(tardy_color)
                    text = f"{shift} {name} (지각)"
                    painter.drawText(rect.x() + 5, entry_y + entry_font_size, text)
                else:
                    shift_color = QColor("green") if shift=="AM" else QColor("darkred")
                    painter.setPen(shift_color)
                    painter.drawText(rect.x() + 5, entry_y + entry_font_size, shift)
                    fm = painter.fontMetrics()
                    shift_width = fm.horizontalAdvance(shift)
                    if name not in self.name_color_map:
                        r = random.randint(0, 255)
                        g = random.randint(0, 255)
                        b = random.randint(0, 255)
                        self.name_color_map[name] = QColor(r, g, b)
                    painter.setPen(self.name_color_map[name])
                    painter.drawText(rect.x() + 5 + shift_width, entry_y + entry_font_size, " " + name)
                entry_y += line_height
        painter.restore()
    
    def contextMenuEvent(self, event):
        menu = QMenu(self)
        action_shift_change = menu.addAction("근무조 변경")
        action_attendance = menu.addAction("출석 상태 변경")
        action_add = menu.addAction("해당일자 근무자 추가")
        action_delete = menu.addAction("해당일자 근무자 삭제")
        action = menu.exec(event.globalPos())
        date = self.selectedDate()
        key = date.toString(Qt.ISODate)
        if action == action_shift_change:
            if key in self.schedule_manager.schedule:
                dlg = ShiftChangeDialog(self.schedule_manager.schedule[key], self)
                if dlg.exec() == QDialog.Accepted:
                    indices = dlg.getSelectedIndices()
                    if indices:
                        self.schedule_manager.toggle_shift(date, indices)
                        self.updateCells()
        elif action == action_attendance:
            if key in self.schedule_manager.schedule:
                dlg = AttendanceDialog(self.schedule_manager.schedule[key], self)
                if dlg.exec() == QDialog.Accepted:
                    self.schedule_manager.save_schedule()
                    self.updateCells()
        elif action == action_add:
            dlg = AddWorkerDialog(self)
            if dlg.exec() == QDialog.Accepted:
                name, shift = dlg.getValues()
                if key not in self.schedule_manager.schedule:
                    self.schedule_manager.schedule[key] = []
                self.schedule_manager.schedule[key].append({"name": name, "shift": shift, "absent": False, "tardy": False})
                self.schedule_manager.save_schedule()
                self.updateCells()
        elif action == action_delete:
            if key in self.schedule_manager.schedule:
                dlg = DeleteWorkerDialog(self.schedule_manager.schedule[key], self)
                if dlg.exec() == QDialog.Accepted:
                    indices = dlg.getSelectedIndices()
                    if indices:
                        for i in sorted(indices, reverse=True):
                            del self.schedule_manager.schedule[key][i]
                        if not self.schedule_manager.schedule[key]:
                            del self.schedule_manager.schedule[key]
                        self.schedule_manager.save_schedule()
                        self.updateCells()

# ----- 공휴일 정보 가져오기 -----
def fetch_holiday_info_for_year(year):
    cache_file = f"holidays_{year}.json"
    if os.path.exists(cache_file):
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                holiday_info = json.load(f)
            logging.debug("캐시 로드 성공: %s", cache_file)
            return holiday_info
        except Exception as e:
            logging.error("캐시 로드 오류: %s", e)
    holiday_info = {}
    try:
        with open("key.txt", "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
            decoding_key = lines[1]
    except Exception as e:
        logging.error("key.txt 오류: %s", e)
        return holiday_info
    for month in range(1, 13):
        month_str = f"{month:02d}"
        url = (f"http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/"
               f"getRestDeInfo?solYear={year}&solMonth={month_str}&ServiceKey={decoding_key}")
        try:
            response = requests.get(url)
            tree = ET.fromstring(response.content)
            for item in tree.iter("item"):
                locdate = item.find("locdate")
                date_nm = item.find("dateName")
                if locdate is not None and locdate.text and date_nm is not None and date_nm.text:
                    d = locdate.text
                    formatted = f"{d[:4]}-{d[4:6]}-{d[6:]}"
                    holiday_info[formatted] = date_nm.text
            logging.debug("월 %s 정보 가져옴", month_str)
        except Exception as e:
            logging.error("API 호출 오류 (월 %s): %s", month_str, e)
    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(holiday_info, f, ensure_ascii=False, indent=2)
        logging.debug("캐시 저장: %s", cache_file)
    except Exception as e:
        logging.error("캐시 저장 오류: %s", e)
    return holiday_info

# ----- DateRangeDialog -----
class DateRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("날짜 범위 선택")
        cal_layout = QHBoxLayout()
        self.startCalendar = QCalendarWidget()
        self.endCalendar = QCalendarWidget()
        self.startCalendar.setMinimumSize(300, 300)
        self.endCalendar.setMinimumSize(300, 300)
        self.startCalendar.setStyleSheet("font-size: 14pt;")
        self.endCalendar.setStyleSheet("font-size: 14pt;")
        cal_layout.addWidget(self.startCalendar)
        cal_layout.addWidget(self.endCalendar)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        mainLayout = QVBoxLayout(self)
        mainLayout.addLayout(cal_layout)
        mainLayout.addWidget(buttons)
    def getDateRange(self):
        start_date = self.startCalendar.selectedDate()
        end_date = self.endCalendar.selectedDate()
        if start_date > end_date:
            start_date, end_date = end_date, start_date
        return start_date, end_date

# ----- Main Window -----
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("알바 근태 관리")
        self.schedule_manager = ScheduleManager()
        self.holiday_info = fetch_holiday_info_for_year(2025)
        self.name_color_map = {}
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        
        # 좌측 패널
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_widget.setMaximumWidth(300)
        
        # [근무자 추가] 그룹
        add_group = QGroupBox("근무자 추가")
        add_layout = QVBoxLayout(add_group)
        add_form = QFormLayout()
        self.name_edit = QLineEdit()
        add_form.addRow("이름:", self.name_edit)
        shift_layout = QHBoxLayout()
        self.am_check = QCheckBox("AM")
        self.pm_check = QCheckBox("PM")
        shift_layout.addWidget(self.am_check)
        shift_layout.addWidget(self.pm_check)
        add_form.addRow("글로벌 근무:", shift_layout)
        weekday_vlayout = QVBoxLayout()
        self.weekday_checks = {}
        for label, num in [("월", 1), ("화", 2), ("수", 3), ("목", 4), ("금", 5), ("토", 6), ("일", 7)]:
            hbox = QHBoxLayout()
            day_cb = QCheckBox(label)
            biweekly_cb = QCheckBox("격주")
            biweekly_combo = QComboBox()
            biweekly_combo.addItems(["AM", "PM"])
            biweekly_combo.setEnabled(False)
            biweekly_cb.toggled.connect(lambda checked, combo=biweekly_combo: combo.setEnabled(checked))
            hbox.addWidget(day_cb)
            hbox.addWidget(biweekly_cb)
            hbox.addWidget(biweekly_combo)
            weekday_vlayout.addLayout(hbox)
            self.weekday_checks[num] = (day_cb, biweekly_cb, biweekly_combo)
        add_form.addRow("요일 선택:", weekday_vlayout)
        date_range_layout = QHBoxLayout()
        self.date_range_display = QLineEdit()
        self.date_range_display.setReadOnly(True)
        self.date_range_button = QPushButton("날짜 선택")
        self.date_range_button.clicked.connect(self.select_add_date_range)
        self.this_month_add = QCheckBox("요번달")
        self.this_month_add.toggled.connect(self.set_this_month_add)
        date_range_layout.addWidget(self.date_range_display)
        date_range_layout.addWidget(self.date_range_button)
        date_range_layout.addWidget(self.this_month_add)
        add_form.addRow("기간:", date_range_layout)
        add_layout.addLayout(add_form)
        self.add_button = QPushButton("근무자 추가")
        self.add_button.clicked.connect(self.add_employee_schedule)
        add_layout.addWidget(self.add_button)
        left_layout.addWidget(add_group)
        
        # [근무자 삭제] 그룹
        del_group = QGroupBox("근무자 삭제")
        del_layout = QVBoxLayout(del_group)
        del_form = QFormLayout()
        self.del_name_combo = QComboBox()
        self.update_del_combo()
        del_form.addRow("이름 선택:", self.del_name_combo)
        del_date_layout = QHBoxLayout()
        self.del_date_range_display = QLineEdit()
        self.del_date_range_display.setReadOnly(True)
        self.del_date_range_button = QPushButton("날짜 선택")
        self.del_date_range_button.clicked.connect(self.select_del_date_range)
        self.this_month_del = QCheckBox("요번달")
        self.this_month_del.toggled.connect(self.set_this_month_del)
        del_date_layout.addWidget(self.del_date_range_display)
        del_date_layout.addWidget(self.del_date_range_button)
        del_date_layout.addWidget(self.this_month_del)
        del_form.addRow("기간:", del_date_layout)
        del_layout.addLayout(del_form)
        self.del_button = QPushButton("근무자 삭제")
        self.del_button.clicked.connect(self.delete_employee_schedule)
        del_layout.addWidget(self.del_button)
        left_layout.addWidget(del_group)
        
        # [엑셀 출력] 그룹
        export_group = QGroupBox("엑셀 출력")
        export_layout = QVBoxLayout(export_group)
        self.export_button = QPushButton("엑셀로 출력하기")
        self.export_button.clicked.connect(self.export_schedule)
        export_layout.addWidget(self.export_button)
        left_layout.addWidget(export_group)
        
        # [공휴일 정보 가져오기] 그룹
        holiday_group = QGroupBox("공휴일 정보 가져오기")
        holiday_layout = QVBoxLayout(holiday_group)
        self.holiday_button = QPushButton("공휴일 정보 가져오기")
        self.holiday_button.clicked.connect(self.fetch_holiday_info)
        holiday_layout.addWidget(self.holiday_button)
        left_layout.addWidget(holiday_group)
        
        # [달력 캡쳐] 그룹
        capture_group = QGroupBox("달력 캡쳐")
        capture_layout = QVBoxLayout(capture_group)
        self.capture_button = QPushButton("캡쳐 저장")
        self.capture_button.clicked.connect(self.capture_calendar)
        capture_layout.addWidget(self.capture_button)
        left_layout.addWidget(capture_group)
        
        main_layout.addWidget(left_widget)
        
        # 우측: 달력 및 요일 헤더
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        self.dayHeader = DayOfWeekHeader()
        right_layout.addWidget(self.dayHeader)
        self.calendar = CalendarWithContextMenu(self.schedule_manager, self.holiday_info, self.name_color_map)
        self.calendar.setMinimumSize(1200, 900)
        right_layout.addWidget(self.calendar)
        main_layout.addWidget(right_widget)
        
        self.add_start_date = None
        self.add_end_date = None
        self.del_start_date = None
        self.del_end_date = None

    def set_this_month_add(self, checked):
        if checked:
            today = QDate.currentDate()
            self.add_start_date = QDate(today.year(), today.month(), 1)
            self.add_end_date = QDate(today.year(), today.month(), today.daysInMonth())
            self.date_range_display.setText(
                f"{self.add_start_date.toString('yyyy-MM-dd')} ~ {self.add_end_date.toString('yyyy-MM-dd')}"
            )
            self.date_range_button.setEnabled(False)
        else:
            self.add_start_date = None
            self.add_end_date = None
            self.date_range_display.clear()
            self.date_range_button.setEnabled(True)

    def set_this_month_del(self, checked):
        if checked:
            today = QDate.currentDate()
            self.del_start_date = QDate(today.year(), today.month(), 1)
            self.del_end_date = QDate(today.year(), today.month(), today.daysInMonth())
            self.del_date_range_display.setText(
                f"{self.del_start_date.toString('yyyy-MM-dd')} ~ {self.del_end_date.toString('yyyy-MM-dd')}"
            )
            self.del_date_range_button.setEnabled(False)
        else:
            self.del_start_date = None
            self.del_end_date = None
            self.del_date_range_display.clear()
            self.del_date_range_button.setEnabled(True)

    def update_del_combo(self):
        names = set()
        for entries in self.schedule_manager.schedule.values():
            for entry in entries:
                names.add(entry["name"])
        current = self.del_name_combo.currentText() if self.del_name_combo.count() > 0 else ""
        self.del_name_combo.clear()
        for name in sorted(names):
            self.del_name_combo.addItem(str(name))
        index = self.del_name_combo.findText(current)
        if index >= 0:
            self.del_name_combo.setCurrentIndex(index)
    
    def select_add_date_range(self):
        if self.this_month_add.isChecked():
            return
        dialog = DateRangeDialog(self)
        if dialog.exec() == QDialog.Accepted:
            start_date, end_date = dialog.getDateRange()
            self.add_start_date = start_date
            self.add_end_date = end_date
            self.date_range_display.setText(
                f"{start_date.toString('yyyy-MM-dd')} ~ {end_date.toString('yyyy-MM-dd')}"
            )
    
    def select_del_date_range(self):
        if self.this_month_del.isChecked():
            return
        dialog = DateRangeDialog(self)
        if dialog.exec() == QDialog.Accepted:
            start_date, end_date = dialog.getDateRange()
            self.del_start_date = start_date
            self.del_end_date = end_date
            self.del_date_range_display.setText(
                f"{start_date.toString('yyyy-MM-dd')} ~ {end_date.toString('yyyy-MM-dd')}"
            )
    
    def add_employee_schedule(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "입력 오류", "이름을 입력하세요.")
            return
        shifts = []
        if self.am_check.isChecked():
            shifts.append("AM")
        if self.pm_check.isChecked():
            shifts.append("PM")
        if not shifts:
            QMessageBox.warning(self, "입력 오류", "글로벌 근무(AM/PM) 중 최소 한 가지를 선택하세요.")
            return
        if self.add_start_date is None or self.add_end_date is None:
            QMessageBox.warning(self, "입력 오류", "근무 기간을 선택하세요.")
            return
        self.schedule_manager.add_schedule(name, shifts, self.weekday_checks,
                                             self.add_start_date, self.add_end_date)
        self.calendar.updateCells()
        self.update_del_combo()
        QMessageBox.information(self, "완료", "근무자 추가가 완료되었습니다.")
        self.name_edit.clear()
        self.am_check.setChecked(False)
        self.pm_check.setChecked(False)
    
    def delete_employee_schedule(self):
        name = self.del_name_combo.currentText()
        if not name:
            QMessageBox.warning(self, "입력 오류", "삭제할 근무자를 선택하세요.")
            return
        if self.del_start_date is None or self.del_end_date is None:
            QMessageBox.warning(self, "입력 오류", "삭제할 기간을 선택하세요.")
            return
        self.schedule_manager.delete_schedule(name, self.del_start_date, self.del_end_date)
        self.calendar.updateCells()
        self.update_del_combo()
        QMessageBox.information(self, "완료", f"{name}님의 근무 스케줄이 삭제되었습니다.")
        self.del_date_range_display.clear()
        self.del_start_date = None
        self.del_end_date = None

    def export_schedule(self):
        filename = self.schedule_manager.export_to_excel()
        QMessageBox.information(self, "엑셀 출력 완료", f"엑셀 파일이 생성되었습니다.\n파일명: {filename}")

    def fetch_holiday_info(self):
        current_year = datetime.datetime.now().year
        self.holiday_info = fetch_holiday_info_for_year(current_year)
        self.calendar.holiday_info = self.holiday_info
        self.calendar.updateCells()
        QMessageBox.information(self, "완료", "공휴일 정보를 가져왔습니다.")

    def capture_calendar(self):
        # 캡쳐 설정 대화상자 실행하여 오프셋과 영역 크기 조정
        settingsDialog = CaptureSettingsDialog(self)
        if settingsDialog.exec() != QDialog.Accepted:
            return
        x_offset, y_offset, desired_width, desired_height = settingsDialog.getSettings()
        
        # 메인 윈도우 전체 캡쳐
        screen = QApplication.primaryScreen()
        full_pixmap = screen.grabWindow(self.winId())
        
        # centralWidget 기준 오프셋 계산
        parent_pos = self.centralWidget().mapToGlobal(QPoint(0, 0))
        cal_pos = self.calendar.mapToGlobal(QPoint(0, 0))
        offset = cal_pos - parent_pos
        
        # 사용자 설정 오프셋 적용
        offset.setX(offset.x() + x_offset)
        offset.setY(offset.y() + y_offset)
        
        # 캘린더 위젯 영역 대신 사용자 지정 영역 사용
        cal_rect = QRect(offset, QSize(desired_width, desired_height))
        cropped_pixmap = full_pixmap.copy(cal_rect)
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M")
        filename = f"{timestamp}.jpg"
        if cropped_pixmap.save(filename, "JPG"):
            QMessageBox.information(self, "캡쳐 완료", f"캡쳐 파일이 저장되었습니다.\n파일명: {filename}")
        else:
            QMessageBox.warning(self, "캡쳐 오류", "캡쳐 저장에 실패하였습니다.")

# ----- CalendarWithContextMenu -----
class CalendarWithContextMenu(CustomCalendar):
    def contextMenuEvent(self, event):
        menu = QMenu(self)
        action_shift_change = menu.addAction("근무조 변경")
        action_attendance = menu.addAction("출석 상태 변경")
        action_add = menu.addAction("해당일자 근무자 추가")
        action_delete = menu.addAction("해당일자 근무자 삭제")
        action = menu.exec(event.globalPos())
        date = self.selectedDate()
        key = date.toString(Qt.ISODate)
        if action == action_shift_change:
            if key in self.schedule_manager.schedule:
                dlg = ShiftChangeDialog(self.schedule_manager.schedule[key], self)
                if dlg.exec() == QDialog.Accepted:
                    indices = dlg.getSelectedIndices()
                    if indices:
                        self.schedule_manager.toggle_shift(date, indices)
                        self.updateCells()
        elif action == action_attendance:
            if key in self.schedule_manager.schedule:
                dlg = AttendanceDialog(self.schedule_manager.schedule[key], self)
                if dlg.exec() == QDialog.Accepted:
                    self.schedule_manager.save_schedule()
                    self.updateCells()
        elif action == action_add:
            dlg = AddWorkerDialog(self)
            if dlg.exec() == QDialog.Accepted:
                name, shift = dlg.getValues()
                if key not in self.schedule_manager.schedule:
                    self.schedule_manager.schedule[key] = []
                self.schedule_manager.schedule[key].append({"name": name, "shift": shift, "absent": False, "tardy": False})
                self.schedule_manager.save_schedule()
                self.updateCells()
        elif action == action_delete:
            if key in self.schedule_manager.schedule:
                dlg = DeleteWorkerDialog(self.schedule_manager.schedule[key], self)
                if dlg.exec() == QDialog.Accepted:
                    indices = dlg.getSelectedIndices()
                    if indices:
                        for i in sorted(indices, reverse=True):
                            del self.schedule_manager.schedule[key][i]
                        if not self.schedule_manager.schedule[key]:
                            del self.schedule_manager.schedule[key]
                        self.schedule_manager.save_schedule()
                        self.updateCells()

# ----- 공휴일 정보 가져오기 -----
def fetch_holiday_info_for_year(year):
    cache_file = f"holidays_{year}.json"
    if os.path.exists(cache_file):
        try:
            with open(cache_file, "r", encoding="utf-8") as f:
                holiday_info = json.load(f)
            logging.debug("캐시 로드 성공: %s", cache_file)
            return holiday_info
        except Exception as e:
            logging.error("캐시 로드 오류: %s", e)
    holiday_info = {}
    try:
        with open("key.txt", "r", encoding="utf-8") as f:
            lines = f.read().splitlines()
            decoding_key = lines[1]
    except Exception as e:
        logging.error("key.txt 오류: %s", e)
        return holiday_info
    for month in range(1, 13):
        month_str = f"{month:02d}"
        url = (f"http://apis.data.go.kr/B090041/openapi/service/SpcdeInfoService/"
               f"getRestDeInfo?solYear={year}&solMonth={month_str}&ServiceKey={decoding_key}")
        try:
            response = requests.get(url)
            tree = ET.fromstring(response.content)
            for item in tree.iter("item"):
                locdate = item.find("locdate")
                date_nm = item.find("dateName")
                if locdate is not None and locdate.text and date_nm is not None and date_nm.text:
                    d = locdate.text
                    formatted = f"{d[:4]}-{d[4:6]}-{d[6:]}"
                    holiday_info[formatted] = date_nm.text
            logging.debug("월 %s 정보 가져옴", month_str)
        except Exception as e:
            logging.error("API 호출 오류 (월 %s): %s", month_str, e)
    try:
        with open(cache_file, "w", encoding="utf-8") as f:
            json.dump(holiday_info, f, ensure_ascii=False, indent=2)
        logging.debug("캐시 저장: %s", cache_file)
    except Exception as e:
        logging.error("캐시 저장 오류: %s", e)
    return holiday_info

# ----- DateRangeDialog -----
class DateRangeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("날짜 범위 선택")
        cal_layout = QHBoxLayout()
        self.startCalendar = QCalendarWidget()
        self.endCalendar = QCalendarWidget()
        self.startCalendar.setMinimumSize(300, 300)
        self.endCalendar.setMinimumSize(300, 300)
        self.startCalendar.setStyleSheet("font-size: 14pt;")
        self.endCalendar.setStyleSheet("font-size: 14pt;")
        cal_layout.addWidget(self.startCalendar)
        cal_layout.addWidget(self.endCalendar)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        mainLayout = QVBoxLayout(self)
        mainLayout.addLayout(cal_layout)
        mainLayout.addWidget(buttons)
    def getDateRange(self):
        start_date = self.startCalendar.selectedDate()
        end_date = self.endCalendar.selectedDate()
        if start_date > end_date:
            start_date, end_date = end_date, start_date
        return start_date, end_date

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    # 우클릭 메뉴가 있는 캘린더 위젯으로 교체
    window.calendar = CalendarWithContextMenu(window.schedule_manager, window.holiday_info, window.name_color_map)
    window.resize(1200, 900)
    window.show()
    sys.exit(app.exec())
