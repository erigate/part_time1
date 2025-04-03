import sys
import os
import datetime
import logging
import pandas as pd

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QCalendarWidget,
    QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit,
    QPushButton, QCheckBox, QMessageBox, QDialog,
    QDialogButtonBox, QGroupBox, QComboBox, QTextEdit
)
from PySide6.QtCore import QDate, Qt, QRect
from PySide6.QtGui import QPainter, QColor, QFont

# 로그 설정: debug_log.txt 파일에 로그 기록 (덮어쓰기 모드)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='debug_log.txt',
    filemode='w'
)
logging.debug("프로그램 시작")

# ----- Log Dialog (로그 내용 보기) -----
class LogDialog(QDialog):
    def __init__(self, log_file, parent=None):
        super().__init__(parent)
        self.setWindowTitle("로그 보기")
        self.resize(600, 400)
        layout = QVBoxLayout(self)
        self.textEdit = QTextEdit()
        self.textEdit.setReadOnly(True)
        layout.addWidget(self.textEdit)
        try:
            with open(log_file, "r", encoding="utf-8") as f:
                self.textEdit.setPlainText(f.read())
        except Exception as e:
            self.textEdit.setPlainText(f"로그 파일을 읽을 수 없습니다: {e}")

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

# ----- Custom Day-Of-Week Header Widget -----
class DayOfWeekHeader(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.days = ["월", "화", "수", "목", "금", "토", "일"]
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
        self.schedule = {}  # key: QDate.toString(Qt.ISODate), value: list of dicts {name, shift}
        self.ensure_excel_file()
        self.load_schedule()
    
    def ensure_excel_file(self):
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame(columns=["date", "name", "shift"])
            df.to_excel(self.excel_file, index=False)
            logging.debug("엑셀 파일이 없어서 새로 생성함")
        else:
            try:
                df = pd.read_excel(self.excel_file)
                if not set(["date", "name", "shift"]).issubset(df.columns):
                    df = pd.DataFrame(columns=["date", "name", "shift"])
                    df.to_excel(self.excel_file, index=False)
                    logging.debug("엑셀 파일의 컬럼이 올바르지 않아 재생성함")
            except Exception as e:
                logging.error("Excel 파일 확인 오류: %s", e)
                df = pd.DataFrame(columns=["date", "name", "shift"])
                df.to_excel(self.excel_file, index=False)
    
    def load_schedule(self):
        try:
            df = pd.read_excel(self.excel_file, converters={"date": lambda x: x})
            logging.debug("엑셀 파일 로드 성공: %s", self.excel_file)
            if "date" not in df.columns:
                self.schedule = {}
                logging.error("엑셀 파일에 'date' 컬럼이 없음")
                return
            for idx, row in df.iterrows():
                logging.debug("Row %d: date=%s, name=%s, shift=%s", idx, row['date'], row['name'], row['shift'])
                try:
                    date_val = pd.to_datetime(row['date'])
                    date_str = date_val.strftime("%Y-%m-%d")
                    logging.debug("pd.to_datetime() 변환 성공: %s", date_str)
                except Exception as e:
                    logging.error("pd.to_datetime() 변환 실패: %s", e)
                    date_str = str(row['date']).strip().split(" ")[0]
                    logging.debug("Fallback 변환: %s", date_str)
                date = QDate.fromString(date_str, "yyyy-MM-dd")
                if date.isValid():
                    key = date.toString(Qt.ISODate)
                    if key not in self.schedule:
                        self.schedule[key] = []
                    # 이름을 문자열로 변환하여 저장
                    self.schedule[key].append({"name": str(row['name']), "shift": row['shift']})
                    logging.debug("스케줄 추가됨: %s -> %s", key, {"name": str(row['name']), "shift": row['shift']})
                else:
                    logging.error("QDate 변환 실패: date_str=%s", date_str)
            logging.debug("최종 스케줄: %s", self.schedule)
        except Exception as e:
            logging.error("Excel 파일 읽기 오류: %s", e)
            self.schedule = {}
    
    def save_schedule(self):
        data = []
        for date_str, entries in self.schedule.items():
            for entry in entries:
                data.append({"date": date_str, "name": entry["name"], "shift": entry["shift"]})
        df = pd.DataFrame(data, columns=["date", "name", "shift"])
        df.to_excel(self.excel_file, index=False)
        logging.debug("스케줄 저장됨: %s", self.schedule)
    
    def export_to_excel(self):
        data = []
        for date_str, entries in self.schedule.items():
            for entry in entries:
                data.append({"date": date_str, "name": entry["name"], "shift": entry["shift"]})
        df = pd.DataFrame(data, columns=["date", "name", "shift"])
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f"{timestamp}.xlsx"
        df.to_excel(filename, index=False)
        logging.debug("엑셀로 출력: %s", filename)
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
                        self.schedule[key].append({"name": str(name), "shift": shift_to_use})
                        logging.debug("격주 스케줄 추가: %s -> %s", key, {"name": str(name), "shift": shift_to_use})
                    else:
                        for shift in shifts:
                            self.schedule[key].append({"name": str(name), "shift": shift})
                            logging.debug("글로벌 스케줄 추가: %s -> %s", key, {"name": str(name), "shift": shift})
            current = current.addDays(1)
        self.save_schedule()
    
    def delete_schedule(self, name, start_date, end_date):
        current = start_date
        while current <= end_date:
            key = current.toString(Qt.ISODate)
            if key in self.schedule:
                before = len(self.schedule[key])
                # 이름 비교를 문자열로 처리
                self.schedule[key] = [entry for entry in self.schedule[key] if entry["name"] != str(name)]
                after = len(self.schedule.get(key, []))
                logging.debug("삭제 전 %d, 삭제 후 %d for key %s", before, after, key)
                if not self.schedule[key]:
                    del self.schedule[key]
                    logging.debug("스케줄 키 삭제됨: %s", key)
            current = current.addDays(1)
        self.save_schedule()
    
    def toggle_shift(self, date, indices):
        key = date.toString(Qt.ISODate)
        if key in self.schedule:
            for idx in indices:
                entry = self.schedule[key][idx]
                entry["shift"] = "AM" if entry["shift"] == "PM" else "PM"
                logging.debug("스케줄 토글: key=%s, index=%d, new shift=%s", key, idx, entry["shift"])
            self.save_schedule()

# ----- Custom Calendar Widget -----
class CustomCalendar(QCalendarWidget):
    def __init__(self, schedule_manager, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.schedule_manager = schedule_manager
        self.setGridVisible(True)
        self.setHorizontalHeaderFormat(QCalendarWidget.NoHorizontalHeader)
        self.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
    
    def paintCell(self, painter, rect, date):
        painter.save()
        super().paintCell(painter, rect, date)
        date_font = QFont("Arial", 14, QFont.Bold)
        painter.setFont(date_font)
        painter.setPen(Qt.black)
        painter.drawText(rect.adjusted(5, 5, -5, -5), Qt.AlignTop | Qt.AlignLeft, str(date.day()))
        key = date.toString(Qt.ISODate)
        if key in self.schedule_manager.schedule:
            entries = self.schedule_manager.schedule[key]
            num_entries = len(entries)
            available_height = rect.height() - 30
            total_default_height = num_entries * (24 + 5)
            scale = min(1.0, available_height / total_default_height) if num_entries > 0 else 1.0
            circle_diameter = int(24 * scale)
            entry_font_size = int(12 * scale)
            entry_y = rect.y() + 30
            for entry in entries:
                shift = entry["shift"]
                name = entry["name"]
                circle_rect = QRect(rect.x() + 5, entry_y, circle_diameter, circle_diameter)
                circle_color = QColor("green") if shift == "AM" else QColor("red") if shift == "PM" else Qt.black
                painter.setPen(circle_color)
                painter.drawEllipse(circle_rect)
                painter.drawText(circle_rect, Qt.AlignCenter, shift)
                entry_font_bold = QFont("Arial", entry_font_size, QFont.Bold)
                painter.setFont(entry_font_bold)
                painter.setPen(Qt.black)
                painter.drawText(
                    int(rect.x() + 5 + circle_diameter + 5),
                    int(entry_y + circle_diameter/2 + entry_font_size/2 - 2),
                    str(name)
                )
                entry_y += circle_diameter + 5
        painter.restore()
    
    def contextMenuEvent(self, event):
        date = self.selectedDate()
        key = date.toString(Qt.ISODate)
        if key in self.schedule_manager.schedule:
            entries = self.schedule_manager.schedule[key]
            if not entries:
                return
            dlg = ShiftChangeDialog(entries, self)
            if dlg.exec() == QDialog.Accepted:
                indices = dlg.getSelectedIndices()
                if indices:
                    self.schedule_manager.toggle_shift(date, indices)
                    self.updateCells()

# ----- 날짜 범위 선택 다이얼로그 -----
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
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        
        # 좌측 패널: 추가/삭제/엑셀 출력/로그 보기 폼 (최대 폭 300)
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
        date_range_layout.addWidget(self.date_range_display)
        date_range_layout.addWidget(self.date_range_button)
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
        del_date_layout.addWidget(self.del_date_range_display)
        del_date_layout.addWidget(self.del_date_range_button)
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
        
        # [로그 확인] 그룹
        log_group = QGroupBox("로그 확인")
        log_layout = QVBoxLayout(log_group)
        self.log_button = QPushButton("로그 보기")
        self.log_button.clicked.connect(self.open_log_dialog)
        log_layout.addWidget(self.log_button)
        left_layout.addWidget(log_group)
        
        main_layout.addWidget(left_widget)
        
        # 우측: 달력 및 요일 헤더
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        self.dayHeader = DayOfWeekHeader()
        right_layout.addWidget(self.dayHeader)
        self.calendar = CustomCalendar(self.schedule_manager)
        self.calendar.setMinimumSize(600, 600)
        right_layout.addWidget(self.calendar)
        main_layout.addWidget(right_widget)
        
        self.add_start_date = None
        self.add_end_date = None
        self.del_start_date = None
        self.del_end_date = None
    
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
        dialog = DateRangeDialog(self)
        if dialog.exec() == QDialog.Accepted:
            start_date, end_date = dialog.getDateRange()
            self.add_start_date = start_date
            self.add_end_date = end_date
            self.date_range_display.setText(f"{start_date.toString('yyyy-MM-dd')} ~ {end_date.toString('yyyy-MM-dd')}")
    
    def select_del_date_range(self):
        dialog = DateRangeDialog(self)
        if dialog.exec() == QDialog.Accepted:
            start_date, end_date = dialog.getDateRange()
            self.del_start_date = start_date
            self.del_end_date = end_date
            self.del_date_range_display.setText(f"{start_date.toString('yyyy-MM-dd')} ~ {end_date.toString('yyyy-MM-dd')}")
    
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

    def open_log_dialog(self):
        log_dialog = LogDialog("debug_log.txt", self)
        log_dialog.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1000, 700)
    window.show()
    sys.exit(app.exec())
