import sys
import os
import datetime
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QCalendarWidget,
                               QVBoxLayout, QHBoxLayout, QFormLayout, QLineEdit,
                               QPushButton, QCheckBox, QMessageBox, QDialog,
                               QDialogButtonBox, QGroupBox, QComboBox)
from PySide6.QtCore import QDate, Qt, QRect
from PySide6.QtGui import QPainter, QColor, QFont
import pandas as pd

# ----- Shift Change Dialog -----
class ShiftChangeDialog(QDialog):
    def __init__(self, entries, parent=None):
        """
        entries: [{'name':..., 'shift':...}, ...]
        기본은 전체 체크된 상태로 보여주며,
        사용자가 원하는 항목만 선택하면 해당 항목만 shift 토글 처리.
        """
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
        # 한글 요일: 월 ~ 일
        self.days = ["월", "화", "수", "목", "금", "토", "일"]
        self.setMinimumHeight(30)
    
    def resizeEvent(self, event):
        self.update()
        super().resizeEvent(event)
    
    def paintEvent(self, event):
        painter = QPainter(self)
        width = self.width() / 7
        height = self.height()
        # 각 칸의 글자 크기를 칸 높이의 절반 정도로 설정
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
        else:
            try:
                df = pd.read_excel(self.excel_file)
                if not set(["date", "name", "shift"]).issubset(df.columns):
                    df = pd.DataFrame(columns=["date", "name", "shift"])
                    df.to_excel(self.excel_file, index=False)
            except Exception as e:
                print("Excel 파일 확인 오류:", e)
                df = pd.DataFrame(columns=["date", "name", "shift"])
                df.to_excel(self.excel_file, index=False)
    
    def load_schedule(self):
        try:
            df = pd.read_excel(self.excel_file)
            if "date" not in df.columns:
                self.schedule = {}
                return
            for _, row in df.iterrows():
                date_str = str(row['date'])
                name = row['name']
                shift = row['shift']
                date = QDate.fromString(date_str, "yyyy-MM-dd")
                if date.isValid():
                    key = date.toString(Qt.ISODate)
                    if key not in self.schedule:
                        self.schedule[key] = []
                    self.schedule[key].append({"name": name, "shift": shift})
        except Exception as e:
            print("Excel 파일 읽기 오류:", e)
            self.schedule = {}
    
    def save_schedule(self):
        data = []
        for date_str, entries in self.schedule.items():
            for entry in entries:
                data.append({"date": date_str, "name": entry["name"], "shift": entry["shift"]})
        df = pd.DataFrame(data, columns=["date", "name", "shift"])
        df.to_excel(self.excel_file, index=False)
    
    def export_to_excel(self):
        """현재 스케줄 데이터를 현재 날짜 및 시간 기반 파일명(YYYYMMDDHHMMSS.xlsx)으로 저장"""
        data = []
        for date_str, entries in self.schedule.items():
            for entry in entries:
                data.append({"date": date_str, "name": entry["name"], "shift": entry["shift"]})
        df = pd.DataFrame(data, columns=["date", "name", "shift"])
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        filename = f"{timestamp}.xlsx"
        df.to_excel(filename, index=False)
        return filename
    
    def add_schedule(self, name, shifts, weekday_info, start_date, end_date):
        """
        weekday_info: {요일번호: (day_checkbox, biweekly_checkbox, biweekly_combo), ...}
        - 만약 biweekly_checkbox가 체크되어 있으면,
          시작일 기준 주차에 따라, 짝수 주에는 biweekly_combo에서 선택한 값을, 홀수 주에는 반대값을 적용.
        - 격주 미체크 시에는 글로벌 shifts(상단의 체크박스)를 모두 적용.
        """
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
                        if week_offset % 2 == 0:
                            shift_to_use = base_shift
                        else:
                            shift_to_use = "PM" if base_shift == "AM" else "AM"
                        self.schedule[key].append({"name": name, "shift": shift_to_use})
                    else:
                        for shift in shifts:
                            self.schedule[key].append({"name": name, "shift": shift})
            current = current.addDays(1)
        self.save_schedule()
    
    def delete_schedule(self, name, start_date, end_date):
        current = start_date
        while current <= end_date:
            key = current.toString(Qt.ISODate)
            if key in self.schedule:
                self.schedule[key] = [entry for entry in self.schedule[key] if entry["name"] != name]
                if not self.schedule[key]:
                    del self.schedule[key]
            current = current.addDays(1)
        self.save_schedule()
    
    def toggle_shift(self, date, indices):
        """
        indices: 리스트로, 해당 날짜의 schedule에서 선택된 항목 인덱스에 대해 shift 토글
        """
        key = date.toString(Qt.ISODate)
        if key in self.schedule:
            for idx in indices:
                entry = self.schedule[key][idx]
                entry["shift"] = "AM" if entry["shift"] == "PM" else "PM"
            self.save_schedule()

# ----- Custom Calendar Widget -----
class CustomCalendar(QCalendarWidget):
    def __init__(self, schedule_manager, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.schedule_manager = schedule_manager
        self.setGridVisible(True)
        # 내장 요일/주차 헤더 숨김
        self.setHorizontalHeaderFormat(QCalendarWidget.NoHorizontalHeader)
        self.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
    
    def paintCell(self, painter, rect, date):
        painter.save()
        # 기본 셀 배경 그리기
        super().paintCell(painter, rect, date)
        
        # 날짜 숫자: 좌측 상단에 굵고 크게 (14pt)
        date_font = QFont("Arial", 14, QFont.Bold)
        painter.setFont(date_font)
        painter.setPen(Qt.black)
        painter.drawText(rect.adjusted(5, 5, -5, -5), Qt.AlignTop | Qt.AlignLeft, str(date.day()))
        
        # 스케줄 항목: 날짜 숫자 아래쪽부터 표시 (반응형 폰트/아이콘 크기)
        key = date.toString(Qt.ISODate)
        if key in self.schedule_manager.schedule:
            entries = self.schedule_manager.schedule[key]
            num_entries = len(entries)
            available_height = rect.height() - 30  # 날짜 숫자 아래 영역
            total_default_height = num_entries * (24 + 5)
            scale = min(1.0, available_height / total_default_height) if num_entries > 0 else 1.0
            circle_diameter = int(24 * scale)
            entry_font_size = int(12 * scale)
            entry_y = rect.y() + 30
            for entry in entries:
                shift = entry["shift"]
                name = entry["name"]
                circle_rect = QRect(rect.x() + 5, entry_y, circle_diameter, circle_diameter)
                if shift == "AM":
                    circle_color = QColor("green")
                elif shift == "PM":
                    circle_color = QColor("red")
                else:
                    circle_color = Qt.black
                painter.setPen(circle_color)
                painter.drawEllipse(circle_rect)
                painter.drawText(circle_rect, Qt.AlignCenter, shift)
                entry_font_bold = QFont("Arial", entry_font_size, QFont.Bold)
                painter.setFont(entry_font_bold)
                painter.setPen(Qt.black)
                # 좌표 값을 정수로 변환하여 전달
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
        # 캘린더 크기 및 폰트 스타일 조정 (눈에 잘 보이도록)
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
        
        # 좌측 패널: 추가/삭제/엑셀 출력 폼 (최대 폭 300)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_widget.setMaximumWidth(300)
        
        # [근무자 추가] 그룹
        add_group = QGroupBox("근무자 추가")
        add_layout = QVBoxLayout(add_group)
        add_form = QFormLayout()
        self.name_edit = QLineEdit()
        add_form.addRow("이름:", self.name_edit)
        
        # 글로벌 AM, PM 체크박스 (격주 미적용 시 사용)
        shift_layout = QHBoxLayout()
        self.am_check = QCheckBox("AM")
        self.pm_check = QCheckBox("PM")
        shift_layout.addWidget(self.am_check)
        shift_layout.addWidget(self.pm_check)
        add_form.addRow("글로벌 근무:", shift_layout)
        
        # 요일 선택: 세로 7줄, 각 줄에 [요일] + [격주] 체크 + [기준 Shift 선택]
        weekday_vlayout = QVBoxLayout()
        self.weekday_checks = {}
        for label, num in [("월", 1), ("화", 2), ("수", 3), ("목", 4), ("금", 5), ("토", 6), ("일", 7)]:
            hbox = QHBoxLayout()
            day_cb = QCheckBox(label)
            biweekly_cb = QCheckBox("격주")
            biweekly_combo = QComboBox()
            biweekly_combo.addItems(["AM", "PM"])
            biweekly_combo.setEnabled(False)
            # 격주 체크 여부에 따라 콤보 활성화
            biweekly_cb.toggled.connect(lambda checked, combo=biweekly_combo: combo.setEnabled(checked))
            hbox.addWidget(day_cb)
            hbox.addWidget(biweekly_cb)
            hbox.addWidget(biweekly_combo)
            weekday_vlayout.addLayout(hbox)
            self.weekday_checks[num] = (day_cb, biweekly_cb, biweekly_combo)
        add_form.addRow("요일 선택:", weekday_vlayout)
        
        # 날짜 범위 선택 (다이얼로그 사용)
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
        
        main_layout.addWidget(left_widget)
        
        # 우측: 커스텀 헤더와 달력 (달력 크게 표시)
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
            self.del_name_combo.addItem(str(name))  # 문자열로 변환
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
        # 글로벌 Shift는 격주 미적용 요일에 적용
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
        # 이름 및 글로벌 shift는 초기화하지만, 요일/격주/날짜 선택은 그대로 유지
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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1000, 700)
    window.show()
    sys.exit(app.exec())
