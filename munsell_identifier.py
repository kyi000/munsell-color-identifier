#!/usr/bin/env python3
"""
먼셀 색상 자동 인식 프로그램
마우스 커서 위치의 색상을 감지하여 가장 가까운 먼셀 색상 코드를 표시합니다.
"""

import sys
import os
import csv
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                            QListWidget, QListWidgetItem, QMessageBox, 
                            QToolTip, QSplitter)
from PyQt5.QtGui import QPixmap, QImage, QColor, QPainter, QPen, QFont, QCursor
from PyQt5.QtCore import Qt, QPoint, QTimer, QEvent, QSize

import numpy as np
from PIL import Image, ImageQt
import cv2

# Windows 환경에서 필요한 모듈
if sys.platform == "win32":
    import win32api
    import win32gui
    import win32ui
    import win32con
# macOS/Linux 환경에서 필요한 모듈
else:
    try:
        import mss
    except ImportError:
        pass
    try:
        from PIL import ImageGrab
    except ImportError:
        pass

from color_utils import find_closest_munsell, rgb_to_hex, get_munsell_color_rgb

class MunsellColorDisplay(QLabel):
    """
    먼셀 색상 및 코드를 표시하는 위젯
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(100, 60)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("border: 1px solid #aaa; border-radius: 4px;")
        
        # 기본 텍스트 설정
        self.setText("색상 코드")
        
    def update_color(self, rgb, munsell_code):
        """
        색상과 먼셀 코드 업데이트
        """
        color_hex = rgb_to_hex(rgb)
        
        # 배경색 설정
        self.setStyleSheet(f"background-color: {color_hex}; color: {'#000' if sum(rgb) > 384 else '#fff'}; "
                         f"font-weight: bold; border: 1px solid #aaa; border-radius: 4px;")
        
        # 텍스트 설정
        self.setText(f"{munsell_code}\n({rgb[0]}, {rgb[1]}, {rgb[2]})")

class ColorHistoryItem:
    """
    색상 히스토리 아이템
    """
    def __init__(self, rgb, munsell_code, timestamp=None):
        self.rgb = rgb
        self.munsell_code = munsell_code
        self.timestamp = timestamp if timestamp else datetime.now()
    
    def __str__(self):
        return f"{self.munsell_code} - RGB({self.rgb[0]}, {self.rgb[1]}, {self.rgb[2]})"

class ImageViewer(QLabel):
    """
    이미지 표시 및 색상 감지를 위한 위젯
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 300)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("border: 1px solid #aaa; background-color: #f0f0f0;")
        self.setText("이미지를 로드하세요")
        
        self.original_image = None
        self.scaled_image = None
        self.image_offset = QPoint(0, 0)
        
        # 마우스 추적 활성화
        self.setMouseTracking(True)
        
        # 먼셀 색상 툴팁 설정
        QToolTip.setFont(QFont('SansSerif', 10))
        
    def load_image(self, image_path):
        """
        이미지 로드 후 표시
        """
        try:
            # PIL.Image로 먼저 로드 (색상 처리를 위해)
            self.pil_image = Image.open(image_path).convert('RGB')
            
            # QPixmap으로 변환
            self.original_image = QPixmap.fromImage(ImageQt.ImageQt(self.pil_image))
            
            # 이미지 크기 조정 및 표시
            self.resize_image()
            
            return True
        except Exception as e:
            QMessageBox.critical(self, "이미지 로드 오류", f"이미지를 로드하는 중 오류가 발생했습니다: {str(e)}")
            return False
    
    def resize_image(self):
        """
        이미지를 위젯 크기에 맞게 조정
        """
        if self.original_image:
            self.scaled_image = self.original_image.scaled(
                self.size(), 
                Qt.KeepAspectRatio, 
                Qt.SmoothTransformation
            )
            self.update()
    
    def paintEvent(self, event):
        """
        이미지 및 추가 요소 그리기
        """
        super().paintEvent(event)
        
        if self.scaled_image:
            painter = QPainter(self)
            
            # 이미지 크기가 위젯보다 작을 경우 중앙에 배치
            self.image_offset = QPoint(
                max(0, (self.width() - self.scaled_image.width()) // 2),
                max(0, (self.height() - self.scaled_image.height()) // 2)
            )
            
            # 이미지 그리기
            painter.drawPixmap(self.image_offset, self.scaled_image)
    
    def get_image_pixel_color(self, pos):
        """
        이미지 내 특정 위치의 픽셀 색상 반환
        """
        if not self.pil_image or not self.scaled_image:
            return None
        
        # 위젯 내 위치를 이미지 내 위치로 변환
        pos_in_image = pos - self.image_offset
        
        # 이미지 영역 밖이면 None 반환
        if (pos_in_image.x() < 0 or pos_in_image.y() < 0 or 
            pos_in_image.x() >= self.scaled_image.width() or 
            pos_in_image.y() >= self.scaled_image.height()):
            return None
        
        # 스케일링된 이미지에서 원본 이미지의 좌표 계산
        orig_x = int(pos_in_image.x() * self.pil_image.width / self.scaled_image.width())
        orig_y = int(pos_in_image.y() * self.pil_image.height / self.scaled_image.height())
        
        # 원본 이미지에서 색상 추출
        try:
            r, g, b = self.pil_image.getpixel((orig_x, orig_y))
            return (r, g, b)
        except:
            return None
    
    def resizeEvent(self, event):
        """
        위젯 크기 변경 시 이미지도 함께 조정
        """
        super().resizeEvent(event)
        self.resize_image()

class ScreenColorPicker:
    """
    화면의 색상을 캡처하는 클래스
    """
    def __init__(self):
        # 스크린샷 관련 변수 초기화
        self.screenshot = None
    
    def update_screenshot(self):
        """
        전체 화면 스크린샷 업데이트
        """
        # OpenCV를 사용하여 화면 캡처
        try:
            # Windows
            if sys.platform == "win32":
                import win32gui
                import win32ui
                import win32con
                import win32api
                
                # 전체 화면 크기 가져오기
                width = win32api.GetSystemMetrics(win32con.SM_CXVIRTUALSCREEN)
                height = win32api.GetSystemMetrics(win32con.SM_CYVIRTUALSCREEN)
                left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
                
                # 화면 캡처
                hwin = win32gui.GetDesktopWindow()
                hwindc = win32gui.GetWindowDC(hwin)
                srcdc = win32ui.CreateDCFromHandle(hwindc)
                memdc = srcdc.CreateCompatibleDC()
                bmp = win32ui.CreateBitmap()
                bmp.CreateCompatibleBitmap(srcdc, width, height)
                memdc.SelectObject(bmp)
                memdc.BitBlt((0, 0), (width, height), srcdc, (left, top), win32con.SRCCOPY)
                
                # Bitmap 정보를 numpy 배열로 변환
                signedIntsArray = bmp.GetBitmapBits(True)
                img = np.frombuffer(signedIntsArray, dtype='uint8')
                img.shape = (height, width, 4)
                
                # 정리
                srcdc.DeleteDC()
                memdc.DeleteDC()
                win32gui.ReleaseDC(hwin, hwindc)
                win32gui.DeleteObject(bmp.GetHandle())
                
                # BGRA에서 RGB로 변환
                self.screenshot = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)
                
            # macOS 및 Linux
            else:
                # mss 라이브러리를 사용한 방법 (더 빠르고 정확함)
                try:
                    import mss
                    with mss.mss() as sct:
                        monitor = sct.monitors[0]  # 기본 모니터
                        sct_img = sct.grab(monitor)
                        # mss 결과를 numpy 배열로 변환
                        img = np.array(sct_img)
                        # BGRA에서 RGB로 변환
                        self.screenshot = cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)
                        return True
                except ImportError:
                    # PIL ImageGrab을 사용한 대체 방법 (macOS/Linux)
                    try:
                        from PIL import ImageGrab
                        img = np.array(ImageGrab.grab())
                        self.screenshot = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                        return True
                    except ImportError:
                        print("스크린샷 캡처를 위해 mss 또는 PIL 라이브러리가 필요합니다.")
                        print("pip install mss 또는 pip install pillow를 실행하여 설치해주세요.")
                        return False
                    except Exception as e:
                        print(f"PIL ImageGrab 오류: {str(e)}")
                        return False
                except Exception as e:
                    print(f"mss 오류: {str(e)}")
                    return False
            
            return True
            
        except Exception as e:
            print(f"스크린샷 오류: {str(e)}")
            return False
    
    def get_color_at(self, x, y):
        """
        지정된 화면 좌표의 색상 반환
        """
        if self.screenshot is None:
            if not self.update_screenshot():
                return None
        
        try:
            # 스크린샷에서 좌표 값이 유효한지 확인
            h, w = self.screenshot.shape[:2]
            if 0 <= x < w and 0 <= y < h:
                # RGB 순서로 색상 반환
                color = self.screenshot[y, x]
                return (int(color[0]), int(color[1]), int(color[2]))
        except:
            pass
            
        return None

class MunsellIdentifierApp(QMainWindow):
    """
    먼셀 색상 인식 애플리케이션의 메인 윈도우
    """
    def __init__(self):
        super().__init__()
        
        # 애플리케이션 기본 설정
        self.setWindowTitle("먼셀 색상 자동 인식 프로그램")
        self.setGeometry(100, 100, 1000, 700)
        
        # 스크린 컬러 피커 초기화
        self.screen_picker = ScreenColorPicker()
        
        # 색상 히스토리 초기화
        self.color_history = []
        
        # UI 설정
        self.setup_ui()
        
        # 마우스 커서 추적을 위한 타이머 설정
        self.cursor_timer = QTimer(self)
        self.cursor_timer.timeout.connect(self.update_cursor_color)
        self.cursor_timer.start(100)  # 100ms마다 업데이트
        
        # 마우스가 윈도우 위에 있는지 여부
        self.mouse_over_window = False
        
        # 이벤트 필터 설치
        self.installEventFilter(self)
    
    def setup_ui(self):
        """
        UI 컴포넌트 초기화 및 배치
        """
        # 중앙 위젯
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 메인 레이아웃
        main_layout = QVBoxLayout(central_widget)
        
        # 상단 컨트롤 영역
        control_layout = QHBoxLayout()
        
        # 버튼들
        self.open_button = QPushButton("이미지 열기", self)
        self.open_button.clicked.connect(self.open_image)
        
        self.export_button = QPushButton("히스토리 내보내기", self)
        self.export_button.clicked.connect(self.export_history)
        
        self.clear_button = QPushButton("히스토리 지우기", self)
        self.clear_button.clicked.connect(self.clear_history)
        
        # 현재 색상 정보 표시
        self.color_display = MunsellColorDisplay(self)
        
        # 컨트롤 영역에 위젯 추가
        control_layout.addWidget(self.open_button)
        control_layout.addWidget(self.export_button)
        control_layout.addWidget(self.clear_button)
        control_layout.addStretch(1)
        control_layout.addWidget(QLabel("현재 색상:"))
        control_layout.addWidget(self.color_display)
        
        # 메인 레이아웃에 컨트롤 영역 추가
        main_layout.addLayout(control_layout)
        
        # 스플리터 (이미지 뷰어와 히스토리 목록 분리)
        splitter = QSplitter(Qt.Horizontal)
        
        # 이미지 뷰어
        self.image_viewer = ImageViewer(self)
        splitter.addWidget(self.image_viewer)
        
        # 히스토리 영역
        history_widget = QWidget()
        history_layout = QVBoxLayout(history_widget)
        
        # 히스토리 레이블
        history_label = QLabel("색상 히스토리 (클릭으로 저장)")
        history_label.setStyleSheet("font-weight: bold;")
        
        # 히스토리 목록
        self.history_list = QListWidget()
        self.history_list.setAlternatingRowColors(True)
        
        # 히스토리 영역에 위젯 추가
        history_layout.addWidget(history_label)
        history_layout.addWidget(self.history_list)
        
        # 스플리터에 히스토리 영역 추가
        splitter.addWidget(history_widget)
        
        # 스플리터 비율 설정
        splitter.setSizes([700, 300])
        
        # 메인 레이아웃에 스플리터 추가
        main_layout.addWidget(splitter)
        
        # 상태 바
        self.statusBar().showMessage("마우스를 화면 위로 움직여 색상을 확인하세요. 색상을 저장하려면 클릭하세요.")
    
    def open_image(self):
        """
        이미지 파일 선택 및 로드
        """
        file_dialog = QFileDialog()
        image_path, _ = file_dialog.getOpenFileName(
            self, "이미지 파일 열기", "", 
            "이미지 파일 (*.png *.jpg *.jpeg *.bmp *.tif *.tiff);;모든 파일 (*)"
        )
        
        if image_path:
            if self.image_viewer.load_image(image_path):
                self.statusBar().showMessage(f"이미지를 로드했습니다: {os.path.basename(image_path)}")
    
    def add_to_history(self, rgb, munsell_code):
        """
        색상 히스토리에 새 항목 추가
        """
        # 히스토리 객체 생성
        history_item = ColorHistoryItem(rgb, munsell_code)
        self.color_history.append(history_item)
        
        # 리스트 위젯에 항목 추가
        list_item = QListWidgetItem(str(history_item))
        
        # 배경색 설정
        color_hex = rgb_to_hex(rgb)
        text_color = "#000" if sum(rgb) > 384 else "#fff"
        list_item.setBackground(QColor(rgb[0], rgb[1], rgb[2]))
        list_item.setForeground(QColor(text_color))
        
        self.history_list.addItem(list_item)
        
        # 새 항목으로 스크롤
        self.history_list.scrollToItem(list_item)
    
    def export_history(self):
        """
        색상 히스토리를 CSV 파일로 내보내기
        """
        if not self.color_history:
            QMessageBox.information(self, "내보내기", "내보낼 히스토리가 없습니다.")
            return
        
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getSaveFileName(
            self, "히스토리 저장", 
            f"munsell_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", 
            "CSV 파일 (*.csv);;모든 파일 (*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    # 헤더 쓰기
                    writer.writerow(["Timestamp", "Munsell Code", "R", "G", "B", "Hex Color"])
                    
                    # 데이터 쓰기
                    for item in self.color_history:
                        writer.writerow([
                            item.timestamp.strftime('%Y-%m-%d %H:%M:%S'),
                            item.munsell_code,
                            item.rgb[0],
                            item.rgb[1],
                            item.rgb[2],
                            rgb_to_hex(item.rgb)
                        ])
                
                QMessageBox.information(self, "내보내기 성공", f"히스토리를 저장했습니다: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "내보내기 오류", f"히스토리 저장 중 오류가 발생했습니다: {str(e)}")
    
    def clear_history(self):
        """
        색상 히스토리 비우기
        """
        if not self.color_history:
            return
            
        reply = QMessageBox.question(
            self, "히스토리 지우기", 
            "정말로 모든 히스토리를 지우시겠습니까?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.color_history.clear()
            self.history_list.clear()
    
    def update_cursor_color(self):
        """
        현재 마우스 커서 위치의 색상 업데이트
        """
        # 마우스가 윈도우 밖에 있으면 스킵
        if not self.mouse_over_window:
            return
            
        # 현재 마우스 커서 위치 가져오기
        cursor_pos = QCursor.pos()
        
        # 마우스가 이미지 뷰어 위에 있는지 확인
        image_viewer_global_rect = self.image_viewer.rect().translated(
            self.image_viewer.mapToGlobal(QPoint(0, 0))
        )
        
        if image_viewer_global_rect.contains(cursor_pos):
            # 이미지 뷰어 로컬 좌표로 변환
            local_pos = self.image_viewer.mapFromGlobal(cursor_pos)
            rgb = self.image_viewer.get_image_pixel_color(local_pos)
            
            if rgb:
                # 먼셀 코드 찾기
                munsell_code = find_closest_munsell(rgb)
                
                # 색상 정보 업데이트
                self.color_display.update_color(rgb, munsell_code)
                
                # 마우스 커서 옆에 먼셀 코드 표시
                QToolTip.showText(cursor_pos, munsell_code, self)
        else:
            # 전체 화면에서 색상 가져오기
            self.screen_picker.update_screenshot()
            rgb = self.screen_picker.get_color_at(cursor_pos.x(), cursor_pos.y())
            
            if rgb:
                # 먼셀 코드 찾기
                munsell_code = find_closest_munsell(rgb)
                
                # 색상 정보 업데이트
                self.color_display.update_color(rgb, munsell_code)
                
                # 마우스 커서 옆에 먼셀 코드 표시
                QToolTip.showText(cursor_pos, munsell_code, self)
    
    def eventFilter(self, obj, event):
        """
        이벤트 필터 - 마우스 이벤트 처리
        """
        if event.type() == QEvent.Enter:
            self.mouse_over_window = True
        elif event.type() == QEvent.Leave:
            self.mouse_over_window = False
        elif event.type() == QEvent.MouseButtonPress:
            if event.button() == Qt.LeftButton:
                # 이미지 뷰어 클릭 시 색상 저장
                if obj == self.image_viewer and self.image_viewer.original_image:
                    cursor_pos = event.pos()
                    rgb = self.image_viewer.get_image_pixel_color(cursor_pos)
                    
                    if rgb:
                        munsell_code = find_closest_munsell(rgb)
                        self.add_to_history(rgb, munsell_code)
                        
                        return True
        
        return super().eventFilter(obj, event)

def main():
    """
    애플리케이션 실행
    """
    app = QApplication(sys.argv)
    window = MunsellIdentifierApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
