from PyQt6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QFrame, QApplication, QWidget, QGraphicsDropShadowEffect
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QPropertyAnimation, QRect
from PyQt6.QtGui import QFont, QColor

class SmoothDialog(QDialog):
    """An ultra-smooth, iOS-like custom dialog to replace the standard QMessageBox."""
    def __init__(self, parent=None, title="Notification", message="", dialog_type="info"):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Dialog)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.result_value = False # Used for Yes/No dialogs

        # Main layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)

        # Background frame for rounding and shadow
        self.bg_frame = QFrame(self)
        self.bg_frame.setObjectName("dialogBg")

        # Add shadow
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(25)
        shadow.setColor(QColor(0, 0, 0, 80))
        shadow.setOffset(0, 4)
        self.bg_frame.setGraphicsEffect(shadow)

        bg_layout = QVBoxLayout(self.bg_frame)
        bg_layout.setContentsMargins(25, 25, 25, 20)
        bg_layout.setSpacing(15)

        # Title
        self.title_lbl = QLabel(title)
        self.title_lbl.setObjectName("dialogTitle")
        self.title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        # Message
        self.msg_lbl = QLabel(message)
        self.msg_lbl.setObjectName("dialogMessage")
        self.msg_lbl.setWordWrap(True)
        self.msg_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

        bg_layout.addWidget(self.title_lbl)
        bg_layout.addWidget(self.msg_lbl)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        if dialog_type in ["info", "error", "warning", "success"]:
            self.ok_btn = QPushButton("OK")
            self.ok_btn.setObjectName("dialogBtnPrimary")
            self.ok_btn.clicked.connect(self.accept)
            self.ok_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_layout.addWidget(self.ok_btn)

            # Auto-focus OK
            self.ok_btn.setFocus()

        elif dialog_type == "question":
            self.no_btn = QPushButton("Cancel")
            self.no_btn.setObjectName("dialogBtnSecondary")
            self.no_btn.clicked.connect(self.reject)
            self.no_btn.setCursor(Qt.CursorShape.PointingHandCursor)

            self.yes_btn = QPushButton("Yes")
            self.yes_btn.setObjectName("dialogBtnPrimary")
            self.yes_btn.clicked.connect(self._accept_yes)
            self.yes_btn.setCursor(Qt.CursorShape.PointingHandCursor)

            btn_layout.addWidget(self.no_btn)
            btn_layout.addWidget(self.yes_btn)

            self.yes_btn.setFocus()

        bg_layout.addLayout(btn_layout)
        layout.addWidget(self.bg_frame)

        self.setStyleSheet("""
            #dialogBg {
                background-color: #FDFDFD;
                border-radius: 18px;
            }
            #dialogTitle {
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
                font-size: 18px;
                font-weight: bold;
                color: #111827;
            }
            #dialogMessage {
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
                font-size: 14px;
                color: #4B5563;
                margin-bottom: 10px;
            }
            QPushButton {
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
                font-size: 14px;
                font-weight: 600;
                padding: 10px 20px;
                border-radius: 10px;
                border: none;
                min-width: 80px;
            }
            #dialogBtnPrimary {
                background-color: #007AFF; /* iOS Blue */
                color: white;
            }
            #dialogBtnPrimary:hover {
                background-color: #0066D6;
            }
            #dialogBtnPrimary:pressed {
                background-color: #0052AC;
            }
            #dialogBtnSecondary {
                background-color: #F3F4F6;
                color: #374151;
            }
            #dialogBtnSecondary:hover {
                background-color: #E5E7EB;
            }
            #dialogBtnSecondary:pressed {
                background-color: #D1D5DB;
            }
        """)

        # Add fade in animation
        self.setWindowOpacity(0.0)
        self.anim = QPropertyAnimation(self, b"windowOpacity")
        self.anim.setDuration(150)
        self.anim.setStartValue(0.0)
        self.anim.setEndValue(1.0)

    def showEvent(self, event):
        self.anim.start()
        super().showEvent(event)

    def _accept_yes(self):
        self.result_value = True
        self.accept()

    @staticmethod
    def show_info(parent, title, message):
        dlg = SmoothDialog(parent, title, message, "info")
        dlg.exec()

    @staticmethod
    def show_error(parent, title, message):
        # We could change colors here if we wanted an error theme
        dlg = SmoothDialog(parent, title, message, "error")
        dlg.exec()

    @staticmethod
    def show_warning(parent, title, message):
        dlg = SmoothDialog(parent, title, message, "warning")
        dlg.exec()

    @staticmethod
    def ask_question(parent, title, message):
        dlg = SmoothDialog(parent, title, message, "question")
        dlg.exec()
        return dlg.result_value
