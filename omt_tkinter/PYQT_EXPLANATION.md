# Beginner's Guide to PyQt6 in Marketplace Automation

Welcome! If you are new to PyQt6 or desktop application development in Python, this guide will walk you through the key concepts used in `marketplaces_automation.py`. We have built this application using a "Zero to Hero" approach, meaning it uses highly scalable, professional standards rather than beginner monolithic scripts.

---

## 1. Application Initialization

At the very bottom of `marketplaces_automation.py`, you will see:

```python
def main():
    app = QApplication(sys.argv)
    ...
    window = POReportApp()
    window.show()
    sys.exit(app.exec())
```

- **`QApplication(sys.argv)`**: This object manages the GUI application's control flow and main settings. You must have exactly one `QApplication` instance before creating any graphical elements.
- **`window.show()`**: Makes the main window visible on the screen.
- **`app.exec()`**: Starts the application's event loop. This loop waits for user interactions (like mouse clicks) and tells the application how to respond. `sys.exit()` ensures a clean exit when the application is closed.

---

## 2. Setting Up the Main Window

Our main class, `POReportApp`, inherits from `QMainWindow`, which provides a framework for building an application's main user interface.

```python
self.setWindowTitle("PO Report Generator")
self.setFixedSize(620, 580)
self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
```

- **`setFixedSize(620, 580)`**: Prevents the user from resizing the window. It will always be 620 pixels wide and 580 pixels high.
- **`FramelessWindowHint`**: Removes the default Windows/macOS title bar and borders. This allows us to draw our own custom, "iOS-like" title bar (`DraggableTitleBar`).
- **`WA_TranslucentBackground`**: Makes the background of the window transparent. This is crucial because it allows us to draw a rounded rectangle with drop shadows that "bleed" over the transparent background.

---

## 3. Layout Management

PyQt6 does not use absolute positioning (like x=10, y=20) by default. Instead, it uses Layouts to automatically arrange widgets. We heavily use two types:
- **`QVBoxLayout`**: Arranges widgets vertically (top to bottom).
- **`QHBoxLayout`**: Arranges widgets horizontally (left to right).

### Example from our code:
```python
wrapper_layout = QVBoxLayout(self.wrapper)
wrapper_layout.setContentsMargins(15, 15, 15, 15)
```

- **`setContentsMargins(Left, Top, Right, Bottom)`**: Adds padding inside the layout. Here, we add 15 pixels of space on all sides so our drop shadows don't get cut off by the edge of the window.

---

## 4. Modular UI Assembly (Zero to Hero)

Instead of dumping hundreds of lines of code into one `_build_ui` method, we extract distinct visual areas into helper methods. This is known as "separation of concerns" and makes the code highly readable and scalable.

```python
# Modular UI Assembly - Header Frame
header_frame = self._build_header()
main_layout.addWidget(header_frame)
```

- **`_build_header()`**: Creates and configures the `QFrame` that holds the title and subtitle labels, returning it to be added to the vertical layout.
- **`_build_content()`**: Creates the central card containing the dropdown and the generate button.
- **`_build_footer()`**: Creates the bottom section showing developer information.

---

## 5. Adding Shadows (`QGraphicsDropShadowEffect`)

To give the application depth, we apply shadows to frames and text.

```python
shadow = QGraphicsDropShadowEffect(self)
shadow.setBlurRadius(25)
shadow.setColor(QColor(0, 0, 0, 60))
shadow.setOffset(0, 8)
card_frame.setGraphicsEffect(shadow)
```

- **`setBlurRadius(25)`**: How soft the shadow is.
- **`setColor(...)`**: The color of the shadow. `QColor(Red, Green, Blue, Alpha)` where Alpha is transparency (60 out of 255 makes it very subtle).
- **`setOffset(0, 8)`**: Pushes the shadow 0 pixels right and 8 pixels down.
- **`setGraphicsEffect(...)`**: Applies the shadow to the specific widget (`card_frame`).

---

## 6. Styling with QSS (Qt Style Sheets)

In `_apply_custom_styles()`, we use a syntax very similar to CSS (used in web development) to style our widgets.

```css
#generateBtn {
    background-color: #007AFF; /* iOS Blue */
    color: white;
    border-radius: 14px;
}
```

- **`#generateBtn`**: This targets the specific button where we previously called `self.generate_btn.setObjectName("generateBtn")`.
- This approach completely decouples logic (Python) from presentation (CSS), which is a professional standard.

---

## 7. Background Threading (`QThread`)

When the "Generate Report" button is clicked, reading huge Excel files takes time. If we did this on the main UI thread, the application would freeze and say "Not Responding".

To prevent this, we use `ReportWorker` which inherits from `QThread`:

```python
self.worker = ReportWorker(marketplace, file_path)
self.worker.success.connect(self._handle_process_success)
self.worker.start()
```

- **`QThread`**: Runs the heavy data processing in the background.
- **`pyqtSignal`**: Threads cannot safely update the UI directly. Instead, the `ReportWorker` emits a `success` signal when it finishes.
- **`.connect(...)`**: We link that signal to a function `_handle_process_success`, which runs safely on the main thread to update the UI (like showing a success dialog).

---

By understanding these core concepts, you can easily read, maintain, and scale this application!