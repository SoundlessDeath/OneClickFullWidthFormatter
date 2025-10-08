r"""
Indentor v1 (Windows-only)

Batch-insert two full-width spaces (U+3000×2) at the start of each line (.txt) or each paragraph (.docx).

- Formats: .txt, .docx (NO .doc)
- Naming: if conflict → _indent; if conflict persists → _copy(2), _copy(3) ...
- Encoding (.txt): detect & preserve; preserves per-line line endings
- .docx: inserts after bullets/numbering; skips empty paragraphs; skips Heading styles
- UI: minimal, pretty; shows up to 3 selected files; shows processing & last 3 finished; has "Open output folder"
- Note line: "Please convert .doc to .docx before using this tool!"

Dependencies (install in Windows):
    pip install PySide6 python-docx charset-normalizer pywin32

Pack (optional):
    pyinstaller -F -w -n Indentor_v1_windows indentor_v1_windows.py

"""
from __future__ import annotations
import os
import sys
import re
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple
from datetime import datetime

# --- Third-party ---
from charset_normalizer import from_bytes
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from PySide6 import QtCore, QtGui, QtWidgets

FULL_WIDTH_SPACE = "\u3000"  # U+3000
FW2 = FULL_WIDTH_SPACE * 2

# -----------------------------
# Logging setup
# -----------------------------

def setup_logger(enable_log=False):
    """设置日志记录器"""
    if not enable_log:
        return None
    
    try:
        # 获取脚本所在目录
        script_dir = Path(__file__).parent
        log_dir = script_dir / "Logs"
        
        # 创建Logs目录（如果不存在）
        log_dir.mkdir(exist_ok=True)
        
        # 使用当前日期作为日志文件名
        today = datetime.now().strftime("%Y-%m-%d")
        log_file = log_dir / f"formatter_{today}.log"
        
        # 创建logger
        logger = logging.getLogger('FormatterLogger')
        logger.setLevel(logging.INFO)
        
        # 清除之前的handlers（避免重复）
        logger.handlers.clear()
        
        # 创建文件handler
        handler = logging.FileHandler(log_file, encoding='utf-8', mode='a')
        formatter = logging.Formatter('%(asctime)s - %(message)s', 
                                    datefmt='%Y-%m-%d %H:%M')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        
        return logger
        
    except Exception as e:
        # 如果日志设置失败，返回None，不影响主功能
        return None

# -----------------------------
# Utility: Filename resolution
# -----------------------------

def resolve_output_path(out_dir: Path, src_name: str) -> Path:
    """Return a non-conflicting output path in out_dir based on src_name.
    Strategy:
      - If no conflict: name.ext
      - If conflict: name_indent.ext
      - If conflict: name_indent1.ext, name_indent2.ext, ...
    """
    base = Path(src_name).stem
    ext = Path(src_name).suffix

    candidate = out_dir / f"{base}{ext}"
    if not candidate.exists():
        return candidate

    candidate = out_dir / f"{base}_indent{ext}"
    if not candidate.exists():
        return candidate

    i = 1
    while True:
        candidate = out_dir / f"{base}_indent{i}{ext}"
        if not candidate.exists():
            return candidate
        i += 1

# -----------------------------
# .txt processing
# -----------------------------

def detect_encoding(data: bytes) -> str:
    """Detect encoding with BOM priority; fallback to charset-normalizer; default utf-8."""
    # BOMs
    if data.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if data.startswith((b"\xff\xfe\x00\x00", b"\x00\x00\xfe\xff")):
        # UTF-32 BOMs (check before UTF-16)
        return "utf-32" if data.startswith(b"\xff\xfe\x00\x00") else "utf-32"
    if data.startswith(b"\xff\xfe"):
        return "utf-16-le"
    if data.startswith(b"\xfe\xff"):
        return "utf-16-be"

    # charset-normalizer
    try:
        result = from_bytes(data).best()
        if result and result.encoding:
            return result.encoding
    except Exception:
        pass
    return "utf-8"


def process_txt_file(src: Path, out_dir: Path) -> Path:
    data = src.read_bytes()
    enc = detect_encoding(data)

    # Preserve per-line endings by operating on bytes → decode → splitlines(keepends=True)
    text = data.decode(enc, errors="replace")
    lines = text.splitlines(keepends=True)

    def transform_line(line: str) -> str:
        # Separate content and line ending
        if line.endswith("\r\n"):
            core, eol = line[:-2], "\r\n"
        elif line.endswith("\n"):
            core, eol = line[:-1], "\n"
        elif line.endswith("\r"):
            core, eol = line[:-1], "\r"
        else:
            core, eol = line, ""

        if core.strip() == "":
            return core + eol
        # Idempotent: ensure startswith 2 full-width spaces
        if core.startswith(FW2):
            return core + eol
        # If there are ASCII spaces/tabs at start, normalize to two full-width spaces
        core_no_leading = core.lstrip(" \t")
        return FW2 + core_no_leading + eol

    new_text = "".join(transform_line(l) for l in lines)

    out_path = resolve_output_path(out_dir, src.name)
    out_path.write_text(new_text, encoding=enc, errors="replace", newline="")
    return out_path

# -----------------------------
# .docx processing
# -----------------------------

HEADING_KEYWORDS = {"Heading", "标题", "Заголовок", "Überschrift", "Rubrik"}  # common locales


def is_heading_style(p: Paragraph) -> bool:
    try:
        name = p.style.name or ""
    except Exception:
        name = ""
    return any(k in name for k in HEADING_KEYWORDS)


def paragraph_has_text(p: Paragraph) -> bool:
    return (p.text or "").strip() != ""


def ensure_fw2_at_paragraph_start(p: Paragraph) -> None:
    """Prepend two full-width spaces to paragraph *content* if not present.
    - Skips empty paragraphs
    - Skips heading styles
    - Inserts after numbering/bullet (since numbering is not part of runs)
    - Idempotent: do nothing if already startswith FW2
    - Try to preserve existing runs by editing only the first run
    - Handle soft line breaks by transforming each line within run texts
    """
    if not paragraph_has_text(p):
        return
    if is_heading_style(p):
        return

    # Quick idempotency check on full paragraph text
    full = p.text
    if full.startswith(FW2):
        return

    # Ensure we only modify the very start of the paragraph content.
    if p.runs:
        first_run = p.runs[0]
        first_run.text = FW2 + first_run.text.lstrip(" \t")
    else:
        # Paragraph has no runs (rare). Set text directly.
        p.add_run(FW2)

    # Soft line breaks inside runs: add FW2 after each break if non-empty follows
    for run in p.runs:
        if "\n" in run.text:
            parts = run.text.split("\n")
            for i in range(1, len(parts)):
                if parts[i] and not parts[i].startswith(FW2):
                    parts[i] = FW2 + parts[i].lstrip(" \t")
            run.text = "\n".join(parts)


def process_docx_file(src: Path, out_dir: Path) -> Path:
    doc = Document(str(src))
    for p in doc.paragraphs:
        ensure_fw2_at_paragraph_start(p)
    # Tables: also iterate cells' paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    ensure_fw2_at_paragraph_start(p)
    out_path = resolve_output_path(out_dir, src.name)
    doc.save(str(out_path))
    return out_path

# -----------------------------
# Worker thread
# -----------------------------

@dataclass
class TaskResult:
    src: Path
    dst: Optional[Path]
    ok: bool
    message: str


class ProcessorWorker(QtCore.QThread):
    progress = QtCore.Signal(object)  # TaskResult per file
    finished_all = QtCore.Signal()

    def __init__(self, files: List[Path], out_dir: Path, logger=None):
        super().__init__()
        self.files = files
        self.out_dir = out_dir
        self.logger = logger
        self.successful_outputs = []
        self.errors = []

    def run(self):
        # 记录转换开始
        if self.logger:
            self.logger.info("=" * 50)
            self.logger.info("转换开始")
            self.logger.info("转换列表：")
            for file_path in self.files:
                self.logger.info(f"  {file_path}")
            self.logger.info(f"输出目录: {self.out_dir}")
        
        for f in self.files:
            try:
                ext = f.suffix.lower()
                if ext == ".txt":
                    dst = process_txt_file(f, self.out_dir)
                elif ext == ".docx":
                    dst = process_docx_file(f, self.out_dir)
                else:
                    raise ValueError("Unsupported extension")
                
                self.successful_outputs.append(dst)
                self.progress.emit(TaskResult(f, dst, True, "Done"))
                
            except Exception as e:
                error_msg = f"{f.name}: {str(e)}"
                self.errors.append(error_msg)
                self.progress.emit(TaskResult(f, None, False, str(e)))
        
        # 记录最终结果
        if self.logger:
            self.logger.info("输出列表：")
            for output_path in self.successful_outputs:
                self.logger.info(f"  {output_path}")
            
            if self.errors:
                self.logger.info("错误记录：")
                for error in self.errors:
                    self.logger.info(f"  {error}")
            
            success_count = len(self.successful_outputs)
            error_count = len(self.errors)
            self.logger.info(f"转换完成 - 成功: {success_count}, 失败: {error_count}")
        
        self.finished_all.emit()

# -----------------------------
# Custom Output Directory Dialog
# -----------------------------

class OutputDirDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, initial_path=""):
        super().__init__(parent)
        self.setWindowTitle("选择输出文件夹")
        self.setModal(True)
        self.resize(750, 500)  # Increased size for better usability
        self.selected_path = ""
        
        # Set initial directory
        self.current_path = Path(initial_path) if initial_path and Path(initial_path).exists() else Path.cwd()
        
        self._setup_ui()
        self._apply_style()
        self._load_directory()
        
    def _setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)
        
        # Path bar with dropdown
        path_layout = QtWidgets.QHBoxLayout()
        
        # Quick access dropdown
        self.quick_combo = QtWidgets.QComboBox()
        self.quick_combo.setMinimumWidth(100)  # Reduced from 120
        self.quick_combo.setMaximumWidth(140)  # Set maximum width to prevent it from expanding too much
        self._populate_quick_access()
        self.quick_combo.currentTextChanged.connect(self._on_quick_access_changed)
        path_layout.addWidget(self.quick_combo)
        
        path_layout.addWidget(QtWidgets.QLabel("当前路径:"))
        
        self.path_edit = QtWidgets.QLineEdit()
        self.path_edit.setText(str(self.current_path))
        self.path_edit.returnPressed.connect(self._navigate_to_path)
        path_layout.addWidget(self.path_edit, 1)
        
        self.btn_up = QtWidgets.QPushButton("上级")
        self.btn_up.clicked.connect(self._go_up)
        path_layout.addWidget(self.btn_up)
        
        layout.addLayout(path_layout)
        
        # File list
        self.file_list = QtWidgets.QListWidget()
        self.file_list.itemDoubleClicked.connect(self._item_double_clicked)
        layout.addWidget(self.file_list)
        
        # Button layout
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addStretch()
        
        self.btn_select = QtWidgets.QPushButton("选择文件夹")
        self.btn_select.clicked.connect(self._select_current)
        button_layout.addWidget(self.btn_select)
        
        self.btn_cancel = QtWidgets.QPushButton("取消")
        self.btn_cancel.clicked.connect(self.reject)
        button_layout.addWidget(self.btn_cancel)
        
        layout.addLayout(button_layout)
        
    def _populate_quick_access(self):
        """Populate quick access dropdown with drives and special folders"""
        self.quick_combo.addItem("选择地址…", "")
        
        # Add desktop
        try:
            desktop_path = Path.home() / "Desktop"
            if desktop_path.exists():
                self.quick_combo.addItem("🖥️ 桌面", str(desktop_path))
        except:
            pass
        
        # Add available drives
        import string
        for drive_letter in string.ascii_uppercase:
            drive_path = Path(f"{drive_letter}:\\")
            if drive_path.exists():
                self.quick_combo.addItem(f"🗃️ {drive_letter}盘", str(drive_path))
        
        # Add Windows Quick Access using pywin32
        quick_access_items = []
        try:
            import win32com.client
            shell = win32com.client.Dispatch("Shell.Application")
            # Get Quick Access folder (FOLDERID_QuickAccess)
            quick_access = shell.Namespace("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
            if quick_access:
                for item in quick_access.Items():
                    try:
                        path_str = item.Path
                        if path_str and Path(path_str).exists() and Path(path_str).is_dir():
                            # Skip if it's already added (like Desktop)
                            already_added = False
                            for i in range(1, self.quick_combo.count()):
                                if self.quick_combo.itemData(i) == path_str:
                                    already_added = True
                                    break
                            if not already_added:
                                display_name = item.Name or Path(path_str).name
                                quick_access_items.append((f"↘️ {display_name}", path_str))
                    except:
                        continue
        except ImportError:
            # pywin32 not available, skip Quick Access
            pass
        except Exception:
            # Other errors, skip Quick Access
            pass
        
        # Add separator and quick access items if any exist
        if quick_access_items:
            # Add separator (non-clickable)
            self.quick_combo.addItem("快速访问", "")
            # Disable the separator item
            separator_index = self.quick_combo.count() - 1
            separator_item = self.quick_combo.model().item(separator_index)
            separator_item.setEnabled(False)
            
            # Add the actual quick access items
            for display_name, path_str in quick_access_items:
                self.quick_combo.addItem(display_name, path_str)
        
    def _on_quick_access_changed(self, text):
        """Handle quick access selection"""
        if text == "选择地址…" or text == "快速访问":
            return
            
        # Get the path from combo data
        current_index = self.quick_combo.currentIndex()
        if current_index > 0:  # Skip the first "选择地址…" item
            path_str = self.quick_combo.itemData(current_index)
            if path_str:
                try:
                    new_path = Path(path_str)
                    if new_path.exists() and new_path.is_dir():
                        self.current_path = new_path
                        self._load_directory()
                except Exception:
                    pass
        
        # Reset combo to "选择地址…"
        self.quick_combo.setCurrentIndex(0)
        
    def _apply_style(self):
        self.setStyleSheet("""
            QDialog { background: white; }
            QLineEdit { 
                padding: 8px 10px; 
                border: 1px solid #ddd; 
                border-radius: 8px; 
                font-size: 14px;
            }
            QComboBox {
                padding: 6px 8px;
                border: 1px solid #ddd;
                border-radius: 8px;
                background: #fafafa;
                font-size: 14px;
            }
            QComboBox:hover {
                background: #f2f2f2;
            }
            QComboBox::drop-down {
                border: none;
                padding-right: 8px;
            }
            QComboBox::down-arrow {
                image: none;
                border: 2px solid #666;
                border-top: none;
                border-left: none;
                width: 6px;
                height: 6px;
                transform: rotate(45deg);
                margin-top: -3px;
            }
            QComboBox QAbstractItemView::item:disabled {
                color: #999;
                background: #f8f8f8;
            }
            QComboBox QAbstractItemView::item:disabled:hover {
                background: #f8f8f8;
            }
            QPushButton { 
                padding: 8px 15px; 
                border: 1px solid #ddd; 
                border-radius: 8px; 
                background: #fafafa; 
                min-width: 70px;
                font-size: 14px;
            }
            QPushButton:hover { background: #f2f2f2; }
            QListWidget { 
                border: 1px solid #ddd; 
                border-radius: 8px; 
                background: white;
                font-size: 14px;
            }
            QListWidget::item {
                padding: 6px 10px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:hover {
                background: #f5f5f5;
            }
            QListWidget::item:selected {
                background: #e3f2fd;
                color: black;
            }
            QLabel {
                font-size: 14px;
            }
        """)
        
    def _load_directory(self):
        self.file_list.clear()
        
        try:
            if not self.current_path.exists():
                self.current_path = Path.cwd()
                
            self.path_edit.setText(str(self.current_path))
            
            # Add parent directory item (if not root)
            if self.current_path.parent != self.current_path:
                item = QtWidgets.QListWidgetItem("📁 ..")
                item.setData(QtCore.Qt.UserRole, str(self.current_path.parent))
                self.file_list.addItem(item)
            
            # Get all items in directory
            items = []
            try:
                for path in self.current_path.iterdir():
                    if path.is_dir():
                        items.append((f"📁 {path.name}", str(path), True))
                    elif path.suffix.lower() in {'.txt', '.docx'}:
                        icon = "📄" if path.suffix.lower() == '.txt' else "📝"
                        items.append((f"{icon} {path.name}", str(path), False))
            except PermissionError:
                item = QtWidgets.QListWidgetItem("❌ 无法访问此目录")
                self.file_list.addItem(item)
                return
                
            # Sort: directories first, then files
            items.sort(key=lambda x: (not x[2], x[0].lower()))
            
            # Add items to list
            for display_name, full_path, is_dir in items:
                item = QtWidgets.QListWidgetItem(display_name)
                item.setData(QtCore.Qt.UserRole, full_path)
                self.file_list.addItem(item)
                
        except Exception as e:
            item = QtWidgets.QListWidgetItem(f"❌ 错误: {str(e)}")
            self.file_list.addItem(item)
            
    def _navigate_to_path(self):
        path_text = self.path_edit.text().strip()
        try:
            new_path = Path(path_text)
            if new_path.exists() and new_path.is_dir():
                self.current_path = new_path
                self._load_directory()
        except Exception:
            # Reset to current path if invalid
            self.path_edit.setText(str(self.current_path))
            
    def _go_up(self):
        if self.current_path.parent != self.current_path:
            self.current_path = self.current_path.parent
            self._load_directory()
            
    def _item_double_clicked(self, item):
        path_str = item.data(QtCore.Qt.UserRole)
        if path_str:
            path = Path(path_str)
            if path.exists() and path.is_dir():
                self.current_path = path
                self._load_directory()
                
    def _select_current(self):
        self.selected_path = str(self.current_path)
        self.accept()
        
    def get_selected_path(self):
        return self.selected_path

# -----------------------------
# GUI
# -----------------------------

class IndentorApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Indentor v1 · Windows")
        self.setMinimumWidth(800)  # Increased width to accommodate longer text
        self.files: List[Path] = []
        self.out_dir: Optional[Path] = None
        self.worker: Optional[ProcessorWorker] = None
        self._build_ui()
        self._apply_style()

    def _build_ui(self):
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        # Title
        title = QtWidgets.QLabel("批量首行缩进（仅支持txt/docx格式）")
        title.setStyleSheet("font-size:20px; font-weight:600;")
        layout.addWidget(title)

        # Split note into two lines
        note1 = QtWidgets.QLabel("请将 .doc 文件先转为 .docx 文件，才能使用本脚本！")
        note1.setStyleSheet("color:#666;")
        note1.setAlignment(QtCore.Qt.AlignLeft)
        layout.addWidget(note1)
        
        note2 = QtWidgets.QLabel("使用前建议先清除原有排版（仅保留分段）")
        note2.setStyleSheet("color:#666;")
        note2.setAlignment(QtCore.Qt.AlignLeft)
        layout.addWidget(note2)

        # Multi-path checkbox row
        checkbox_row = QtWidgets.QHBoxLayout()
        
        self.chk_clear_list = QtWidgets.QCheckBox("清除转换列表")
        self.chk_clear_list.setChecked(True)  # Default checked
        self.chk_clear_list.setStyleSheet("color:#666;")
        checkbox_row.addWidget(self.chk_clear_list)
        
        # Add spacing between checkboxes
        checkbox_row.addSpacing(20)
        
        self.chk_multipath = QtWidgets.QCheckBox("多路径文件")
        self.chk_multipath.setStyleSheet("color:#666;")
        checkbox_row.addWidget(self.chk_multipath)
        
        # Add spacing and log checkbox
        checkbox_row.addSpacing(20)
        self.chk_enable_log = QtWidgets.QCheckBox("记录日志")
        self.chk_enable_log.setStyleSheet("color:#666;")
        checkbox_row.addWidget(self.chk_enable_log)
        
        checkbox_row.addStretch()  # Push checkboxes to the left
        layout.addLayout(checkbox_row)

        # File picker row
        file_row = QtWidgets.QHBoxLayout()
        self.btn_pick = QtWidgets.QPushButton("选择文件（可多选）…")
        self.btn_pick.clicked.connect(self.pick_files)
        file_row.addWidget(self.btn_pick)
        
        self.btn_clear = QtWidgets.QPushButton("清除")
        self.btn_clear.clicked.connect(self.clear_files)
        file_row.addWidget(self.btn_clear)

        self.sel_summary = QtWidgets.QLabel("未选择文件")
        self.sel_summary.setStyleSheet("color:#444;")
        self.sel_summary.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)
        file_row.addWidget(self.sel_summary, 1)
        layout.addLayout(file_row)

        # Output dir row
        out_row = QtWidgets.QHBoxLayout()
        self.out_edit = QtWidgets.QLineEdit()
        self.out_edit.setPlaceholderText("输出文件夹路径（可直接输入或右侧选择）")
        out_row.addWidget(self.out_edit, 1)
        self.btn_out = QtWidgets.QPushButton("输出地址…")
        self.btn_out.clicked.connect(self.pick_out_dir)
        out_row.addWidget(self.btn_out)
        layout.addLayout(out_row)

        # Start button
        self.btn_start = QtWidgets.QPushButton("开始")
        self.btn_start.setFixedHeight(36)
        self.btn_start.setStyleSheet("font-weight:600;")
        self.btn_start.clicked.connect(self.start_processing)
        layout.addWidget(self.btn_start)

        # Status area
        grp = QtWidgets.QGroupBox("状态")
        gl = QtWidgets.QGridLayout(grp)
        gl.setContentsMargins(12, 12, 12, 12)
        gl.setHorizontalSpacing(10)
        gl.setVerticalSpacing(6)

        self.lbl_processing = QtWidgets.QLabel("正在处理：－")
        self.lbl_done = QtWidgets.QLabel("已完成：－")
        self.lbl_err = QtWidgets.QLabel("错误：－")
        for w in (self.lbl_processing, self.lbl_done, self.lbl_err):
            w.setTextInteractionFlags(QtCore.Qt.TextSelectableByMouse)

        gl.addWidget(self.lbl_processing, 0, 0)
        gl.addWidget(self.lbl_done, 1, 0)
        gl.addWidget(self.lbl_err, 2, 0)

        self.btn_open_out = QtWidgets.QPushButton("打开输出文件夹")
        self.btn_open_out.clicked.connect(self.open_out_dir)
        gl.addWidget(self.btn_open_out, 3, 0, 1, 1)

        layout.addWidget(grp)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

    def _apply_style(self):
        # Minimal, modern, light theme
        self.setStyleSheet(
            """
            QWidget { font-size:14px; }
            QPushButton { padding:6px 12px; border:1px solid #ddd; border-radius:8px; background:#fafafa; }
            QPushButton:hover { background:#f2f2f2; }
            QLineEdit { padding:6px 8px; border:1px solid #ddd; border-radius:8px; }
            QGroupBox { border:1px solid #eee; border-radius:10px; margin-top:10px; }
            QGroupBox::title { subcontrol-origin: margin; left:10px; padding:0 4px; color:#666; }
            QProgressBar { height:12px; border:1px solid #ddd; border-radius:6px; background:#f5f5f5; }
            QProgressBar::chunk { border-radius:6px; background:#6aa3ff; }
            """
        )

    # ---------------- Events ----------------
    def pick_files(self):
        dlg = QtWidgets.QFileDialog(self, "选择文件")
        dlg.setFileMode(QtWidgets.QFileDialog.ExistingFiles)
        # Only one combined filter, no "All files" option
        dlg.setNameFilters(["文本与 Word (*.txt *.docx)"])
        dlg.selectNameFilter("文本与 Word (*.txt *.docx)")
        if dlg.exec():
            selected = [Path(p) for p in dlg.selectedFiles()]
            # Filter to .txt / .docx just in case
            selected = [p for p in selected if p.suffix.lower() in {'.txt', '.docx'}]
            
            if self.chk_multipath.isChecked():
                # Remove duplicates: check if file already exists in current list
                existing_paths = {p.resolve() for p in self.files}
                new_files = [p for p in selected if p.resolve() not in existing_paths]
                self.files.extend(new_files)
            else:
                self.files = selected
            self.update_selected_summary()

    def update_selected_summary(self):
        if not self.files:
            self.sel_summary.setText("未选择文件")
            return
        paths = [str(p) for p in self.files[:3]]
        extra = len(self.files) - 3
        text = "\n".join(paths)
        if extra > 0:
            text += f"\n……等 {extra} 个"
        self.sel_summary.setText(text)

    def pick_out_dir(self):
        dlg = OutputDirDialog(self, self.out_edit.text().strip())
        if dlg.exec():
            self.out_edit.setText(dlg.get_selected_path())

    def open_out_dir(self):
        path = self.out_edit.text().strip()
        if not path:
            QtWidgets.QMessageBox.information(self, "提示", "请先选择输出文件夹")
            return
        if not Path(path).exists():
            QtWidgets.QMessageBox.warning(self, "提示", "输出文件夹不存在")
            return
        # Open in Explorer with error handling
        try:
            os.startfile(path)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "错误", f"无法打开文件夹：{str(e)}")

    def start_processing(self):
        if not self.files:
            QtWidgets.QMessageBox.information(self, "提示", "请先选择文件（.txt / .docx）")
            return
        out_dir = self.out_edit.text().strip()
        if not out_dir:
            QtWidgets.QMessageBox.information(self, "提示", "请先选择或输入输出文件夹")
            return
        out_path = Path(out_dir)
        if not out_path.exists():
            QtWidgets.QMessageBox.information(self, "提示", "输出文件夹不存在")
            return
        if not os.access(out_path, os.W_OK):
            QtWidgets.QMessageBox.warning(self, "提示", "没有输出文件夹的写入权限")
            return

        # Clean up previous worker if exists
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()

        # 设置日志
        logger = setup_logger(self.chk_enable_log.isChecked())
        if self.chk_enable_log.isChecked() and logger is None:
            # 日志设置失败但用户想要日志，给出警告
            reply = QtWidgets.QMessageBox.question(
                self, "日志警告", 
                "日志功能初始化失败，可能是权限问题。\n是否继续不记录日志？",
                QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No
            )
            if reply == QtWidgets.QMessageBox.No:
                return

        self.btn_start.setEnabled(False)
        self.btn_pick.setEnabled(False)
        self.btn_out.setEnabled(False)
        self.progress_bar.setValue(0)
        self.lbl_processing.setText("正在处理：－")
        self.lbl_done.setText("已完成：－")
        self.lbl_err.setText("错误：－")
        
        # Reset completion tracking
        self._completed_names = []
        self._error_msgs = []

        self.worker = ProcessorWorker(self.files, out_path, logger)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished_all.connect(self.on_finished_all)
        self.worker.start()

    def on_progress(self, result: TaskResult):
        # Update progress bar & labels
        try:
            idx = self.files.index(result.src)
            pct = int((idx + 1) / max(1, len(self.files)) * 100)
            self.progress_bar.setValue(pct)
        except ValueError:
            # File not in list, shouldn't happen but handle gracefully
            pass

        # Show current processing file
        if idx < len(self.files) - 1:  # Not the last file
            self.lbl_processing.setText("正在处理：" + result.src.name)
        else:
            self.lbl_processing.setText("正在处理：完成")

        # Completed list (last up to 3)
        if not hasattr(self, "_completed_names"):
            self._completed_names = []
        if result.ok:
            self._completed_names.append(result.src.name)
            self._completed_names = self._completed_names[-3:]
            self.lbl_done.setText("已完成：" + ", ".join(self._completed_names))
        else:
            # Error list accumulate (last up to 3)
            if not hasattr(self, "_error_msgs"):
                self._error_msgs = []
            self._error_msgs.append(f"{result.src.name}: {result.message}")
            self._error_msgs = self._error_msgs[-3:]
            self.lbl_err.setText("错误：" + " | ".join(self._error_msgs))

    def on_finished_all(self):
        self.lbl_processing.setText("正在处理：全部完成")
        self.btn_start.setEnabled(True)
        self.btn_pick.setEnabled(True)
        self.btn_out.setEnabled(True)
        
        # Clear file list if option is checked
        if self.chk_clear_list.isChecked():
            self.files = []
            self.update_selected_summary()
        
        # Clean up worker
        if self.worker:
            self.worker.deleteLater()
            self.worker = None
            
        QtWidgets.QMessageBox.information(self, "完成", "处理完毕！")

    def clear_files(self):
        self.files = []
        self.update_selected_summary()


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = IndentorApp()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
