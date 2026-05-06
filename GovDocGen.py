import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import json
import os
import subprocess
import platform
import shutil
import time
import importlib.util
import sys
import traceback
from pathlib import Path
from dataclasses import dataclass, asdict
from typing import Optional, Tuple, Dict, Any, List
from PIL import Image, ImageTk
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# Импорты с проверкой наличия
try:
    from docxtpl import DocxTemplate, InlineImage
    from docx.shared import Mm
except ImportError:
    messagebox.showerror("Ошибка", "Установите docxtpl: pip install docxtpl")
    sys.exit(1)

try:
    from pdf2image import convert_from_path
    from PIL import Image
except ImportError:
    messagebox.showerror("Ошибка", "Установите pdf2image и pillow: pip install pdf2image pillow")
    sys.exit(1)

try:
    import comtypes.client
    import pythoncom
    WORD_AVAILABLE = True
except ImportError:
    WORD_AVAILABLE = False

# Константы
APP_TITLE = "Система подготовки документов"
APP_SIZE = "700x650"
MIN_SIZE = "700x650"
DISCLAIMER = "Данная программа - эмулятор, была создана для игрового проекта RMRP. Все совпадения с реальной жизнью случайны."
VERSION = "1.3"

# Пути
BASE_DIR = Path.cwd()
TEMPLATES_DIR = BASE_DIR / "TEMPLATES"
MODULES_DIR = BASE_DIR / "MODULES"
COMPLETE_DIR = BASE_DIR / "COMPLETE"
IMAGES_DIR = BASE_DIR / "IMAGES"
LOGS_DIR = BASE_DIR / "LOGS"
POPPLER_DIR = BASE_DIR / "poppler" / "bin"
DATA_FILE = BASE_DIR / "gov_data.json"
HELP_FILE = BASE_DIR / "HELP.txt"
CANVAS_IMAGE_PATH = BASE_DIR / "CANVAS.jpg"

# Создаем папки
for dir_path in [COMPLETE_DIR, IMAGES_DIR, MODULES_DIR, TEMPLATES_DIR, LOGS_DIR]:
    dir_path.mkdir(exist_ok=True)

# Месяцы
MONTHS = {
    1: "января", 2: "февраля", 3: "марта", 4: "апреля",
    5: "мая", 6: "июня", 7: "июля", 8: "августа",
    9: "сентября", 10: "октября", 11: "ноября", 12: "декабря"
}

@dataclass
class GovData:
    last_order_number: int = 0
    last_signer: str = ""
    modules_data: Dict[str, Any] = None
    
    def __post_init__(self):
        if self.modules_data is None:
            self.modules_data = {}

class DataManager:
    def __init__(self):
        self.data_file = DATA_FILE
        self.data = self.load()
    
    def load(self) -> GovData:
        if not self.data_file.exists():
            return GovData()
        try:
            with open(self.data_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return GovData(
                    last_order_number=data.get('last_order_number', 0),
                    last_signer=data.get('last_signer', ''),
                    modules_data=data.get('modules_data', {})
                )
        except:
            return GovData()
    
    def save(self, data: GovData):
        try:
            with open(self.data_file, 'w', encoding='utf-8') as f:
                json.dump(asdict(data), f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log_error(f"Ошибка сохранения данных: {e}")
    
    @staticmethod
    def log_error(error_msg: str):
        """Запись ошибки в лог-файл"""
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_file = LOGS_DIR / f"error_{timestamp}.log"
            with open(log_file, 'w', encoding='utf-8') as f:
                f.write(f"Ошибка: {error_msg}\n")
                f.write(f"Время: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(traceback.format_exc())
        except:
            pass

class ConversionChecker:
    @staticmethod
    def check_word() -> bool:
        if platform.system() == 'Windows':
            try:
                import winreg
                key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application\\CurVer")
                winreg.QueryValueEx(key, "")
                winreg.CloseKey(key)
                return True
            except:
                pass
        if not WORD_AVAILABLE:
            return False
        try:
            pythoncom.CoInitialize()
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            word.Quit()
            pythoncom.CoUninitialize()
            return True
        except:
            return False
    
    @staticmethod
    def check_libreoffice() -> Optional[str]:
        if platform.system() == 'Windows':
            paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            ]
            for path in paths:
                if os.path.exists(path):
                    return path
        elif platform.system() == 'Linux':
            if shutil.which('libreoffice'):
                return 'libreoffice'
        elif platform.system() == 'Darwin':
            mac_path = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
            if os.path.exists(mac_path):
                return mac_path
        return None
    
    @staticmethod
    def check_poppler() -> bool:
        if platform.system() == 'Windows':
            return (POPPLER_DIR / "pdftoppm.exe").exists()
        return shutil.which('pdftoppm') is not None

class DocumentConverter:
    def __init__(self, word_available: bool, libreoffice_path: Optional[str]):
        self.word_available = word_available
        self.libreoffice_path = libreoffice_path
    
    def docx_to_pdf(self, docx_path: Path, pdf_path: Path) -> Tuple[bool, str]:
        if self.word_available:
            success, msg = self._word_convert(docx_path, pdf_path)
            if success:
                return True, "Word"
            if self.libreoffice_path:
                success, msg = self._libreoffice_convert(docx_path, pdf_path)
                if success:
                    return True, "LibreOffice (Word не сработал)"
        elif self.libreoffice_path:
            success, msg = self._libreoffice_convert(docx_path, pdf_path)
            if success:
                return True, "LibreOffice"
        
        return False, "Нет доступных методов конвертации"
    
    def _word_convert(self, docx_path: Path, pdf_path: Path) -> Tuple[bool, str]:
        try:
            pythoncom.CoInitialize()
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path))
            doc.SaveAs(str(pdf_path), FileFormat=17)
            doc.Close()
            word.Quit()
            pythoncom.CoUninitialize()
            return True, ""
        except Exception as e:
            return False, str(e)
    
    def _libreoffice_convert(self, docx_path: Path, pdf_path: Path) -> Tuple[bool, str]:
        try:
            cmd = [
                self.libreoffice_path,
                '--headless',
                '--convert-to', 'pdf:writer_pdf_Export',
                '--infilter="Microsoft Word 2007 XML"',
                '--outdir', str(pdf_path.parent),
                str(docx_path)
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            
            if result.returncode == 0:
                generated = pdf_path.parent / (docx_path.stem + ".pdf")
                time.sleep(2)
                
                if generated.exists():
                    if generated != pdf_path:
                        if pdf_path.exists():
                            os.remove(str(pdf_path))
                        os.rename(str(generated), str(pdf_path))
                    return True, ""
                else:
                    return False, "PDF не создан"
            else:
                return False, f"Ошибка LibreOffice: {result.stderr}"
                
        except subprocess.TimeoutExpired:
            return False, "Таймаут при конвертации (120 сек)"
        except Exception as e:
            return False, str(e)
    
    def pdf_to_jpg_multipage(self, pdf_path: Path, output_dir: Path, base_filename: str) -> List[Path]:
        try:
            if not ConversionChecker.check_poppler():
                return []
            
            params = {}
            if platform.system() == 'Windows' and POPPLER_DIR.exists():
                params['poppler_path'] = str(POPPLER_DIR)
            
            images = convert_from_path(str(pdf_path), dpi=300, **params)
            result_files = []
            
            for i, image in enumerate(images):
                page_num = i + 1
                if len(images) == 1:
                    jpg_path = output_dir / f"{base_filename}.jpg"
                else:
                    jpg_path = output_dir / f"{base_filename}_страница_{page_num}.jpg"
                image.save(str(jpg_path), 'JPEG', quality=95)
                result_files.append(jpg_path)
            
            if pdf_path.exists():
                os.remove(str(pdf_path))
            
            return result_files
        except Exception as e:
            DataManager.log_error(f"Ошибка конвертации PDF в JPG: {e}")
            return []

class DocumentGenerator:
    def __init__(self, converter: DocumentConverter):
        self.converter = converter
    
    def create_docx(self, template_path: Path, context: Dict[str, Any], output_path: Path) -> Tuple[bool, str]:
        try:
            doc = DocxTemplate(str(template_path))
            
            if 'CANVAS' in context and isinstance(context['CANVAS'], str) and context['CANVAS']:
                if os.path.exists(context['CANVAS']):
                    try:
                        context['CANVAS'] = InlineImage(doc, context['CANVAS'], width=Mm(85))
                    except Exception as e:
                        DataManager.log_error(f"Ошибка создания InlineImage: {e}")
                        context['CANVAS'] = ""
            
            doc.render(context)
            doc.save(str(output_path))
            return True, str(output_path)
        except Exception as e:
            DataManager.log_error(f"Ошибка создания DOCX: {e}")
            return False, str(e)
    
    def docx_to_jpg_pipeline_multipage(self, docx_path: Path, output_dir: Path, base_filename: str) -> Tuple[List[Path], str]:
        temp_pdf = output_dir / f"temp_{datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
        try:
            success, _ = self.converter.docx_to_pdf(docx_path, temp_pdf)
            if not success:
                return [], "Ошибка конвертации в PDF"
            
            jpg_files = self.converter.pdf_to_jpg_multipage(temp_pdf, output_dir, base_filename)
            
            if temp_pdf.exists():
                os.remove(str(temp_pdf))
            
            if jpg_files:
                return jpg_files, ""
            else:
                return [], "Не удалось создать JPG"
        except Exception as e:
            DataManager.log_error(f"Ошибка в конвейере DOCX->JPG: {e}")
            if temp_pdf.exists():
                os.remove(str(temp_pdf))
            return [], str(e)

class GovDocApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(APP_SIZE)
        self.root.minsize(int(MIN_SIZE.split('x')[0]), int(MIN_SIZE.split('x')[1]))
        
        self.default_font = ('Times New Roman', 10)
        self.header_font = ('Times New Roman', 11, 'bold')
        self.small_font = ('Times New Roman', 8)
        
        self.data_manager = DataManager()
        self.data = self.data_manager.data
        
        self.word_available = ConversionChecker.check_word()
        self.libreoffice_path = ConversionChecker.check_libreoffice()
        self.converter = DocumentConverter(self.word_available, self.libreoffice_path)
        self.generator = DocumentGenerator(self.converter)
        self.poppler_available = ConversionChecker.check_poppler()
        
        self.canvas_photo = None
        self.load_canvas_image()
        
        self.modules = {}
        self.current_module_frames = {}
        
        self.setup_styles()
        self.create_widgets()
        self.set_current_date()
        
        # Загружаем модули (они создают свои поля)
        self.load_modules()
        
        # Привязываем контекстное меню ПОСЛЕ создания всех полей
        self.setup_paste_bindings()
    
    def setup_styles(self):
        style = ttk.Style()
        style.configure('Header.TLabel', font=self.header_font)
        style.configure('Generate.TButton', font=('Times New Roman', 11, 'bold'))
        style.configure('Prepare.TButton', font=('Times New Roman', 11))
        style.configure('Tool.TButton', font=('Times New Roman', 10))
        style.configure('Disclaimer.TLabel', font=('Times New Roman', 8, 'italic'), foreground='gray')
        self.root.option_add('*Font', self.default_font)
    
    def load_canvas_image(self):
        if CANVAS_IMAGE_PATH.exists():
            try:
                img = Image.open(CANVAS_IMAGE_PATH)
                img.thumbnail((150, 80), Image.Resampling.LANCZOS)
                self.canvas_photo = ImageTk.PhotoImage(img)
            except Exception as e:
                DataManager.log_error(f"Ошибка загрузки CANVAS.jpg: {e}")
    
    def get_canvas_path(self) -> str:
        return str(CANVAS_IMAGE_PATH) if CANVAS_IMAGE_PATH.exists() else ""
    
    def setup_paste_bindings(self):
        """Настройка вставки из буфера обмена через контекстное меню"""
        
        def paste_from_menu(event):
            """Вставка из контекстного меню"""
            try:
                widget = event.widget
                clipboard_text = self.root.clipboard_get()
                widget.insert(tk.INSERT, clipboard_text)
            except:
                pass
            return "break"
        
        def show_context_menu(event):
            """Показать контекстное меню с пунктом Вставить"""
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Вставить", command=lambda: paste_from_menu(event))
            menu.post(event.x_root, event.y_root)
        
        def bind_context_menu(widget):
            for child in widget.winfo_children():
                if isinstance(child, ttk.Entry) or isinstance(child, tk.Entry):
                    child.bind('<Button-3>', show_context_menu)
                else:
                    bind_context_menu(child)
        
        bind_context_menu(self.root)
    
    def load_modules(self):
        for py_file in MODULES_DIR.glob("*.py"):
            if py_file.name == "__init__.py":
                continue
            
            module_name = py_file.stem
            
            try:
                spec = importlib.util.spec_from_file_location(module_name, py_file)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                
                if hasattr(module, 'MODULE_INFO') and hasattr(module, 'create_ui'):
                    self.modules[module_name] = module
                    self.add_module_tab(module_name, module.MODULE_INFO, module.create_ui)
            except Exception as e:
                DataManager.log_error(f"Ошибка загрузки модуля {module_name}: {e}")
        
        if not self.modules:
            empty_frame = ttk.Frame(self.notebook)
            self.notebook.add(empty_frame, text="Нет модулей")
            ttk.Label(empty_frame, text="Нет загруженных модулей.\n\nПоместите файлы модулей в папку MODULES.").pack(expand=True)
    
    def add_module_tab(self, module_name, module_info, create_ui_func):
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text=module_info.get("tab_name", module_name))
        
        def get_common_data():
            return {
                'dd': self.day_entry.get().strip().zfill(2),
                'mm': self.get_month_name(self.month_entry.get().strip()),
                'yyyy': self.year_entry.get().strip(),
                'num': self.order_entry.get().strip(),
                'signer': self.signer_entry.get().strip(),
                'CANVAS': self.get_canvas_path()
            }
        
        def get_template_path():
            # Если у модуля есть свой get_template_path - используем его
            if hasattr(module_frame, 'get_template_path'):
                return module_frame.get_template_path()
            # Иначе используем стандартный из MODULE_INFO
            template_name = module_info.get("template")
            if template_name:
                return TEMPLATES_DIR / template_name
            return None
        
        def get_module_data(module_name, key, default=None):
            return self.get_module_data(module_name, key, default)
        
        def save_module_data(module_name, key, value):
            self.save_module_data(module_name, key, value)
        
        module_frame = create_ui_func(frame, self, module_name)
        
        self.current_module_frames[module_name] = {
            'frame': module_frame,
            'info': module_info,
            'get_common_data': get_common_data,
            'get_template_path': get_template_path,
            'get_context': module_frame.get_context if hasattr(module_frame, 'get_context') else lambda d: {},
            'save_data': module_frame.save_data if hasattr(module_frame, 'save_data') else lambda: {},
            'load_data': module_frame.load_data if hasattr(module_frame, 'load_data') else lambda d: None,
            'validate_inputs': module_frame.validate_inputs if hasattr(module_frame, 'validate_inputs') else lambda: (True, "")
        }
        
        if module_name in self.data.modules_data and hasattr(module_frame, 'load_data'):
            module_frame.load_data(self.data.modules_data[module_name])
    
    def get_module_data(self, module_name, key, default=None):
        if module_name in self.data.modules_data:
            return self.data.modules_data[module_name].get(key, default)
        return default
    
    def save_module_data(self, module_name, key, value):
        if module_name not in self.data.modules_data:
            self.data.modules_data[module_name] = {}
        self.data.modules_data[module_name][key] = value
    
    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, pady=2)
        
        info_text = " | ".join([
            "LibreOffice" if self.libreoffice_path else ("MS Word" if self.word_available else "Нет ПО"),
            "Poppler: да" if self.poppler_available else "Poppler: нет"
        ])
        ttk.Label(info_frame, text=info_text, font=self.small_font).pack(side=tk.LEFT)
        
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        common_frame = ttk.LabelFrame(main_frame, text="Общие данные", padding="5")
        common_frame.pack(fill=tk.X, pady=5)
        
        row1 = ttk.Frame(common_frame)
        row1.pack(fill=tk.X, pady=2)
        
        ttk.Label(row1, text="Дата:", width=6, anchor="w").pack(side=tk.LEFT)
        self.day_entry = ttk.Entry(row1, width=4)
        self.day_entry.pack(side=tk.LEFT)
        ttk.Label(row1, text=".").pack(side=tk.LEFT)
        self.month_entry = ttk.Entry(row1, width=4)
        self.month_entry.pack(side=tk.LEFT)
        ttk.Label(row1, text=".").pack(side=tk.LEFT)
        self.year_entry = ttk.Entry(row1, width=6)
        self.year_entry.pack(side=tk.LEFT)
        ttk.Button(row1, text="Сегодня", command=self.set_current_date, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="Номер:", width=7, anchor="w").pack(side=tk.LEFT, padx=(10,0))
        self.order_entry = ttk.Entry(row1, width=10)
        self.order_entry.pack(side=tk.LEFT)
        ttk.Label(row1, text=f"({self.data.last_order_number:04d})", font=self.small_font, foreground='gray').pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1, text="Подписывающий:", width=17, anchor="w").pack(side=tk.LEFT, padx=(10,0))
        self.signer_entry = ttk.Entry(row1, width=20)
        self.signer_entry.pack(side=tk.LEFT)
        if self.data.last_signer:
            self.signer_entry.insert(0, self.data.last_signer)
        
        row3 = ttk.Frame(common_frame)
        row3.pack(fill=tk.X, pady=2)
        
        if self.canvas_photo:
            canvas_label = ttk.Label(row3, image=self.canvas_photo)
            canvas_label.pack(side=tk.LEFT)
        else:
            ttk.Label(row3, text="[CANVAS.jpg]", font=self.small_font, foreground='gray').pack(side=tk.LEFT)
        
        ttk.Label(row3, text="(замените CANVAS.JPG)", font=self.small_font, foreground='gray').pack(side=tk.LEFT, padx=5)
        
        btn_frame = ttk.Frame(row3)
        btn_frame.pack(side=tk.RIGHT)
        
        self.prepare_btn = ttk.Button(btn_frame, text="ПОДГОТОВИТЬ", 
                                       command=lambda: self.generate_document(prepare_mode=True), width=16)
        self.prepare_btn.pack(side=tk.LEFT, padx=2)
        
        self.generate_btn = ttk.Button(btn_frame, text="СФОРМИРОВАТЬ", style='Generate.TButton',
                                        command=lambda: self.generate_document(prepare_mode=False), width=25)
        self.generate_btn.pack(side=tk.LEFT, padx=2)
        
        tools_frame = ttk.LabelFrame(main_frame, text="Инструменты", padding="5")
        tools_frame.pack(fill=tk.X, pady=5)
        
        btn_container = ttk.Frame(tools_frame)
        btn_container.pack()
        
        ttk.Button(btn_container, text="Проверить Poppler", command=self.check_poppler_manual, width=18).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_container, text="DOCX → JPG", command=self.convert_docx_to_jpg, width=16).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_container, text="Открыть IMG", command=lambda: self.open_folder('images'), width=16).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_container, text="Открыть DOCX", command=lambda: self.open_folder('complete'), width=16).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_container, text="Справка", command=self.show_help, width=16).pack(side=tk.LEFT, padx=2)
        
        disclaimer_frame = ttk.Frame(main_frame)
        disclaimer_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))
        ttk.Label(disclaimer_frame, text=DISCLAIMER, style='Disclaimer.TLabel', wraplength=900).pack()
    
    def set_current_date(self):
        now = datetime.now()
        self.day_entry.delete(0, tk.END)
        self.day_entry.insert(0, f"{now.day:02d}")
        self.month_entry.delete(0, tk.END)
        self.month_entry.insert(0, f"{now.month:02d}")
        self.year_entry.delete(0, tk.END)
        self.year_entry.insert(0, str(now.year))
    
    def get_month_name(self, month_num):
        return MONTHS.get(int(month_num), "")
    
    def get_current_tab_module(self):
        current_tab = self.notebook.select()
        if not current_tab:
            return None, None
        
        tab_text = self.notebook.tab(current_tab, "text")
        
        for module_name, module_data in self.current_module_frames.items():
            if module_data['info'].get('tab_name') == tab_text:
                return module_name, module_data
        return None, None
    
    def generate_document(self, prepare_mode: bool = False):
        module_name, module_data = self.get_current_tab_module()
        
        if not module_data:
            messagebox.showerror("Ошибка", "Не выбран тип документа")
            return
        
        if not self.order_entry.get().strip():
            messagebox.showerror("Ошибка", "Введите номер документа")
            return
        
        if not self.signer_entry.get().strip():
            messagebox.showerror("Ошибка", "Введите подписывающего")
            return
        
        valid, error_msg = module_data['validate_inputs']()
        if not valid:
            messagebox.showerror("Ошибка ввода", error_msg)
            return
        
        template_path = module_data['get_template_path']()
        if not template_path or not template_path.exists():
            messagebox.showerror("Ошибка", f"Шаблон не найден: {template_path}")
            return
        
        if not prepare_mode and not self.word_available and not self.libreoffice_path:
            if not messagebox.askyesno("Внимание", "Не обнаружено ПО для конвертации.\nJPG не будет создан.\nПродолжить?"):
                return
        
        try:
            common_data = module_data['get_common_data']()
            module_context = module_data['get_context'](common_data)
            
            context = {
                'dd': common_data['dd'],
                'mm': common_data['mm'],
                'yyyy': common_data['yyyy'],
                'num': common_data['num'],
                'signer': common_data['signer'],
                'CANVAS': common_data.get('CANVAS', ''),
                **module_context
            }
            
            safe_order = "".join(c for c in common_data['num'] if c.isalnum() or c in ('-', '_')).strip()
            doc_name = f"{safe_order} {module_data['info']['name']}"
            docx_path = COMPLETE_DIR / f"{doc_name}.docx"
            
            self.root.config(cursor="watch")
            self.root.update()
            
            success, msg = self.generator.create_docx(template_path, context, docx_path)
            if not success:
                raise Exception(f"Ошибка создания DOCX: {msg}")
            
            results = [f"DOCX: {docx_path.name}"]
            jpg_files = []
            
            if not prepare_mode and self.poppler_available:
                jpg_files, error = self.generator.docx_to_jpg_pipeline_multipage(docx_path, IMAGES_DIR, doc_name)
                if jpg_files:
                    if len(jpg_files) == 1:
                        results.append(f"JPG: {jpg_files[0].name}")
                    else:
                        results.append(f"JPG ({len(jpg_files)} стр.):")
                        for f in jpg_files:
                            results.append(f"  - {f.name}")
                else:
                    results.append(f"JPG не создан: {error}")
            elif not prepare_mode and not self.poppler_available:
                results.append("JPG не создан: нет Poppler")
            else:
                results.append("JPG не создан (режим подготовки)")
            
            self.data.last_order_number = int(common_data['num'])
            self.data.last_signer = common_data['signer']
            
            module_save_data = module_data['save_data']()
            if module_save_data:
                self.save_module_data(module_name, "saved", module_save_data)
            
            self.data_manager.save(self.data)
            
            self.root.config(cursor="")
            
            result_text = "Готово.\n\n" + "\n".join(results)
            messagebox.showinfo("Результат", result_text)
            
            if jpg_files:
                if messagebox.askyesno("Открыть папку", "Открыть папку с JPG?"):
                    self.open_folder('images')
            elif prepare_mode:
                if messagebox.askyesno("Открыть папку", "Открыть папку с DOCX?"):
                    self.open_folder('complete')
            
        except Exception as e:
            self.root.config(cursor="")
            DataManager.log_error(f"Ошибка при генерации: {e}")
            messagebox.showerror("Ошибка", str(e))
    
    def check_poppler_manual(self):
        if ConversionChecker.check_poppler():
            messagebox.showinfo("Poppler", "Poppler найден")
        else:
            msg = ("Poppler не найден!\n\n"
                   "Windows: распакуйте poppler в папку программы\n"
                   "Linux: sudo apt-get install poppler-utils\n"
                   "Mac: brew install poppler")
            messagebox.showerror("Poppler", msg)
    
    def convert_docx_to_jpg(self):
        file_path = filedialog.askopenfilename(
            title="Выберите DOCX файл",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )
        if not file_path:
            return
        path = Path(file_path)
        self.root.config(cursor="watch")
        self.root.update()
        base_name = path.stem
        jpg_files, error = self.generator.docx_to_jpg_pipeline_multipage(path, IMAGES_DIR, base_name)
        self.root.config(cursor="")
        if jpg_files:
            msg = f"Создано JPG: {len(jpg_files)} файлов"
            if messagebox.askyesno("Успех", msg + "\n\nОткрыть папку?"):
                self.open_folder('images')
        else:
            messagebox.showerror("Ошибка", f"Не удалось конвертировать:\n{error}")
    
    def open_folder(self, folder_type):
        folders = {'images': IMAGES_DIR, 'complete': COMPLETE_DIR}
        path = folders.get(folder_type)
        if path and path.exists():
            if platform.system() == 'Windows':
                os.startfile(str(path))
            elif platform.system() == 'Darwin':
                subprocess.run(['open', str(path)])
            else:
                subprocess.run(['xdg-open', str(path)])
    
    def show_help(self):
        method = "Word" if self.word_available else "LibreOffice" if self.libreoffice_path else "НЕТ"
        poppler = "есть" if self.poppler_available else "нет"
        help_text = f"""Система подготовки документов
Версия {VERSION}

Конвертация: {method}
Poppler: {poppler}

Документы сохраняются в папку COMPLETE
Изображения — в папку IMAGES

Для добавления модулей поместите .py файлы в папку MODULES
"""
        messagebox.showinfo("О программе", help_text)

def main():
    root = tk.Tk()
    app = GovDocApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()