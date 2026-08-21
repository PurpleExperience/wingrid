import tkinter as tk
from tkinter import ttk, messagebox, filedialog, BooleanVar
import configparser
import os
import sys
import math
import logging
import webbrowser
import win32gui
import win32con
import win32api
from win32api import GetSystemMetrics, EnumDisplayMonitors, GetMonitorInfo
from ctypes import windll

# ============================================================
# Constants (previously magic numbers scattered through code)
# ============================================================
APP_NAME = "WinGrid"
DEFAULT_WINDOW_GEOMETRY = "650x380"
DEFAULT_MIN_SIZE = (650, 380)
DEFAULT_SPACING_PX = 8
TIGHT_OVERLAP_BASE_PX = 15
TIGHT_SPACING_Y = -8
AUTO_APPLY_DELAY_MS = 10
AUTO_CLOSE_TOAST_MS = 3000
ARRANGE_REAPPLY_DELAY_MS = 10
SCROLLABLE_FRAME_WIDTH = 660
CONFIG_FILE_NAME = "settings.ini"
ABOUT_CONTACT_HANDLE = "t.me/alex_dev404"
ABOUT_CONTACT_URL = "https://t.me/alex_dev404"

# Readability: larger, consistent fonts across the whole UI instead of the
# tiny tkinter defaults.
FONT_FAMILY = "Segoe UI"
FONT_NORMAL = (FONT_FAMILY, 11)
FONT_BOLD = (FONT_FAMILY, 11, "bold")
FONT_BUTTON = (FONT_FAMILY, 11)
FONT_MONITOR_LABEL = (FONT_FAMILY, 11, "bold")
FONT_TOOLTIP = (FONT_FAMILY, 10)
FONT_SMALL_BUTTON = (FONT_FAMILY, 11, "bold")  # for compact "..."/"X"/"+" buttons

# ============================================================
# Logging (was: bare print(), invisible once frozen as .exe)
# ============================================================
def _init_logging(application_path):
    log_path = os.path.join(application_path, f"{APP_NAME}.log")
    logger = logging.getLogger(APP_NAME)
    logger.setLevel(logging.INFO)
    # logging.basicConfig(..., encoding=...) only accepts the encoding
    # keyword on Python 3.9+; building a FileHandler by hand works on any
    # Python 3 version.
    handler = logging.FileHandler(log_path, encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(handler)
    return logger


TRANSLATIONS = {
    'ru': {
        'window_title': APP_NAME,
        'config_file': "Файл конфигурации:",
        'load': "Загрузить",
        'auto_apply': "Применять при запуске",
        'window_titles': "Заголовки окон:",
        'monitor': "Монитор:",
        'width': "Ширина:",
        'height': "Высота:",
        'save_settings': "Сохранить настройки",
        'apply': "Применить",
        'exit': "Выход",
        'primary': "Основной",
        'windows_label': "Окна",
        'custom_geometry_title': "Свой размер и позиция",
        'use_custom_geometry': "Задать свой размер/позицию для этого окна",
        'custom_width': "Ширина:",
        'custom_height': "Высота:",
        'custom_x': "Позиция X (от левого края монитора):",
        'custom_y': "Позиция Y (от верхнего края монитора):",
        'clear_override': "Сбросить (автоматически)",
        'close': "Закрыть",
        'live_preview_unavailable': "Окно не найдено — предпросмотр недоступен",
        'warning': "Предупреждение",
        'error': "Ошибка",
        'success': "Успех",
        'attention': "Внимание",
        'select_config': "Выберите файл конфигурации",
        'save_as_title': "Сохранить конфигурацию как",
        'overwrite_prompt': "Файл \"{}\" уже существует.\nПерезаписать его или сохранить под новым именем?",
        'overwrite': "Перезаписать",
        'save_as': "Сохранить как...",
        'config_loaded': "Конфигурация загружена успешно",
        'config_not_found': "Файл конфигурации не найден",
        'add_title': "Добавьте хотя бы один заголовок",
        'windows_not_found': "Не найдены следующие окна:",
        'settings_saved': "Настройки сохранены в",
        'auto_close': "Автозакрытие через {} сек",
        'monitor_info': "Монитор {} ({}x{}):",
        'notepad_title': '{}.txt – Блокнот',
        'tight_windows': "Окна вплотную",
        'consider_taskbar': "С учётом таскбара",
        'not_found_hint': " (не найдено)",
        'cancel': "Отмена",
        'unsaved_marker': "*",
        'tooltip_auto_apply': "Автоматически расставит окна при запуске программы",
        'tooltip_tight_windows': "Окна слегка перекрываются, чтобы устранить зазоры между ними",
        'tooltip_consider_taskbar': "Не занимать окнами область панели задач",
        'about_title': "О программе",
        'about_text': (
            "Программа автоматически расставляет окна по сетке "
            "на выбранных мониторах согласно заданным заголовкам."
        ),
        'about_contact_label': "Контакт:",
        'about_button': "(?)",
    },
    'en': {
        'window_title': APP_NAME,
        'config_file': "Configuration file:",
        'load': "Load",
        'auto_apply': "Apply on startup",
        'window_titles': "Window titles:",
        'monitor': "Monitor:",
        'width': "Width:",
        'height': "Height:",
        'save_settings': "Save settings",
        'apply': "Apply",
        'exit': "Exit",
        'primary': "Primary",
        'windows_label': "Windows",
        'custom_geometry_title': "Custom size and position",
        'use_custom_geometry': "Use a custom size/position for this window",
        'custom_width': "Width:",
        'custom_height': "Height:",
        'custom_x': "X position (from monitor's left edge):",
        'custom_y': "Y position (from monitor's top edge):",
        'clear_override': "Clear (use automatic)",
        'close': "Close",
        'live_preview_unavailable': "Window not found — live preview unavailable",
        'warning': "Warning",
        'error': "Error",
        'success': "Success",
        'attention': "Attention",
        'select_config': "Select configuration file",
        'save_as_title': "Save configuration as",
        'overwrite_prompt': "The file \"{}\" already exists.\nOverwrite it or save under a new name?",
        'overwrite': "Overwrite",
        'save_as': "Save as...",
        'config_loaded': "Configuration loaded successfully",
        'config_not_found': "Configuration file not found",
        'add_title': "Add at least one title",
        'windows_not_found': "Following windows not found:",
        'settings_saved': "Settings saved to",
        'auto_close': "Auto-close in {} sec",
        'monitor_info': "Monitor {} ({}x{}):",
        'notepad_title': '{}.txt - Notepad',
        'tight_windows': "Tight windows",
        'consider_taskbar': "Consider taskbar",
        'not_found_hint': " (not found)",
        'cancel': "Cancel",
        'unsaved_marker': "*",
        'tooltip_auto_apply': "Automatically arrange windows when the program starts",
        'tooltip_tight_windows': "Windows slightly overlap to eliminate gaps between them",
        'tooltip_consider_taskbar': "Don't place windows over the taskbar area",
        'about_title': "About",
        'about_text': (
            "Automatically arranges windows in a grid across the selected "
            "monitors, based on the window titles you provide."
        ),
        'about_contact_label': "Contact:",
        'about_button': "(?)",
    }
}


class MonitorInfo:
    def __init__(self, logger=None):
        self.monitors = []
        self.logger = logger or logging.getLogger(APP_NAME)
        try:
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception as e:
            self.logger.warning(f"Error setting DPI awareness: {e}")
        self.get_monitors_info()

    def get_monitors_info(self):
        try:
            monitors = EnumDisplayMonitors()

            for i, monitor in enumerate(monitors):
                monitor_info = GetMonitorInfo(monitor[0])
                monitor_area = monitor_info.get('Monitor')
                work_area = monitor_info.get('Work')

                try:
                    import ctypes
                    dpi_x = ctypes.c_uint()
                    dpi_y = ctypes.c_uint()
                    # Bug fix: monitor[0] is a pywin32 PyHANDLE wrapper, not
                    # a plain int — ctypes can't convert it directly and
                    # raised "Don't know how to convert parameter 1" here,
                    # silently falling back to scale=100 every time.
                    ctypes.windll.shcore.GetDpiForMonitor(
                        int(monitor[0]),
                        0,
                        ctypes.byref(dpi_x),
                        ctypes.byref(dpi_y)
                    )
                    scale_percentage = int(dpi_x.value * 100 / 96)
                except Exception as e:
                    self.logger.warning(f"Error getting scale for monitor {i + 1}: {e}")
                    scale_percentage = 100

                physical_width = monitor_area[2] - monitor_area[0]
                physical_height = monitor_area[3] - monitor_area[1]

                monitor_data = {
                    'handle': monitor[0],
                    'left': monitor_area[0],
                    'top': monitor_area[1],
                    'width': physical_width,
                    'height': physical_height,
                    'physical_width': physical_width,
                    'physical_height': physical_height,
                    'work_left': work_area[0],
                    'work_top': work_area[1],
                    'work_width': work_area[2] - work_area[0],
                    'work_height': work_area[3] - work_area[1],
                    'device': monitor_info.get('Device'),
                    'is_primary': monitor_info.get('Flags') == 1,
                    'scale': scale_percentage
                }
                self.monitors.append(monitor_data)

        except Exception as e:
            self.logger.error(f"Error detecting monitors: {e}")

        return self.monitors


class MonitorFrame(ttk.Frame):
    """
    Represents one physical monitor row in the monitor panel.

    GUI improvement: now shows a small live layout preview canvas is handled
    by MonitorLayoutPreview (separate widget) — this frame stays focused on
    the textual info row, but exposes width/height as real StringVars so
    load_monitor_settings() no longer crashes on missing attributes.
    """

    def __init__(self, parent, monitor_info, index, colors, current_language='ru', **kwargs):
        super().__init__(parent, **kwargs)
        self.colors = colors
        self.monitor_info = monitor_info
        self.current_language = current_language
        self.index = index

        # Bug fix: these StringVars previously didn't exist, but
        # load_monitor_settings() referenced them -> AttributeError on load.
        self.width_var = tk.StringVar(value=str(monitor_info['physical_width']))
        self.height_var = tk.StringVar(value=str(monitor_info['physical_height']))

        monitor_label = TRANSLATIONS[self.current_language]['monitor_info'].format(
            index + 1,
            monitor_info['physical_width'],
            monitor_info['physical_height']
        )
        if monitor_info['is_primary']:
            monitor_label += f" [{TRANSLATIONS[self.current_language]['primary']}]"

        self.monitor_label = tk.Label(self,
                                       text=monitor_label,
                                       bg=self.colors['bg'],
                                       fg=self.colors['text'],
                                       font=FONT_MONITOR_LABEL)
        self.monitor_label.pack(side=tk.LEFT, padx=(0, 10), pady=4)

        self.left = monitor_info['left']
        self.top = monitor_info['top']
        self.work_left = monitor_info['work_left']
        self.work_top = monitor_info['work_top']
        self.work_width = monitor_info['work_width']
        self.work_height = monitor_info['work_height']
        self.is_primary = monitor_info['is_primary']
        self.width = monitor_info['width']
        self.height = monitor_info['height']

    def get_config(self):
        return {
            'width': self.width,
            'height': self.height,
            'left': self.left,
            'top': self.top,
            'work_left': self.work_left,
            'work_top': self.work_top,
            'work_width': self.work_width,
            'work_height': self.work_height,
            'is_primary': self.is_primary
        }


class MonitorLayoutPreview(tk.Canvas):
    """
    GUI improvement #4: a small proportional preview of the physical monitor
    layout with numbered rectangles, so picking "Monitor: 2" in a dropdown
    actually corresponds to something the user can see, instead of forcing
    them to remember which number is which screen.

    Each rectangle carries its own resolution, a "Primary" badge, and a
    live count of how many configured window rows are currently assigned
    to it — replacing the separate "Monitor N (WxH): [Primary]" text row
    that used to sit underneath. Call refresh() whenever the window list
    or monitor assignments change to keep the counts current.
    """

    PREVIEW_HEIGHT = 110  # used only as the floor/starting height
    MAX_PREVIEW_HEIGHT = 320  # cap so a large monitor grid can't blow up the UI
    MIN_RECT_H_FOR_TEXT = 34  # below this, resolution text gets dropped (see below)
    PADDING = 12

    def __init__(self, parent, monitors, colors, translations, get_window_counts=None, **kwargs):
        super().__init__(parent, height=self.PREVIEW_HEIGHT,
                          bg=colors['bg'], highlightthickness=0, **kwargs)
        self.monitors = monitors
        self.colors = colors
        self.translations = translations  # {'primary': ..., 'windows_label': ...}
        self.get_window_counts = get_window_counts or (lambda: [0] * len(monitors))
        # Debounced: <Configure> fires on every single pixel while the
        # user drags-resizes the main window, and a full redraw (delete
        # all canvas items + recreate rectangles/text for every monitor)
        # on each one of those events was visibly janky during live
        # resize. Coalesce rapid-fire resize events into one redraw
        # shortly after they stop.
        self._redraw_after_id = None
        self.bind("<Configure>", self._schedule_redraw)
        self._redraw()

    def _schedule_redraw(self, _event=None):
        if self._redraw_after_id is not None:
            try:
                self.after_cancel(self._redraw_after_id)
            except (ValueError, tk.TclError):
                pass
        self._redraw_after_id = self.after(30, self._debounced_redraw)

    def _debounced_redraw(self):
        self._redraw_after_id = None
        self._redraw()

    def refresh(self):
        """Call after the window list or a monitor assignment changes."""
        self._redraw()

    def _redraw(self):
        self.delete("all")
        if not self.monitors:
            return

        canvas_width = max(self.winfo_width(), 200)

        min_left = min(m['left'] for m in self.monitors)
        max_right = max(m['left'] + m['width'] for m in self.monitors)
        min_top = min(m['top'] for m in self.monitors)
        max_bottom = max(m['top'] + m['height'] for m in self.monitors)

        total_w = max(max_right - min_left, 1)
        total_h = max(max_bottom - min_top, 1)

        available_w = canvas_width - 2 * self.PADDING

        # The canvas used to have a fixed height (PREVIEW_HEIGHT), so any
        # monitor arrangement with more than one row (e.g. a 2x3 grid of six
        # monitors) forced the scale down so far that each rectangle's
        # rendered height dropped below MIN_RECT_H_FOR_TEXT and the
        # resolution label was silently skipped - only the monitor number
        # showed.
        #
        # Fix attempt #1 sized the canvas to the raw bounding-box aspect
        # ratio (total_h / total_w), but that also grows the canvas for a
        # SINGLE tall-ish monitor (e.g. one 2560x1440 screen), which doesn't
        # need extra height at all. What actually needs more room is having
        # multiple ROWS of monitors, not any one monitor's own aspect ratio.
        # So estimate the row count and size the canvas for that instead.
        avg_monitor_h = sum(m['height'] for m in self.monitors) / len(self.monitors)
        estimated_rows = max(1, round(total_h / avg_monitor_h)) if avg_monitor_h > 0 else 1

        per_row_height = 80  # enough for the 3-line label at a reasonable scale
        target_height = max(self.PREVIEW_HEIGHT, estimated_rows * per_row_height + 2 * self.PADDING)
        target_height = min(target_height, self.MAX_PREVIEW_HEIGHT)

        scale = min(available_w / total_w, (target_height - 2 * self.PADDING) / total_h)

        preview_height = target_height
        if int(self.cget('height')) != int(preview_height):
            self.configure(height=int(preview_height))

        offset_x = (canvas_width - total_w * scale) / 2
        offset_y = (preview_height - total_h * scale) / 2

        counts = self.get_window_counts()

        for i, m in enumerate(self.monitors):
            x0 = offset_x + (m['left'] - min_left) * scale
            y0 = offset_y + (m['top'] - min_top) * scale
            x1 = x0 + m['width'] * scale
            y1 = y0 + m['height'] * scale

            fill = self.colors['button_bg']
            outline = self.colors['text']
            self.create_rectangle(x0, y0, x1, y1, fill=fill, outline=outline, width=1)

            cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
            rect_h = y1 - y0

            heading = f"{i + 1}"
            if m.get('is_primary'):
                heading += f" · {self.translations['primary']}"

            count = counts[i] if i < len(counts) else 0
            count_text = f"{self.translations['windows_label']}: {count}"

            # Three stacked lines: number (+ Primary badge), resolution,
            # window count — laid out so they stay centered even if the
            # rectangle is short and the text ends up slightly overflowing.
            if rect_h >= 60:
                self.create_text(cx, cy - 16, text=heading, fill=self.colors['text'],
                                  font=(FONT_FAMILY, 12, "bold"))
                self.create_text(cx, cy, text=f"{m['physical_width']}x{m['physical_height']}",
                                  fill=self.colors['text'], font=(FONT_FAMILY, 10))
                self.create_text(cx, cy + 16, text=count_text, fill=self.colors['text'],
                                  font=(FONT_FAMILY, 10))
            elif rect_h >= self.MIN_RECT_H_FOR_TEXT:
                # Not enough vertical room for three lines — drop the count.
                self.create_text(cx, cy - 8, text=heading, fill=self.colors['text'],
                                  font=(FONT_FAMILY, 12, "bold"))
                self.create_text(cx, cy + 8, text=f"{m['physical_width']}x{m['physical_height']}",
                                  fill=self.colors['text'], font=(FONT_FAMILY, 9))
            else:
                # Very short rectangle — just the number.
                self.create_text(cx, cy, text=str(i + 1), fill=self.colors['text'],
                                  font=(FONT_FAMILY, 12, "bold"))


class AutoCloseMessageBox:
    """Non-blocking toast that auto-dismisses. Used for 'windows not found'."""

    def __init__(self, title, message, duration=AUTO_CLOSE_TOAST_MS):
        self.root = tk.Toplevel()
        self.root.title(title)
        self.root.resizable(False, False)
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)

        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        label = tk.Label(self.root, text=message, padx=24, pady=18, font=FONT_NORMAL, justify='left')
        label.pack()

        self.root.update_idletasks()
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()

        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        self.root.geometry(f"+{x}+{y}")
        self.root.after(duration, self.close)

    def close(self):
        self.root.destroy()


class WindowManager:
    def __init__(self, root, logger, override_config_path=None, headless=False, force_apply=False):
        self.logger = logger
        self.headless = headless

        self.root = root

        self.tight_windows_var = BooleanVar(value=False)
        self.consider_taskbar_var = BooleanVar(value=True)

        self.dirty = False  # informational only, not used to block closing

        self.colors = {
            'bg': '#2B2B2B',
            'input_bg': '#3C3F41',
            'text': '#E8E8E8',
            'button_bg': '#4C5052',
            'button_active': '#595B5D',
            'scroll_bg': '#383838',
            'error_bg': '#5C2B2B',    # title matches no window at all
            'found_bg': '#2B5C3A',    # will actually get a window when applied
            'warning_bg': '#5C4E1F',  # title exists, but not enough duplicates for this row
        }

        self.auto_apply_var = BooleanVar(value=False)
        self.config = configparser.ConfigParser()
        self.current_language = 'ru'
        self.window_inputs = []
        self.monitor_frames = []
        self._validity_after_id = None
        self._resize_repaint_after_id = None

        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        # This fixed path next to the program always exists/is checked
        # first. It normally holds the settings themselves, but it can
        # also just point ("LastConfigPath") at a different config file
        # the user loaded/saved elsewhere — so the app reopens with
        # whichever config was in use last, not always this default file.
        self.bootstrap_config_path = os.path.join(application_path, CONFIG_FILE_NAME)
        # --config on the command line overrides the normal "reopen the
        # last used config" logic for this run only (doesn't update the
        # LastConfigPath pointer, so interactive runs are unaffected).
        self.config_file = override_config_path or self._resolve_last_config_path()
        self.load_config()

        self.root.minsize(*DEFAULT_MIN_SIZE)
        self._apply_saved_geometry()

        self.root.configure(bg=self.colors['bg'])
        self.setup_styles()

        self.main_frame = ttk.Frame(root, style='Dark.TFrame')
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self._make_draggable(self.main_frame)

        self.monitor_info = MonitorInfo(logger=self.logger)

        # Registry of widgets that need text updates on language switch.
        # GUI/code improvement #7: replaces fragile text-matching recursion.
        self.translatable_labels = {}   # widget -> translation key
        self.translatable_buttons = {}  # widget -> translation key
        self.translatable_checks = {}   # widget -> translation key

        self.create_config_file_section()
        self.create_monitor_selection()
        self.create_window_list_section()
        self.create_control_buttons()
        self._add_resize_grip()

        self.load_window_titles()
        self.window_titles_menu = None
        self._picker_outside_bind_id = None

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        # The picker popup is an overrideredirect Toplevel, so it doesn't
        # automatically minimize along with the main window on Windows —
        # close it explicitly whenever the main window gets minimized.
        self.root.bind('<Unmap>', self._on_root_unmap)
        # Same idea for moving/resizing: the popup doesn't follow the
        # window, so close it rather than leave it stranded.
        self.root.bind('<Configure>', self._on_root_configure)

        if headless:
            # CLI mode: apply the layout once the window/monitor info is
            # settled, then close automatically — no GUI is ever shown.
            self.root.after(AUTO_APPLY_DELAY_MS, self._run_headless_apply)
        elif force_apply or self.auto_apply_var.get():
            self.root.after(AUTO_APPLY_DELAY_MS, self.start_auto_sequence)

    # ------------------------------------------------------------------
    # Language handling (rewritten to use a widget registry instead of
    # recursive text-matching, which broke easily and was slow)
    # ------------------------------------------------------------------
    def toggle_language(self):
        self.current_language = 'en' if self.current_language == 'ru' else 'ru'
        self.lang_button.config(text="ENG" if self.current_language == 'ru' else "RUS")

        # Bug fix: this used to write self.config to disk immediately, but
        # self.config is only synced with the live UI state (window titles,
        # monitor assignments, etc.) inside _persist_settings() - so an
        # immediate write here could silently overwrite the saved config
        # file with a stale mix of old data + the new language, bypassing
        # the app's "no auto-save, only the Save settings button persists"
        # policy (see on_closing). Just update the in-memory config here;
        # it'll be written out next time the user actually saves.
        self.config.set('Language', 'Language', 'RUS' if self.current_language == 'ru' else 'ENG')
        self.mark_dirty()

        self.update_interface_language()

    def register_label(self, widget, key):
        self.translatable_labels[widget] = key
        widget.config(text=TRANSLATIONS[self.current_language][key])
        return widget

    def register_button(self, widget, key):
        self.translatable_buttons[widget] = key
        return widget

    def register_check(self, widget, key):
        self.translatable_checks[widget] = key
        return widget

    def update_interface_language(self):
        lang = self.current_language
        self._update_window_title()

        for widget, key in self.translatable_labels.items():
            widget.config(text=TRANSLATIONS[lang][key])
        for widget, key in self.translatable_buttons.items():
            widget.config(text=TRANSLATIONS[lang][key])
        for widget, key in self.translatable_checks.items():
            widget.config(text=TRANSLATIONS[lang][key])

        # Monitor layout preview: refresh its translated labels (Primary /
        # Windows count) and redraw.
        if hasattr(self, 'layout_preview'):
            self.layout_preview.translations = {
                'primary': TRANSLATIONS[lang]['primary'],
                'windows_label': TRANSLATIONS[lang]['windows_label'],
            }
            self.layout_preview.refresh()

        # Window title rows: update "Monitor:" labels, dropdown options,
        # and default Notepad titles
        new_monitor_names = self._build_monitor_names(lang)
        for entry in self.window_inputs:
            parent_frame = entry.master
            for child in parent_frame.winfo_children():
                if isinstance(child, tk.Label) and getattr(child, 'is_monitor_label', False):
                    child.config(text=TRANSLATIONS[lang]['monitor'])

            # Bug fix: the dropdown's option list (and the currently shown
            # value, if it carries a translated "[Primary]" suffix) used to
            # stay frozen in whichever language was active when the row was
            # created. Rebuild the option list and re-render the current
            # selection in the new language, without changing which
            # monitor is actually selected.
            if hasattr(entry, 'monitor_dropdown') and new_monitor_names:
                try:
                    selected_index = int(entry.monitor_var.get().split()[0]) - 1
                except (ValueError, IndexError):
                    selected_index = None
                entry.monitor_dropdown['values'] = new_monitor_names
                if selected_index is not None and 0 <= selected_index < len(new_monitor_names):
                    entry.monitor_var.set(new_monitor_names[selected_index])

            current_text = entry.get()
            # Only rewrite default notepad-style titles, leave custom ones alone
            for src_lang in ('ru', 'en'):
                for number in range(1, 100):
                    if current_text == TRANSLATIONS[src_lang]['notepad_title'].format(number):
                        entry.delete(0, tk.END)
                        entry.insert(0, TRANSLATIONS[lang]['notepad_title'].format(number))
                        break

        self._refresh_window_titles_validity()

    # ------------------------------------------------------------------
    def setup_styles(self):
        # Readability: set application-wide default fonts so every plain
        # tk widget (Label, Button, Entry, Menu, Listbox, message boxes)
        # picks up the larger font without having to touch each one.
        self.root.option_add("*Font", FONT_NORMAL)
        self.root.option_add("*Dialog.msg.font", FONT_NORMAL)

        self.style = ttk.Style()
        # NOTE: intentionally NOT switching to the 'clam' ttk theme here.
        # 'clam' does let us fully recolor the Combobox, but it also
        # replaces the native Windows checkbox indicator (blue checkmark)
        # with its own small square/"x" glyph — worse for readability and
        # not what's wanted. We stay on the native theme ('vista' on
        # Windows) so checkboxes look normal, and fix the Combobox text
        # visibility below without a theme switch.

        self.style.configure('Dark.TFrame', background=self.colors['bg'])
        self.style.configure('Dark.TButton',
                              background=self.colors['button_bg'],
                              foreground=self.colors['text'],
                              font=FONT_BUTTON,
                              padding=6)
        self.style.map('Dark.TButton',
                        background=[('active', self.colors['button_active'])])
        self.style.configure('Dark.TEntry',
                              fieldbackground=self.colors['input_bg'],
                              foreground=self.colors['text'],
                              font=FONT_NORMAL)
        self.style.configure('Dark.TScrollbar',
                              background=self.colors['button_bg'],
                              arrowcolor=self.colors['text'],
                              troughcolor=self.colors['scroll_bg'])
        self.style.configure('Dark.TCheckbutton',
                              background=self.colors['bg'],
                              foreground=self.colors['text'],
                              font=FONT_NORMAL)
        self.style.map('Dark.TCheckbutton',
                        background=[('active', self.colors['bg'])],
                        foreground=[('active', self.colors['text'])])
        self.style.configure('Dark.TSizegrip', background=self.colors['bg'])

        # Comboboxes (monitor dropdowns). Bug fix: on Windows' native theme
        # the field background can't actually be recolored (it stays the
        # system default, usually white/light gray) — the earlier version
        # set a *light* foreground for a dark theme, which on that
        # uncontrollable light background rendered as near-invisible
        # light-on-light text ("empty"-looking box). Using a dark,
        # readable foreground instead fixes that without touching the theme.
        self.style.configure('Dark.TCombobox',
                              foreground='#1A1A1A',
                              font=FONT_NORMAL,
                              padding=4)
        self.style.map('Dark.TCombobox',
                        foreground=[('readonly', '#1A1A1A'),
                                    ('disabled', '#1A1A1A')])
        self.root.option_add('*TCombobox*Listbox.font', FONT_NORMAL)

    # ------------------------------------------------------------------
    # Auto-apply on startup
    # Fix: this used to invoke Apply and then start a 3-second countdown
    # that closed the whole program automatically — surprising and
    # unwanted. Now it only arranges the windows and leaves the app open.
    # ------------------------------------------------------------------
    def start_auto_sequence(self):
        self.apply_btn.invoke()

    def _run_headless_apply(self):
        """CLI build: arrange the windows, then close automatically."""
        self.arrange_windows()
        # arrange_windows() re-applies once more after
        # ARRANGE_REAPPLY_DELAY_MS; wait past that before exiting so both
        # passes actually complete.
        #self.root.after(ARRANGE_REAPPLY_DELAY_MS + 100, self._headless_quit)

    def _headless_quit(self):
        self.root.destroy()

    def _close_title_picker_popup(self):
        """
        The single place the window-title picker popup ever gets closed
        from (Escape, an outside click, a successful selection, or the
        main window minimizing/moving). Centralizing this matters because
        show_window_titles_menu also binds a global <Button-1> handler on
        self.root to detect outside clicks - bug fix: that binding used to
        only get cleaned up by the outside-click handler itself, so
        closing via Escape or selecting a title left it dangling. Each
        reopen added another one (add='+'), so stale handlers accumulated
        for the life of the app instead of being replaced.
        """
        if self.window_titles_menu:
            if getattr(self, '_picker_outside_bind_id', None):
                try:
                    self.root.unbind('<Button-1>', self._picker_outside_bind_id)
                except tk.TclError:
                    pass
                self._picker_outside_bind_id = None
            self.window_titles_menu.destroy()
            self.window_titles_menu = None

    def _on_root_unmap(self, event):
        """
        Fires when the main window's visibility state changes, including
        minimize. The title-picker popup (an overrideredirect Toplevel)
        doesn't minimize on its own along with the main window, so close
        it explicitly here rather than leaving it floating over the desktop.
        """
        if self.root.state() == 'iconic':
            self._close_title_picker_popup()

    def _on_root_configure(self, event):
        """
        Fires whenever the main window moves or resizes — including our
        own custom drag-by-background (_make_draggable) and dragging the
        OS title bar. The picker popup is positioned once, relative to the
        entry that opened it, and never follows the window afterward, so
        it would otherwise get left behind at its old screen position.
        Simplest fix: close it rather than try to keep it glued in place.
        """
        if event.widget is self.root:
            self._close_title_picker_popup()
            self._schedule_resize_repaint()

    def _schedule_resize_repaint(self):
        """
        A Canvas with embedded native child widgets (create_window — used
        for the scrollable window-title list) can fail to repaint some of
        those children during a live drag-resize on Windows, leaving
        stale blank strips where a row's monitor label/dropdown used to
        be until something forces a fresh paint pass. Once resizing has
        been quiet for a moment, force one so nothing is left stuck.
        """
        if self._resize_repaint_after_id is not None:
            try:
                self.root.after_cancel(self._resize_repaint_after_id)
            except (ValueError, tk.TclError):
                pass
        self._resize_repaint_after_id = self.root.after(80, self._force_resize_repaint)

    def _force_resize_repaint(self):
        self._resize_repaint_after_id = None
        try:
            # update_idletasks alone doesn't always fix this on Windows —
            # it only flushes pending geometry/idle work. A full update()
            # additionally processes pending Expose/paint events, which is
            # what actually clears the stale strips.
            self.root.update_idletasks()
            self.root.update()
        except tk.TclError:
            pass

    def on_closing(self):
        # No auto-save on close: persisting is now explicit (the "Save
        # settings" button), so closing never silently overwrites a config
        # file with unsaved changes. The next launch still reopens whatever
        # config was last loaded/saved (see _resolve_last_config_path).
        self.root.quit()

    # ------------------------------------------------------------------
    def mark_dirty(self, *_args):
        self.dirty = True

    def mark_clean(self):
        self.dirty = False

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    def _on_config_path_changed(self):
        """
        Fires whenever config_path_var changes (loading, saving, or the
        initial value at startup). Keeps the filename shown in the entry
        and the config name shown in the window title in sync with the
        actual full path, without the UI ever showing that full path.
        """
        self.config_display_var.set(os.path.basename(self.config_path_var.get()) or CONFIG_FILE_NAME)
        self._update_window_title()

    def _update_window_title(self):
        base = TRANSLATIONS[self.current_language]['window_title']
        filename = os.path.basename(self.config_path_var.get()) or CONFIG_FILE_NAME
        self.root.title(f"[{filename}] {base}")

    def _block_entry_edit(self, event):
        """
        Keeps the config-file entry read-only (it shows just the filename,
        which isn't a loadable path on its own) while still allowing
        selection and copying — only lets navigation/selection/copy key
        combos through, blocks anything that would modify the text.
        """
        allowed_keysyms = {
            'Left', 'Right', 'Home', 'End', 'Tab',
            'Shift_L', 'Shift_R', 'Control_L', 'Control_R'
        }
        ctrl_held = bool(event.state & 0x4)
        if event.keysym in allowed_keysyms:
            return None
        if ctrl_held and event.keysym.lower() in ('c', 'a'):
            return None
        return "break"

    def create_config_file_section(self):
        config_frame = ttk.Frame(self.main_frame, style='Dark.TFrame')
        config_frame.pack(fill=tk.X, pady=(5, 10))
        self._make_draggable(config_frame)

        label = tk.Label(config_frame, bg=self.colors['bg'], fg=self.colors['text'])
        label.pack(side=tk.LEFT)
        self.register_label(label, 'config_file')
        self._make_draggable(label)

        # config_path_var holds the real, full path and is what's actually
        # used for file I/O. config_display_var mirrors just the filename
        # for the UI — showing a long absolute path here added noise
        # without helping the user.
        self.config_path_var = tk.StringVar(value=self.config_file)
        self.config_display_var = tk.StringVar()
        self.config_path_var.trace_add('write', lambda *args: self._on_config_path_changed())
        self._on_config_path_changed()  # initial sync — trace doesn't fire for the constructor value

        self.config_entry = self.create_dark_entry(config_frame, width=24)
        self.config_entry.config(textvariable=self.config_display_var)
        self.config_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        # Read-only display: typing a bare filename here wouldn't be a
        # loadable path anymore, so block edits but still allow selecting
        # and copying the text (e.g. Ctrl+C, Ctrl+A).
        self.config_entry.bind('<Key>', self._block_entry_edit)

        # One button instead of two: opens the file picker and immediately
        # loads whatever config file the user selects, instead of making
        # them click "..." to browse and then "Load" to actually load it.
        load_btn = self.create_dark_button(config_frame, "", self.browse_config, width=10)
        self.register_button(load_btn, 'load')
        load_btn.config(text=TRANSLATIONS[self.current_language]['load'])
        load_btn.pack(side=tk.LEFT)

        checkboxes_frame = ttk.Frame(self.main_frame, style='Dark.TFrame')
        checkboxes_frame.pack(fill=tk.X, pady=(0, 5))
        self._make_draggable(checkboxes_frame)

        self.auto_apply_check = ttk.Checkbutton(
            checkboxes_frame,
            variable=self.auto_apply_var,
            style='Dark.TCheckbutton',
            command=self.mark_dirty
        )
        self.register_check(self.auto_apply_check, 'auto_apply')
        self.auto_apply_check.config(text=TRANSLATIONS[self.current_language]['auto_apply'])
        self.auto_apply_check.pack(side=tk.LEFT)
        self._add_tooltip(self.auto_apply_check, 'tooltip_auto_apply')

        self.tight_windows_check = ttk.Checkbutton(
            checkboxes_frame,
            variable=self.tight_windows_var,
            style='Dark.TCheckbutton',
            command=self.mark_dirty
        )
        self.register_check(self.tight_windows_check, 'tight_windows')
        self.tight_windows_check.config(text=TRANSLATIONS[self.current_language]['tight_windows'])
        self.tight_windows_check.pack(side=tk.LEFT, padx=(20, 0))
        self._add_tooltip(self.tight_windows_check, 'tooltip_tight_windows')

        self.consider_taskbar_check = ttk.Checkbutton(
            checkboxes_frame,
            variable=self.consider_taskbar_var,
            style='Dark.TCheckbutton',
            command=self.mark_dirty
        )
        self.register_check(self.consider_taskbar_check, 'consider_taskbar')
        self.consider_taskbar_check.config(text=TRANSLATIONS[self.current_language]['consider_taskbar'])
        self.consider_taskbar_check.pack(side=tk.LEFT, padx=(20, 0))
        self._add_tooltip(self.consider_taskbar_check, 'tooltip_consider_taskbar')

    def _add_tooltip(self, widget, translation_key):
        """
        GUI improvement #8: lightweight tooltip on hover for checkboxes.
        Takes a translation key (not raw text) and looks it up in the
        current language on every hover, so the tooltip text switches
        along with the rest of the UI when the language is toggled.
        """
        tooltip = {'win': None}

        def show(_event):
            if tooltip['win'] is not None:
                return
            text = TRANSLATIONS[self.current_language][translation_key]
            x = widget.winfo_rootx() + 10
            y = widget.winfo_rooty() + widget.winfo_height() + 5
            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.wm_geometry(f"+{x}+{y}")
            label = tk.Label(win, text=text, bg='#FFFFE0', fg='black',
                              relief='solid', borderwidth=1, padx=8, pady=5,
                              font=FONT_TOOLTIP, wraplength=320, justify='left')
            label.pack()
            tooltip['win'] = win

        def hide(_event):
            if tooltip['win'] is not None:
                tooltip['win'].destroy()
                tooltip['win'] = None

        widget.bind('<Enter>', show)
        widget.bind('<Leave>', hide)

    # ------------------------------------------------------------------
    def _monitor_window_counts(self):
        """How many configured window rows are currently assigned to each
        monitor — feeds the live count shown on the layout preview."""
        counts = [0] * len(self.monitor_frames)
        for entry in self.window_inputs:
            try:
                idx = int(entry.monitor_var.get().split()[0]) - 1
            except (ValueError, IndexError):
                continue
            if 0 <= idx < len(counts):
                counts[idx] += 1
        return counts

    def create_monitor_selection(self):
        monitors_frame = ttk.Frame(self.main_frame, style='Dark.TFrame')
        monitors_frame.pack(fill=tk.X)
        self._make_draggable(monitors_frame)

        # GUI improvement #4: visual layout preview — now shows resolution,
        # the Primary badge, and a live per-monitor window count directly
        # on each rectangle, so the separate "Monitor N (WxH): [Primary]"
        # text rows underneath (MonitorFrame) are no longer shown; they're
        # still created below purely to hold config data (get_config()).
        if self.monitor_info.monitors:
            self.layout_preview = MonitorLayoutPreview(
                monitors_frame, self.monitor_info.monitors, self.colors,
                translations={
                    'primary': TRANSLATIONS[self.current_language]['primary'],
                    'windows_label': TRANSLATIONS[self.current_language]['windows_label'],
                },
                get_window_counts=self._monitor_window_counts
            )
            self.layout_preview.pack(fill=tk.X, pady=(0, 5))

        for i, monitor in enumerate(self.monitor_info.monitors):
            monitor_frame = MonitorFrame(monitors_frame, monitor, i, self.colors,
                                          current_language=self.current_language,
                                          style='Dark.TFrame')
            # Not packed — kept only for its stored config data (left/top/
            # work area/etc, used by save_settings), the visible text row
            # it used to draw is gone now that the preview shows it.
            self.monitor_frames.append(monitor_frame)

    # ------------------------------------------------------------------
    def create_window_list_section(self):
        window_list_container = ttk.Frame(self.main_frame, style='Dark.TFrame')
        window_list_container.pack(fill=tk.BOTH, expand=True)

        header_frame = ttk.Frame(window_list_container, style='Dark.TFrame')
        header_frame.pack(fill=tk.X)
        self._make_draggable(header_frame)

        label = tk.Label(header_frame, bg=self.colors['bg'], fg=self.colors['text'])
        label.pack(side=tk.LEFT)
        self.register_label(label, 'window_titles')
        self._make_draggable(label)

        add_btn = self.create_dark_button(header_frame, "+", self.add_window_input, width=3)
        add_btn.pack(side=tk.LEFT, padx=5)

        self.create_scrollable_area(window_list_container)

    def create_scrollable_area(self, parent):
        self.canvas = tk.Canvas(parent,
                                 bg=self.colors['bg'],
                                 highlightthickness=0,
                                 height=10)
        self.scrollbar = ttk.Scrollbar(parent,
                                        orient="vertical",
                                        command=self.canvas.yview,
                                        style='Dark.Vertical.TScrollbar')

        self.scrollable_frame = ttk.Frame(self.canvas, style='Dark.TFrame')
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self._canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw", width=SCROLLABLE_FRAME_WIDTH
        )
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Code/GUI improvement #11: keep inner frame width in sync with the
        # canvas so the window-list area actually uses extra horizontal
        # space if the window is ever made wider (was previously fixed).
        #
        # Debounced: dragging the main window's edge fires <Configure> on
        # every single pixel. Each itemconfig(width=...) here forces Tk to
        # re-run pack layout for every row inside scrollable_frame (label +
        # combobox + entry + 3 buttons per window title row), which gets
        # visibly janky once there are more than a handful of rows.
        # Coalescing into one reflow shortly after resizing stops keeps the
        # window responsive while still ending up at the right width.
        self._width_sync_after_id = None
        self.canvas.bind("<Configure>", self._schedule_canvas_width_sync)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # GUI/code improvement #4: scope the mousewheel binding to this
        # canvas only (enter/leave), instead of bind_all which hijacked
        # scrolling everywhere in the app, including comboboxes.
        self.canvas.bind('<Enter>', lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind('<Leave>', lambda e: self.canvas.unbind_all("<MouseWheel>"))

    def _schedule_canvas_width_sync(self, event):
        pending_width = event.width
        if self._width_sync_after_id is not None:
            try:
                self.canvas.after_cancel(self._width_sync_after_id)
            except (ValueError, tk.TclError):
                pass
        self._width_sync_after_id = self.canvas.after(
            30, lambda: self._apply_canvas_width_sync(pending_width))

    def _apply_canvas_width_sync(self, width):
        self._width_sync_after_id = None
        try:
            self.canvas.itemconfig(self._canvas_window, width=width)
        except tk.TclError:
            pass

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # ------------------------------------------------------------------
    def create_control_buttons(self):
        buttons_frame = ttk.Frame(self.main_frame, style='Dark.TFrame')
        buttons_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))
        self._make_draggable(buttons_frame)

        centered_frame = ttk.Frame(buttons_frame, style='Dark.TFrame')
        centered_frame.pack(expand=True)

        self.lang_button = self.create_dark_button(
            centered_frame,
            "ENG" if self.current_language == 'ru' else "RUS",
            self.toggle_language,
            width=5
        )

        self.about_btn = self.create_dark_button(
            centered_frame,
            TRANSLATIONS[self.current_language]['about_button'],
            self.show_about,
            width=3
        )

        self.apply_btn = self.create_dark_button(centered_frame, "", self.arrange_windows, width=15)
        self.register_button(self.apply_btn, 'apply')
        self.apply_btn.config(text=TRANSLATIONS[self.current_language]['apply'])

        save_btn = self.create_dark_button(centered_frame, "", self.save_settings, width=17)
        self.register_button(save_btn, 'save_settings')
        save_btn.config(text=TRANSLATIONS[self.current_language]['save_settings'])

        exit_btn = self.create_dark_button(centered_frame, "", self.on_closing, width=15)
        self.register_button(exit_btn, 'exit')
        exit_btn.config(text=TRANSLATIONS[self.current_language]['exit'])

        self.lang_button.pack(side=tk.LEFT, padx=5)
        self.about_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.apply_btn.pack(side=tk.LEFT, padx=5)
        save_btn.pack(side=tk.LEFT, padx=5)
        # GUI improvement #11: extra gap before Exit so it's visually
        # separated from Apply and less likely to be mis-clicked.
        exit_btn.pack(side=tk.LEFT, padx=(20, 5))

    def _add_resize_grip(self):
        """
        A small diagonal-lines handle in the bottom-right corner — the
        standard visual cue that a window can be resized by dragging that
        corner (the window already is resizable via any edge; this just
        makes it obvious at a glance).
        """
        grip = ttk.Sizegrip(self.root, style='Dark.TSizegrip')
        grip.place(relx=1.0, rely=1.0, anchor='se')

    def show_about(self):
        """"(?)" button — custom-styled About dialog matching the app's dark theme."""
        lang = self.current_language

        about = tk.Toplevel(self.root)
        about.title(TRANSLATIONS[lang]['about_title'])
        about.configure(bg=self.colors['bg'])
        about.resizable(False, False)
        about.transient(self.root)
        about.attributes('-topmost', True)

        container = ttk.Frame(about, style='Dark.TFrame')
        container.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)

        title_label = tk.Label(container,
                                text=APP_NAME,
                                bg=self.colors['bg'],
                                fg=self.colors['text'],
                                font=(FONT_FAMILY, 14, 'bold'))
        title_label.pack(anchor='w', pady=(0, 12))

        body_label = tk.Label(container,
                               text=TRANSLATIONS[lang]['about_text'],
                               bg=self.colors['bg'],
                               fg=self.colors['text'],
                               font=FONT_NORMAL,
                               justify='left',
                               wraplength=380)
        body_label.pack(anchor='w', pady=(0, 16))

        contact_frame = ttk.Frame(container, style='Dark.TFrame')
        contact_frame.pack(anchor='w', pady=(0, 20))

        contact_caption = tk.Label(contact_frame,
                                    text=TRANSLATIONS[lang]['about_contact_label'],
                                    bg=self.colors['bg'],
                                    fg=self.colors['text'],
                                    font=FONT_NORMAL)
        contact_caption.pack(side=tk.LEFT, padx=(0, 6))

        contact_link = tk.Label(contact_frame,
                                 text=ABOUT_CONTACT_HANDLE,
                                 bg=self.colors['bg'],
                                 fg='#6EA8FE',
                                 font=(FONT_FAMILY, 11, 'underline'),
                                 cursor='hand2')
        contact_link.pack(side=tk.LEFT)
        contact_link.bind('<Button-1>', lambda e: webbrowser.open(ABOUT_CONTACT_URL))
        contact_link.bind('<Enter>', lambda e: contact_link.config(fg='#9CC4FF'))
        contact_link.bind('<Leave>', lambda e: contact_link.config(fg='#6EA8FE'))

        ok_btn = self.create_dark_button(container, "OK", about.destroy, width=10)
        ok_btn.pack(anchor='e')

        about.update_idletasks()
        # Center the About window over the main window.
        px = self.root.winfo_rootx() + (self.root.winfo_width() - about.winfo_width()) // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - about.winfo_height()) // 2
        about.geometry(f"+{max(px, 0)}+{max(py, 0)}")

        about.bind('<Escape>', lambda e: about.destroy())
        about.protocol("WM_DELETE_WINDOW", about.destroy)
        ok_btn.focus_set()
        about.grab_set()

    def create_dark_button(self, parent, text, command, width=None):
        return tk.Button(parent, text=text, command=command,
                          bg=self.colors['button_bg'],
                          fg=self.colors['text'],
                          activebackground=self.colors['button_active'],
                          activeforeground=self.colors['text'],
                          relief='raised',
                          font=FONT_BUTTON,
                          padx=6, pady=4,
                          width=width)

    def _get_gear_icon(self, size=22):
        """
        A small flat gear icon, drawn pixel-by-pixel with plain Tkinter
        (no PIL / external image files needed). Replaces the Unicode "⚙"
        glyph, which at this small a size renders in some fonts as an
        indistinct blob rather than a recognizable gear. Edges are
        supersampled and blended toward the button's background color for
        anti-aliasing (Tk 8.6's photo images don't support alpha via
        put(), so true transparency is only used for fully-empty pixels).
        Built once and cached - every "⚙" button shares the same image.
        """
        if getattr(self, '_gear_icon_cache', None) is not None:
            return self._gear_icon_cache

        img = tk.PhotoImage(width=size, height=size)
        cx = cy = size / 2
        outer_r = size * 0.47   # tip of each tooth
        body_r = size * 0.31    # solid disc radius (between teeth)
        hole_r = size * 0.14    # hollow center
        teeth = 8
        duty = 0.5              # fraction of each tooth's angular slot that's filled
        slot = (2 * math.pi) / teeth
        supersample = 4         # subsamples per axis, for anti-aliased edges

        fg = self.colors['text']
        bg = self.colors['button_bg']
        fg_rgb = tuple(int(fg[i:i + 2], 16) for i in (1, 3, 5))
        bg_rgb = tuple(int(bg[i:i + 2], 16) for i in (1, 3, 5))

        def is_gear_point(px, py):
            dx, dy = px - cx, py - cy
            r = math.hypot(dx, dy)
            if r <= hole_r:
                return False
            if r <= body_r:
                return True
            if r <= outer_r:
                angle = math.atan2(dy, dx) % (2 * math.pi)
                return (angle % slot) / slot < duty
            return False

        for y in range(size):
            for x in range(size):
                hits = 0
                for sy in range(supersample):
                    for sx in range(supersample):
                        px = x + (sx + 0.5) / supersample
                        py = y + (sy + 0.5) / supersample
                        if is_gear_point(px, py):
                            hits += 1
                coverage = hits / (supersample * supersample)
                if coverage <= 0:
                    img.transparency_set(x, y, True)
                elif coverage >= 1:
                    img.put(fg, (x, y))
                else:
                    mixed = tuple(round(bg_rgb[i] + (fg_rgb[i] - bg_rgb[i]) * coverage)
                                   for i in range(3))
                    img.put("#%02x%02x%02x" % mixed, (x, y))

        self._gear_icon_cache = img
        return img

    def create_dark_entry(self, parent, width=None):
        return tk.Entry(parent,
                         bg=self.colors['input_bg'],
                         fg=self.colors['text'],
                         insertbackground=self.colors['text'],
                         relief='solid',
                         font=FONT_NORMAL,
                         width=width)

    def _make_draggable(self, widget):
        """
        Lets the user move the whole app window by clicking and dragging
        on `widget` — not just the OS title bar. Bound only to
        non-interactive surfaces (background frames, static labels), never
        to buttons/entries/checkboxes/comboboxes, so normal clicks on
        controls are unaffected.
        """
        drag_state = {'x': 0, 'y': 0}

        def start_drag(event):
            drag_state['x'] = event.x_root - self.root.winfo_x()
            drag_state['y'] = event.y_root - self.root.winfo_y()

        def do_drag(event):
            x = event.x_root - drag_state['x']
            y = event.y_root - drag_state['y']
            self.root.geometry(f"+{x}+{y}")

        widget.bind('<ButtonPress-1>', start_drag)
        widget.bind('<B1-Motion>', do_drag)

    # ------------------------------------------------------------------
    def browse_config(self):
        filename = filedialog.askopenfilename(
            title=TRANSLATIONS[self.current_language]['select_config'],
            filetypes=[("INI files", "*.ini")]
        )
        if filename:
            self.config_path_var.set(filename)
            self.load_selected_config()

    def load_selected_config(self):
        new_config_path = self.config_path_var.get()
        if os.path.exists(new_config_path):
            self.config_file = new_config_path
            for entry in list(self.window_inputs):
                entry.master.master.destroy()
            self.window_inputs.clear()

            self.load_config()
            self.load_window_titles()
            self.load_monitor_settings()
            self._apply_saved_geometry()
            self.mark_clean()
            self._remember_last_config_path(new_config_path)
        else:
            messagebox.showerror(TRANSLATIONS[self.current_language]['error'],
                                  TRANSLATIONS[self.current_language]['config_not_found'])

    def load_monitor_settings(self):
        # Bug fix #1: width_var/height_var now actually exist on MonitorFrame.
        for i, monitor_frame in enumerate(self.monitor_frames):
            width = self.config.get('Monitors', f'Monitor{i}_Width',
                                     fallback=str(monitor_frame.get_config()['width']))
            height = self.config.get('Monitors', f'Monitor{i}_Height',
                                      fallback=str(monitor_frame.get_config()['height']))

            monitor_frame.width_var.set(width)
            monitor_frame.height_var.set(height)

    def _resolve_last_config_path(self):
        """
        Peeks into the bootstrap settings.ini (next to the program) just
        for a 'LastConfigPath' pointer, without disturbing self.config.
        Returns that path if it points to a file that still exists,
        otherwise falls back to the bootstrap file itself.
        """
        if not os.path.exists(self.bootstrap_config_path):
            return self.bootstrap_config_path
        try:
            probe = configparser.ConfigParser()
            probe.read(self.bootstrap_config_path)
            last_used = probe.get('Settings', 'LastConfigPath', fallback=self.bootstrap_config_path)
            if os.path.exists(last_used):
                return last_used
        except (OSError, configparser.Error) as e:
            self.logger.warning(f"Could not read last-used config pointer: {e}")
        return self.bootstrap_config_path

    def _remember_last_config_path(self, path):
        """
        Records which config file is currently in use, so the next launch
        reopens it automatically instead of always defaulting to the
        bootstrap settings.ini. Stored inside the bootstrap file itself,
        separately from whatever config content it may also hold.
        """
        try:
            pointer_cfg = configparser.ConfigParser()
            if os.path.exists(self.bootstrap_config_path):
                pointer_cfg.read(self.bootstrap_config_path)
            if not pointer_cfg.has_section('Settings'):
                pointer_cfg.add_section('Settings')
            pointer_cfg.set('Settings', 'LastConfigPath', path)
            os.makedirs(os.path.dirname(self.bootstrap_config_path), exist_ok=True)
            with open(self.bootstrap_config_path, 'w') as f:
                pointer_cfg.write(f)
        except OSError as e:
            self.logger.error(f"Could not remember last-used config path: {e}")

    def load_config(self):
        self.config.read(self.config_file)
        for section in ('Monitors', 'Windows', 'Settings', 'Language'):
            if not self.config.has_section(section):
                self.config.add_section(section)

        self.auto_apply_var.set(self.config.getboolean('Settings', 'AutoApply', fallback=False))
        self.tight_windows_var.set(self.config.getboolean('Settings', 'TightWindows', fallback=False))
        self.consider_taskbar_var.set(self.config.getboolean('Settings', 'ConsiderTaskbar', fallback=True))

        saved_language = self.config.get('Language', 'Language', fallback='RUS')
        self.current_language = 'ru' if saved_language == 'RUS' else 'en'
        if hasattr(self, 'lang_button'):
            self.lang_button.config(text="ENG" if self.current_language == 'ru' else "RUS")

    def _apply_saved_geometry(self):
        """
        Restores the app window's own size/position from whichever config
        is currently loaded (self.config). Falls back to the built-in
        default if that config has no WindowGeometry yet. Called both at
        startup and whenever a different config file is loaded, so the
        window resizes to match the config actually in use.
        """
        saved_geometry = self.config.get('Settings', 'WindowGeometry', fallback='')
        self.root.geometry(saved_geometry if saved_geometry else DEFAULT_WINDOW_GEOMETRY)

    def _build_monitor_names(self, lang):
        """The translated option list for a window row's monitor dropdown
        (e.g. "1 [Primary]" / "1 [Основной]"). Shared by add_window_input
        (initial creation) and update_interface_language (refresh on
        language switch) so both stay in sync."""
        return [
            f"{i + 1}" + (f" [{TRANSLATIONS[lang]['primary']}]" if mf.is_primary else "")
            for i, mf in enumerate(self.monitor_frames)
        ]

    # ------------------------------------------------------------------
    def add_window_input(self, title=""):
        frame = ttk.Frame(self.scrollable_frame, style='Dark.TFrame')
        frame.pack(fill=tk.X, pady=4)

        if not title:
            existing_numbers = [int(inp.get().split('.')[0]) for inp in self.window_inputs
                                 if inp.get().split('.')[0].isdigit()]
            next_num = max(existing_numbers, default=0) + 1
            title = TRANSLATIONS[self.current_language]['notepad_title'].format(next_num)

        input_row = ttk.Frame(frame, style='Dark.TFrame')
        input_row.pack(fill=tk.X)

        monitor_label = tk.Label(input_row,
                                  text=TRANSLATIONS[self.current_language]['monitor'],
                                  bg=self.colors['bg'],
                                  fg=self.colors['text'],
                                  font=FONT_NORMAL)
        monitor_label.is_monitor_label = True  # used by update_interface_language
        monitor_label.pack(side=tk.LEFT, padx=(0, 5))

        monitor_names = self._build_monitor_names(self.current_language)
        monitor_var = tk.StringVar(value="1")
        monitor_dropdown = ttk.Combobox(input_row,
                                         textvariable=monitor_var,
                                         values=monitor_names or [str(i + 1) for i in range(len(self.monitor_frames))],
                                         width=8,
                                         state="readonly",
                                         style='Dark.TCombobox',
                                         font=FONT_NORMAL)
        monitor_dropdown.pack(side=tk.LEFT, padx=(0, 5), ipady=2)
        monitor_dropdown.bind("<<ComboboxSelected>>", self._on_monitor_assignment_changed)

        entry = self.create_dark_entry(input_row)
        entry.insert(0, title)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5), ipady=2)
        entry.monitor_var = monitor_var
        # Bug fix: needed so update_interface_language() can find and
        # refresh this dropdown's option list (it embeds the translated
        # "[Primary]"/"[Основной]" suffix, which was previously frozen in
        # whichever language was active when the row was created).
        entry.monitor_dropdown = monitor_dropdown

        # Per-window custom size/position override (set via the "⚙" dialog).
        # When use_custom_var is False, this window is arranged
        # automatically in the grid as usual.
        entry.use_custom_var = BooleanVar(value=False)
        entry.custom_width_var = tk.IntVar(value=800)
        entry.custom_height_var = tk.IntVar(value=600)
        entry.custom_x_var = tk.IntVar(value=100)
        entry.custom_y_var = tk.IntVar(value=100)

        # GUI improvement #2: live "not found" highlight instead of only
        # discovering missing windows after clicking Apply.
        entry.bind('<KeyRelease>', lambda e, ent=entry: self._on_entry_changed(ent))
        entry.bind('<FocusOut>', lambda e, ent=entry: self._on_entry_changed(ent))

        select_btn = self.create_dark_button(input_row, "...",
                                              lambda ent=entry: self.show_window_titles_menu(ent),
                                              width=3)
        select_btn.pack(side=tk.RIGHT, padx=(0, 10))

        # GUI improvement #5: extra spacing between "..." and "X" so they
        # aren't misclicked, plus a confirmation before removing a row.
        delete_btn = self.create_dark_button(input_row, "X",
                                              lambda: self.remove_window_input(frame, entry),
                                              width=3)
        delete_btn.pack(side=tk.RIGHT, padx=(8, 0))

        gear_btn = self.create_dark_button(input_row, "",
                                            lambda ent=entry: self.show_custom_geometry_dialog(ent),
                                            width=3)
        gear_icon = self._get_gear_icon()
        gear_btn.config(image=gear_icon, width=28, height=28)
        gear_btn.image = gear_icon  # keep a reference so it isn't garbage-collected
        gear_btn.pack(side=tk.RIGHT)

        self.window_inputs.append(entry)
        self.mark_dirty()
        self._refresh_window_titles_validity()
        if hasattr(self, 'layout_preview'):
            self.layout_preview.refresh()

    def _on_monitor_assignment_changed(self, event=None):
        self.mark_dirty()
        if hasattr(self, 'layout_preview'):
            self.layout_preview.refresh()

    def _on_entry_changed(self, entry):
        """
        Debounced: this fires on every <KeyRelease>, and validity refresh
        does a full EnumWindows-based scan. Without debouncing, typing a
        single window title fired one full scan per keystroke. Collapse
        rapid-fire calls into a single scan ~250ms after the user stops
        typing.
        """
        self.mark_dirty()
        if self._validity_after_id is not None:
            try:
                self.root.after_cancel(self._validity_after_id)
            except (ValueError, tk.TclError):
                pass
        self._validity_after_id = self.root.after(250, self._refresh_window_titles_validity)

    def _count_matching_windows(self, title, snapshot=None):
        """
        How many currently open windows' text contains `title` as a
        substring. Accepts an optional pre-collected snapshot (see
        _snapshot_visible_windows) to avoid a fresh EnumWindows pass per
        call when checking many titles in one go.
        """
        title_lower = title.lower()

        if snapshot is not None:
            return sum(1 for _h, text in snapshot if title_lower in text.lower())

        count = 0

        def enum_cb(h, _):
            nonlocal count
            if win32gui.IsWindowVisible(h):
                text = win32gui.GetWindowText(h)
                if text and title_lower in text.lower():
                    count += 1

        win32gui.EnumWindows(enum_cb, None)
        return count

    def _refresh_window_titles_validity(self, snapshot=None):
        """
        Colors every window-title entry green or red. This has to look at
        ALL entries together, not one at a time: if five rows share the
        same title and only five matching windows are actually open, all
        five are fine (green) — but a sixth row with that same title has
        no window left for it and should show red, even though the title
        itself is "valid" and exists elsewhere in the list. This mirrors
        exactly how _arrange_windows_impl claims windows one-by-one per
        duplicate title, so the color always matches what Apply will do.

        Performance: takes a single EnumWindows snapshot up front (unless
        one is already supplied by the caller, e.g. _arrange_windows_impl
        reusing the snapshot it already took) instead of re-enumerating
        all windows once per distinct title.
        """
        self._validity_after_id = None

        if snapshot is None:
            snapshot = self._snapshot_visible_windows()

        seen_so_far = {}    # title (lowercased) -> how many rows with it we've passed
        available_cache = {}  # title (lowercased) -> how many real windows match it

        for entry in self.window_inputs:
            title = entry.get().strip()
            if not title:
                continue

            key = title.lower()
            rank = seen_so_far.get(key, 0) + 1
            seen_so_far[key] = rank

            if key not in available_cache:
                available_cache[key] = self._count_matching_windows(title, snapshot=snapshot)
            available = available_cache[key]

            will_be_found = rank <= available
            entry.config(bg=self.colors['found_bg'] if will_be_found else self.colors['error_bg'])

    def remove_window_input(self, frame, entry):
        self.window_inputs.remove(entry)
        frame.destroy()
        self.mark_dirty()
        if hasattr(self, 'layout_preview'):
            self.layout_preview.refresh()

    def load_window_titles(self):
        count = self.config.getint('Windows', 'Count', fallback=0)
        if count == 0:
            self.add_window_input()
        else:
            for i in range(1, count + 1):
                title = self.config.get('Windows', f'Window{i}', fallback=f"{i}.txt - Notepad")
                monitor = self.config.get('Windows', f'Window{i}_Monitor', fallback="1")
                self.add_window_input(title)
                new_entry = self.window_inputs[-1]
                # Cosmetic: only the bare number is saved, but the dropdown's
                # own option list uses "N [Primary]"/"N [Основной]" for the
                # primary monitor. The bare number still works fine
                # everywhere it's parsed (int(...).split()[0]) - this just
                # makes a loaded row display identically to one where the
                # user picked the same monitor by hand from the dropdown.
                try:
                    monitor_idx = int(monitor.split()[0]) - 1
                except (ValueError, IndexError):
                    monitor_idx = None
                monitor_names = self._build_monitor_names(self.current_language)
                if monitor_idx is not None and 0 <= monitor_idx < len(monitor_names):
                    new_entry.monitor_var.set(monitor_names[monitor_idx])
                else:
                    new_entry.monitor_var.set(monitor)
                new_entry.use_custom_var.set(
                    self.config.getboolean('Windows', f'Window{i}_UseCustom', fallback=False))
                new_entry.custom_width_var.set(
                    self.config.getint('Windows', f'Window{i}_CustomWidth', fallback=800))
                new_entry.custom_height_var.set(
                    self.config.getint('Windows', f'Window{i}_CustomHeight', fallback=600))
                new_entry.custom_x_var.set(
                    self.config.getint('Windows', f'Window{i}_CustomX', fallback=100))
                new_entry.custom_y_var.set(
                    self.config.getint('Windows', f'Window{i}_CustomY', fallback=100))
        self.mark_clean()  # loading isn't a user edit

    def get_window_titles(self):
        titles = []

        def enum_window_callback(hwnd, titles):
            if win32gui.IsWindowVisible(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                if window_title:
                    titles.append(window_title)

        win32gui.EnumWindows(enum_window_callback, titles)
        return sorted(titles)

    def show_window_titles_menu(self, entry):
        """
        GUI improvement #2: replaced the flat tk.Menu (no search, hard to
        use with many open windows) with a small filterable popup listbox.
        """
        self._close_title_picker_popup()

        titles = self.get_window_titles()

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.configure(bg=self.colors['bg'])
        self.window_titles_menu = popup

        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        popup.geometry(f"440x280+{x}+{y}")

        filter_var = tk.StringVar()
        filter_entry = self.create_dark_entry(popup)
        filter_entry.config(textvariable=filter_var)
        filter_entry.pack(fill=tk.X, padx=6, pady=6, ipady=3)
        filter_entry.focus_set()

        listbox = tk.Listbox(popup,
                              bg=self.colors['input_bg'],
                              fg=self.colors['text'],
                              selectbackground=self.colors['button_active'],
                              highlightthickness=0,
                              relief='flat',
                              font=FONT_NORMAL)
        listbox.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))

        def populate(filter_text=""):
            listbox.delete(0, tk.END)
            for t in titles:
                if filter_text.lower() in t.lower():
                    listbox.insert(tk.END, t)

        populate()

        def on_filter_change(*_args):
            populate(filter_var.get())

        filter_var.trace_add('write', on_filter_change)

        def choose(_event=None):
            selection = listbox.curselection()
            if selection:
                self.select_window_title(entry, listbox.get(selection[0]))

        # Bug fix: a <FocusOut> binding on the popup itself was destroying
        # the window mid-click (focus shifts internally when you click the
        # listbox), so the click never reached the selection handler. Select
        # on a single click release instead of double-click, and only close
        # the popup explicitly (Escape, or after a successful selection) —
        # no more auto-destroy on focus changes.
        listbox.bind('<ButtonRelease-1>', choose)
        filter_entry.bind('<Return>', lambda e: (listbox.selection_set(0), choose()) if listbox.size() else None)
        filter_entry.bind('<Down>', lambda e: (listbox.focus_set(), listbox.selection_set(0)) if listbox.size() else None)
        popup.bind('<Escape>', lambda e: self._close_title_picker_popup())

        # Close the popup if the user clicks somewhere else entirely
        # (outside both the popup and the entry that opened it), without
        # relying on fragile FocusOut semantics.
        def _close_if_click_outside(event):
            widget_under_click = event.widget
            if popup.winfo_exists() and widget_under_click not in (popup, filter_entry, listbox):
                # Only close if the click landed outside the popup's own tree
                x, y = event.x_root, event.y_root
                px, py = popup.winfo_rootx(), popup.winfo_rooty()
                pw, ph = popup.winfo_width(), popup.winfo_height()
                if not (px <= x <= px + pw and py <= y <= py + ph):
                    self._close_title_picker_popup()

        self._picker_outside_bind_id = self.root.bind('<Button-1>', _close_if_click_outside, add='+')

    def select_window_title(self, entry, title):
        entry.delete(0, tk.END)
        entry.insert(0, title)
        self._close_title_picker_popup()
        self._on_entry_changed(entry)

    def show_custom_geometry_dialog(self, entry):
        """
        "⚙" button: lets the user pin an exact width/height/position for
        this one window (via sliders) instead of letting it fall into the
        automatic grid layout. Bounds are based on whichever monitor is
        currently selected for this row. Dragging a slider live-moves the
        real window (if it's currently open) so the effect is visible
        immediately, not just after Apply.
        """
        lang = self.current_language

        try:
            monitor_index = int(entry.monitor_var.get().split()[0]) - 1
        except (ValueError, IndexError):
            monitor_index = 0
        if 0 <= monitor_index < len(self.monitor_info.monitors):
            bounds = self.monitor_info.monitors[monitor_index]
        else:
            bounds = {'physical_width': 1920, 'physical_height': 1080, 'left': 0, 'top': 0}

        # Resolve the live target window once, up front, and remember its
        # current geometry so we can put it back if the user closes this
        # dialog without actually turning the override on.
        target_hwnd = self._find_window_by_title(entry.get().strip())
        original_rect = None
        if target_hwnd:
            try:
                original_rect = win32gui.GetWindowRect(target_hwnd)  # (left, top, right, bottom)
            except Exception:
                target_hwnd = None

        dialog = tk.Toplevel(self.root)
        dialog.title(TRANSLATIONS[lang]['custom_geometry_title'])
        dialog.configure(bg=self.colors['bg'])
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        container = ttk.Frame(dialog, style='Dark.TFrame')
        container.pack(fill=tk.BOTH, expand=True, padx=20, pady=16)

        if not target_hwnd:
            tk.Label(container, text=TRANSLATIONS[lang]['live_preview_unavailable'],
                     bg=self.colors['bg'], fg='#E67E73',
                     font=FONT_NORMAL).pack(anchor='w', pady=(0, 10))

        def live_update(*_args):
            # Touching a slider signals intent to actually use these
            # values — flip the checkbox on automatically (the "Clear"
            # button is the explicit way back out).
            entry.use_custom_var.set(True)
            self.mark_dirty()
            if target_hwnd:
                try:
                    w = max(entry.custom_width_var.get(), 50)
                    h = max(entry.custom_height_var.get(), 50)
                    x = bounds['left'] + entry.custom_x_var.get()
                    y = bounds['top'] + entry.custom_y_var.get()
                    win32gui.SetWindowPos(target_hwnd, win32con.HWND_TOP, x, y, w, h,
                                           win32con.SWP_NOZORDER | win32con.SWP_SHOWWINDOW)
                except Exception:
                    pass

        use_custom_check = tk.Checkbutton(
            container,
            text=TRANSLATIONS[lang]['use_custom_geometry'],
            variable=entry.use_custom_var,
            bg=self.colors['bg'], fg=self.colors['text'],
            activebackground=self.colors['bg'], activeforeground=self.colors['text'],
            selectcolor=self.colors['input_bg'],
            highlightthickness=0, font=FONT_NORMAL,
            command=lambda: (self.mark_dirty(), live_update())
        )
        use_custom_check.pack(anchor='w', pady=(0, 10))

        def add_slider(label_key, variable, max_value):
            tk.Label(container, text=TRANSLATIONS[lang][label_key],
                     bg=self.colors['bg'], fg=self.colors['text'],
                     font=FONT_NORMAL).pack(anchor='w')
            row = ttk.Frame(container, style='Dark.TFrame')
            row.pack(fill=tk.X, pady=(0, 10))
            scale = tk.Scale(row, from_=0, to=max(max_value, 1), orient=tk.HORIZONTAL,
                              variable=variable, bg=self.colors['bg'], fg=self.colors['text'],
                              troughcolor=self.colors['input_bg'],
                              highlightthickness=0, activebackground=self.colors['button_active'],
                              length=340, font=FONT_NORMAL,
                              command=live_update)
            scale.pack(fill=tk.X)
            return scale

        add_slider('custom_width', entry.custom_width_var, bounds['physical_width'])
        add_slider('custom_height', entry.custom_height_var, bounds['physical_height'])
        add_slider('custom_x', entry.custom_x_var, bounds['physical_width'])
        add_slider('custom_y', entry.custom_y_var, bounds['physical_height'])

        btn_frame = ttk.Frame(container, style='Dark.TFrame')
        btn_frame.pack(fill=tk.X, pady=(6, 0))

        def restore_if_not_custom():
            """Snaps the live window back to where it was before this
            dialog opened, but only if the override was never enabled —
            once "use custom" is on, the live position IS the intended
            result, so it's left alone."""
            if target_hwnd and original_rect and not entry.use_custom_var.get():
                try:
                    left, top, right, bottom = original_rect
                    win32gui.SetWindowPos(target_hwnd, win32con.HWND_TOP, left, top,
                                           right - left, bottom - top,
                                           win32con.SWP_NOZORDER | win32con.SWP_SHOWWINDOW)
                except Exception:
                    pass

        def clear_override():
            entry.use_custom_var.set(False)
            self.mark_dirty()
            restore_if_not_custom()

        def close_dialog():
            restore_if_not_custom()
            dialog.destroy()

        clear_btn = self.create_dark_button(btn_frame, TRANSLATIONS[lang]['clear_override'],
                                             clear_override, width=20)
        clear_btn.pack(side=tk.LEFT)

        close_btn = self.create_dark_button(btn_frame, TRANSLATIONS[lang]['close'],
                                             close_dialog, width=12)
        close_btn.pack(side=tk.RIGHT)

        dialog.update_idletasks()
        px = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(px, 0)}+{max(py, 0)}")

        dialog.bind('<Escape>', lambda e: close_dialog())
        dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        dialog.grab_set()

    # ------------------------------------------------------------------
    def arrange_windows(self):
        count = len(self.window_inputs)
        if count == 0:
            if self.headless:
                self.logger.warning("No window titles configured — nothing to arrange.")
            else:
                messagebox.showinfo(TRANSLATIONS[self.current_language]['attention'],
                                     TRANSLATIONS[self.current_language]['add_title'])
            return

        self._arrange_windows_impl()
        # Windows sometimes ignores the first SetWindowPos right after a
        # SW_RESTORE; re-applying shortly after makes layout reliable.
        # report=False: don't log/toast the same "not found" list twice.
        #self.root.after(ARRANGE_REAPPLY_DELAY_MS, lambda: self._arrange_windows_impl(report=False))

    def _snapshot_visible_windows(self):
        """
        One single EnumWindows pass over all top-level visible windows,
        collecting (hwnd, title) pairs. This used to be re-done from
        scratch — via a fresh EnumWindows + GetWindowText per call — for
        every single row in _find_window_by_title, and again for every
        distinct title in _count_matching_windows. With N configured
        windows and M open windows that was O(N*M) GetWindowText calls
        (each a cross-process SendMessage, the most expensive part of
        this). Taking one snapshot per arrange/validity pass and matching
        against it in memory makes it O(N+M) instead.
        """
        windows = []

        def enum_cb(h, _):
            if win32gui.IsWindowVisible(h):
                text = win32gui.GetWindowText(h)
                if text:
                    windows.append((h, text))

        win32gui.EnumWindows(enum_cb, None)
        return windows

    def _find_window_by_title(self, title, exclude=frozenset(), snapshot=None):
        """
        Bug fix #3: exact FindWindow match breaks the moment a title changes
        slightly (e.g. Notepad adding '*' for unsaved changes). Fall back to
        a substring search over all visible windows.

        `exclude` lets duplicate-titled rows (e.g. two Notepad windows both
        titled "Untitled - Notepad") each resolve to a *different* actual
        window instead of every row grabbing the same one — see
        _arrange_windows_impl, which tracks handles already claimed in the
        current pass and excludes them here.

        `snapshot` is an optional pre-collected list of (hwnd, title) pairs
        from _snapshot_visible_windows(). When the caller is doing this for
        many rows in one pass (arranging, or refreshing validity), it
        should take one snapshot up front and pass it in here every time,
        instead of letting this method re-enumerate all windows on every
        call.
        """
        hwnd = win32gui.FindWindow(None, title)
        if hwnd and hwnd not in exclude:
            return hwnd

        title_lower = title.lower()

        if snapshot is not None:
            for h, text in snapshot:
                if h not in exclude and title_lower in text.lower():
                    return h
            return None

        # No snapshot supplied (e.g. a one-off lookup like the "⚙" custom
        # geometry dialog) — fall back to a fresh enumeration.
        matches = []

        def enum_cb(h, _):
            if win32gui.IsWindowVisible(h) and h not in exclude:
                text = win32gui.GetWindowText(h)
                if text and title_lower in text.lower():
                    matches.append(h)

        win32gui.EnumWindows(enum_cb, None)
        return matches[0] if matches else None

    def _position_window(self, hwnd, x, y, width, height, window_title):
        """
        Restores/positions/raises one window — shared by both the
        automatic grid layout and custom per-window overrides. A single
        SetWindowPos(HWND_TOP, ...) call per window; no topmost/notopmost
        toggle and no SetForegroundWindow — both were extra Win32 round
        trips that aren't needed just to lay windows out.
        """
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        try:
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y,
                                   width, height, win32con.SWP_SHOWWINDOW)
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        except Exception as e:
            self.logger.error(f"Failed to position window '{window_title}': {e}")

    def _arrange_windows_impl(self, report=True):
        windows_by_monitor = {}
        for entry in self.window_inputs:
            # Bug fix: every other place that parses monitor_var guards
            # against a malformed/empty value (e.g. a hand-edited or
            # corrupted settings.ini) with try/except - this path, the
            # actual Apply logic, didn't, so a bad value here used to
            # raise inside a Tk callback and silently do nothing (no
            # error shown, nothing logged in a windowed build where
            # stdout/stderr may not exist).
            try:
                monitor_index = int(entry.monitor_var.get().split()[0]) - 1
            except (ValueError, IndexError):
                self.logger.warning(
                    f"Skipping window '{entry.get()}': invalid monitor value "
                    f"'{entry.monitor_var.get()}'")
                continue
            windows_by_monitor.setdefault(monitor_index, []).append(entry)

        window_not_found = []
        # Tracks windows already claimed in this pass so that duplicate
        # titles (e.g. two windows both "Untitled - Notepad") each get a
        # different physical window instead of all rows grabbing the same one.
        claimed_handles = set()

        # Performance: take ONE EnumWindows snapshot for the whole arrange
        # pass instead of letting _find_window_by_title re-enumerate every
        # visible window on every single lookup. This is the single
        # biggest win for large window lists / busy desktops.
        snapshot = self._snapshot_visible_windows()

        for monitor_index, monitor_windows in windows_by_monitor.items():
            if monitor_index >= len(self.monitor_frames):
                continue

            monitor_info = self.monitor_info.monitors[monitor_index]

            # Rows with a "⚙" custom size/position override are positioned
            # exactly where the user pinned them, and don't take part in —
            # or skew — the automatic grid math for the remaining rows.
            auto_windows = [e for e in monitor_windows if not e.use_custom_var.get()]
            custom_windows = [e for e in monitor_windows if e.use_custom_var.get()]

            windows_count = len(auto_windows)

            if windows_count:
                windows_per_row = math.ceil(math.sqrt(windows_count))
                windows_per_column = math.ceil(windows_count / windows_per_row)

                # Bug fix: this used to derive a single "taskbar_height" as
                # physical_height - work_height and always place windows
                # starting at the monitor's physical left/top. That only
                # works when the taskbar is docked at the bottom. Docked at
                # top, windows started under the taskbar and the bottom row
                # ran past the visible area; docked left/right,
                # physical_height == work_height so taskbar_height came out
                # to 0 and windows ignored it entirely (started under a
                # side-docked taskbar, sized as if it weren't there). Using
                # the monitor's work-area rect directly handles all four
                # edges correctly, since Windows always shrinks/offsets the
                # work area away from wherever the taskbar actually is.
                if self.consider_taskbar_var.get():
                    area_left = monitor_info['work_left']
                    area_top = monitor_info['work_top']
                    area_width = monitor_info['work_width']
                    area_height = monitor_info['work_height']
                else:
                    area_left = monitor_info['left']
                    area_top = monitor_info['top']
                    area_width = monitor_info['physical_width']
                    area_height = monitor_info['physical_height']

                if self.tight_windows_var.get():
                    overlap = int(TIGHT_OVERLAP_BASE_PX * (monitor_info.get('scale', 100) / 100))
                    window_width = (area_width + (windows_per_row - 1) * overlap) // windows_per_row
                    window_height = (area_height +
                                      (windows_per_column - 1) * overlap) // windows_per_column
                    spacing_x = -overlap
                    spacing_y = TIGHT_SPACING_Y
                else:
                    total_spacing_x = (windows_per_row - 1) * DEFAULT_SPACING_PX
                    total_spacing_y = (windows_per_column - 1) * DEFAULT_SPACING_PX
                    window_width = (area_width - total_spacing_x) // windows_per_row
                    window_height = (area_height - total_spacing_y) // windows_per_column
                    spacing_x = DEFAULT_SPACING_PX
                    spacing_y = DEFAULT_SPACING_PX

                for i, entry in enumerate(auto_windows):
                    window_title = entry.get()
                    hwnd = self._find_window_by_title(window_title, exclude=claimed_handles, snapshot=snapshot)

                    if hwnd:
                        claimed_handles.add(hwnd)
                        row = i // windows_per_row
                        col = i % windows_per_row
                        x = area_left + (col * (window_width + spacing_x))
                        y = area_top + (row * (window_height + spacing_y))
                        self._position_window(hwnd, x, y, window_width, window_height, window_title)
                    else:
                        window_not_found.append(window_title)

            for entry in custom_windows:
                window_title = entry.get()
                hwnd = self._find_window_by_title(window_title, exclude=claimed_handles, snapshot=snapshot)

                if hwnd:
                    claimed_handles.add(hwnd)
                    width = max(entry.custom_width_var.get(), 50)
                    height = max(entry.custom_height_var.get(), 50)
                    x = monitor_info['left'] + entry.custom_x_var.get()
                    y = monitor_info['top'] + entry.custom_y_var.get()
                    self._position_window(hwnd, x, y, width, height, window_title)
                else:
                    window_not_found.append(window_title)

        self._refresh_window_titles_validity(snapshot=snapshot)

        if report and window_not_found:
            if self.headless:
                self.logger.warning("Windows not found: " + ", ".join(window_not_found))
            else:
                message = f"{TRANSLATIONS[self.current_language]['windows_not_found']}\n" + "\n".join(window_not_found)
                AutoCloseMessageBox(TRANSLATIONS[self.current_language]['attention'], message, AUTO_CLOSE_TOAST_MS)

    # ------------------------------------------------------------------
    def save_settings(self):
        """
        Explicit "Save settings" button. If the current config file already
        exists, first asks whether to overwrite it or save under a new
        name; otherwise goes straight to the "Save as" dialog.
        """
        current_path = self.config_path_var.get()

        if current_path and os.path.exists(current_path):
            choice = self._ask_overwrite_or_save_as(current_path)
            if choice == 'cancel':
                return
            if choice == 'overwrite':
                chosen_path = current_path
            else:
                chosen_path = self._prompt_save_as_path(current_path)
                if not chosen_path:
                    return
        else:
            chosen_path = self._prompt_save_as_path(current_path)
            if not chosen_path:
                return

        self.config_path_var.set(chosen_path)
        self.config_file = chosen_path
        self._persist_settings()

    def _prompt_save_as_path(self, current_path):
        initial_dir = os.path.dirname(current_path) or os.getcwd()
        initial_file = os.path.basename(current_path) or CONFIG_FILE_NAME
        return filedialog.asksaveasfilename(
            title=TRANSLATIONS[self.current_language]['save_as_title'],
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=".ini",
            filetypes=[("INI files", "*.ini")]
        )

    def _ask_overwrite_or_save_as(self, current_path):
        """
        Small custom modal (matches the app's dark theme) with three
        choices: overwrite the current file, save under a new name, or
        cancel. Blocks until the user picks one and returns
        'overwrite' | 'save_as' | 'cancel'.
        """
        lang = self.current_language
        result = {'choice': 'cancel'}

        dialog = tk.Toplevel(self.root)
        dialog.title(TRANSLATIONS[lang]['attention'])
        dialog.configure(bg=self.colors['bg'])
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.attributes('-topmost', True)

        container = ttk.Frame(dialog, style='Dark.TFrame')
        container.pack(fill=tk.BOTH, expand=True, padx=24, pady=20)

        message = TRANSLATIONS[lang]['overwrite_prompt'].format(os.path.basename(current_path))
        label = tk.Label(container, text=message,
                          bg=self.colors['bg'], fg=self.colors['text'],
                          font=FONT_NORMAL, justify='left', wraplength=380)
        label.pack(anchor='w', pady=(0, 16))

        btn_frame = ttk.Frame(container, style='Dark.TFrame')
        btn_frame.pack(fill=tk.X)

        def choose(value):
            result['choice'] = value
            dialog.destroy()

        overwrite_btn = self.create_dark_button(
            btn_frame, TRANSLATIONS[lang]['overwrite'], lambda: choose('overwrite'), width=14)
        overwrite_btn.pack(side=tk.LEFT, padx=(0, 8))

        save_as_btn = self.create_dark_button(
            btn_frame, TRANSLATIONS[lang]['save_as'], lambda: choose('save_as'), width=14)
        save_as_btn.pack(side=tk.LEFT, padx=(0, 8))

        cancel_btn = self.create_dark_button(
            btn_frame, TRANSLATIONS[lang]['cancel'], lambda: choose('cancel'), width=10)
        cancel_btn.pack(side=tk.LEFT)

        dialog.update_idletasks()
        px = self.root.winfo_rootx() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        py = self.root.winfo_rooty() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{max(px, 0)}+{max(py, 0)}")

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose('cancel'))
        dialog.bind('<Escape>', lambda e: choose('cancel'))
        dialog.grab_set()
        self.root.wait_window(dialog)

        return result['choice']

    def _persist_settings(self):
        """
        Writes the full current state to the config file: monitor info,
        window list + their monitor assignment, checkboxes, language, and
        the app window's own geometry — so the next launch can restore
        everything exactly as it was left, without asking anything.
        """
        config_path = self.config_path_var.get()

        for i, monitor_frame in enumerate(self.monitor_frames):
            monitor_config = monitor_frame.get_config()
            self.config.set('Monitors', f'Monitor{i}_Width', str(monitor_frame.monitor_info['physical_width']))
            self.config.set('Monitors', f'Monitor{i}_Height', str(monitor_frame.monitor_info['physical_height']))
            self.config.set('Monitors', f'Monitor{i}_Left', str(monitor_config['left']))
            self.config.set('Monitors', f'Monitor{i}_Top', str(monitor_config['top']))
            self.config.set('Monitors', f'Monitor{i}_WorkWidth', str(monitor_config['work_width']))
            self.config.set('Monitors', f'Monitor{i}_WorkHeight', str(monitor_config['work_height']))
            self.config.set('Monitors', f'Monitor{i}_WorkLeft', str(monitor_config['work_left']))
            self.config.set('Monitors', f'Monitor{i}_WorkTop', str(monitor_config['work_top']))
            self.config.set('Monitors', f'Monitor{i}_IsPrimary', str(monitor_config['is_primary']))

        self.config.set('Windows', 'Count', str(len(self.window_inputs)))
        for i, entry in enumerate(self.window_inputs, 1):
            self.config.set('Windows', f'Window{i}', entry.get())
            self.config.set('Windows', f'Window{i}_Monitor', entry.monitor_var.get().split()[0])
            self.config.set('Windows', f'Window{i}_UseCustom', str(entry.use_custom_var.get()))
            self.config.set('Windows', f'Window{i}_CustomWidth', str(entry.custom_width_var.get()))
            self.config.set('Windows', f'Window{i}_CustomHeight', str(entry.custom_height_var.get()))
            self.config.set('Windows', f'Window{i}_CustomX', str(entry.custom_x_var.get()))
            self.config.set('Windows', f'Window{i}_CustomY', str(entry.custom_y_var.get()))

        self.config.set('Settings', 'AutoApply', str(self.auto_apply_var.get()))
        self.config.set('Settings', 'TightWindows', str(self.tight_windows_var.get()))
        self.config.set('Settings', 'ConsiderTaskbar', str(self.consider_taskbar_var.get()))

        # Remember the app window's own size/position so it reopens the
        # same way next time. normal geometry (not while minimized).
        try:
            if self.root.state() != 'iconic':
                self.config.set('Settings', 'WindowGeometry', self.root.geometry())
        except tk.TclError:
            pass

        self.config.set('Language', 'Language', 'RUS' if self.current_language == 'ru' else 'ENG')

        try:
            # Bug fix: os.path.dirname(config_path) is "" when config_path
            # is a bare relative filename (e.g. launched with a relative
            # --config path), and os.makedirs("", exist_ok=True) raises
            # FileNotFoundError - which used to be caught below and logged,
            # but the file itself then never got written, silently. Only
            # call makedirs when there's actually a directory component.
            config_dir = os.path.dirname(config_path)
            if config_dir:
                os.makedirs(config_dir, exist_ok=True)
            with open(config_path, 'w') as configfile:
                self.config.write(configfile)
        except OSError as e:
            self.logger.error(f"Failed to save settings to {config_path}: {e}")
            messagebox.showerror(TRANSLATIONS[self.current_language]['error'], str(e))
            return

        # So the next launch reopens this same config file automatically.
        self._remember_last_config_path(config_path)

        self.mark_clean()


# ============================================================
# Main execution
# ============================================================
def _build_bilingual_help(exe_name, is_cli_build):
    """
    Full bilingual (RU + EN) CLI help text. argparse's own --help
    formatting is English-only and can't easily be made bilingual, so
    this is built by hand and used instead.
    """
    if is_cli_build:
        ru_behavior = (
            "Это консольная сборка: запуск без аргументов просто показывает эту справку.\n"
            "Чтобы применить раскладку, укажите --config — конфиг загрузится, окна\n"
            "расставятся и программа закроется; окно программы не показывается.\n"
        )
        en_behavior = (
            "This is the console build: running it with no arguments just shows\n"
            "this help. Pass --config to actually apply a layout — the config loads,\n"
            "windows are arranged, and the program exits; no window is ever shown.\n"
        )
    else:
        ru_behavior = (
            "Это обычная сборка: окно программы открывается как всегда.\n"
            "--config лишь подгружает указанный файл вместо последнего использованного.\n"
        )
        en_behavior = (
            "This is the regular build: the app window always opens as usual.\n"
            "--config just preloads the given file instead of the last one used.\n"
        )

    bare_run_ru = (
        "      Показать эту справку (для запуска нужен хотя бы один аргумент).\n"
        if is_cli_build else
        "      Открыть программу с последним использованным конфигом.\n"
    )
    bare_run_en = (
        "      Shows this help (running it for real needs at least one argument).\n"
        if is_cli_build else
        "      Open the app with the last-used config.\n"
    )

    return (
        "=" * 60 + "\n"
        "РУССКИЙ\n" + "=" * 60 + "\n"
        f"{APP_NAME}.\n"
        "Расставляет окна по сетке на выбранных мониторах по их заголовкам.\n"
        "\n"
        + ru_behavior +
        "\n"
        "Использование:\n"
        f"  {exe_name} [-h] [--config ПУТЬ] [--list-monitors]\n"
        "\n"
        "Параметры:\n"
        "  -h, --help        показать эту справку и выйти\n"
        "  --config ПУТЬ     использовать указанный .ini файл вместо последнего использованного\n"
        "  --list-monitors   вывести список обнаруженных мониторов и выйти — окно не открывается\n"
        "\n"
        "Примеры:\n"
        f"  {exe_name}\n"
        + bare_run_ru +
        "\n"
        f"  {exe_name} --config \"C:\\configs\\work.ini\"\n"
        "      Использовать конкретный конфиг вместо последнего использованного.\n"
        "\n"
        f"  {exe_name} --list-monitors\n"
        "      Вывести список мониторов (разрешение, позиция, масштаб DPI) и выйти.\n"
        "\n\n"
        + "=" * 60 + "\n"
        "ENGLISH\n" + "=" * 60 + "\n"
        f"{APP_NAME}.\n"
        "Arranges windows in a grid across the selected monitors by title.\n"
        "\n"
        + en_behavior +
        "\n"
        "Usage:\n"
        f"  {exe_name} [-h] [--config PATH] [--list-monitors]\n"
        "\n"
        "Options:\n"
        "  -h, --help        show this help message and exit\n"
        "  --config PATH     Use this .ini config file instead of the last one used\n"
        "  --list-monitors   Print detected monitors to the console and exit — no window is opened\n"
        "\n"
        "Examples:\n"
        f"  {exe_name}\n"
        + bare_run_en +
        "\n"
        f"  {exe_name} --config \"C:\\configs\\work.ini\"\n"
        "      Use a specific config instead of the last-used one.\n"
        "\n"
        f"  {exe_name} --list-monitors\n"
        "      Print detected monitors (resolution, position, DPI scale) and exit.\n"
    )


def _print_monitor_list(logger):
    """--list-monitors: pure console diagnostic, no GUI involved at all."""
    info = MonitorInfo(logger=logger)
    if not info.monitors:
        print("No monitors detected.")
        return
    for i, m in enumerate(info.monitors):
        primary = " [Primary]" if m['is_primary'] else ""
        print(f"Monitor {i + 1}{primary}: {m['physical_width']}x{m['physical_height']} "
              f"at ({m['left']},{m['top']}), scale {m['scale']}%")


def _ensure_console_output():
    """
    A windowed/GUI-subsystem build (the common way to package a Tkinter
    app) has no console at all: sys.stdout/sys.stderr are None, so
    print() either raises or is silently swallowed — that's why --help,
    --list-monitors, etc. produced no output and cmd returned instantly.

    If we already have a working stdout (e.g. running the .py directly,
    or the exe was built with a console), do nothing. Otherwise, try to
    attach to whichever console launched us (cmd.exe) and redirect
    stdout/stderr/stdin to it, so CLI output actually reaches the user.
    """
    try:
        if sys.stdout is not None:
            sys.stdout.fileno()
            return  # already have a real, usable stream
    except Exception:
        pass

    try:
        import ctypes
        ATTACH_PARENT_PROCESS = -1
        if ctypes.windll.kernel32.AttachConsole(ATTACH_PARENT_PROCESS):
            sys.stdout = open('CONOUT$', 'w')
            sys.stderr = open('CONOUT$', 'w')
            sys.stdin = open('CONIN$', 'r')
    except Exception:
        # No parent console to attach to (e.g. launched by double-click) —
        # leave stdout/stderr as-is; there's no CLI user watching anyway.
        pass


if __name__ == "__main__":
    _ensure_console_output()
    import argparse

    exe_name = os.path.basename(sys.argv[0])
    # Behavior is decided by which exe you run, not by flags: a build
    # whose filename contains "CLI" always runs headless (load config,
    # arrange, exit — no window); the regular build always opens the
    # window as usual. --config just picks which .ini file to use either way.
    is_cli_build = 'cli' in exe_name.lower()

    # -h/--help is handled by hand (add_help=False below) so the help
    # text can be fully bilingual (RU + EN) — argparse's own --help
    # formatting only supports one language at a time.
    if '-h' in sys.argv[1:] or '--help' in sys.argv[1:]:
        print(_build_bilingual_help(exe_name, is_cli_build))
        sys.exit(0)

    # The CLI build launched with no arguments at all (e.g. a stray
    # double-click) shows help instead of silently applying the
    # last-used config — running it unattended requires an explicit
    # argument (at minimum --config). Waits for Enter so the console
    # doesn't vanish instantly if this was a double-click.
    if is_cli_build and len(sys.argv) == 1:
        print(_build_bilingual_help(exe_name, is_cli_build))
        try:
            input("Press Enter to exit / Нажмите Enter для выхода...")
        except (EOFError, OSError):
            pass
        sys.exit(0)

    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--config', metavar='PATH')
    parser.add_argument('--list-monitors', action='store_true')
    cli_args = parser.parse_args()

    if getattr(sys, 'frozen', False):
        _app_path = os.path.dirname(sys.executable)
    else:
        _app_path = os.path.dirname(os.path.abspath(__file__))

    log = _init_logging(_app_path)

    if cli_args.list_monitors:
        try:
            _print_monitor_list(log)
        except Exception as e:
            log.exception("Failed to list monitors")
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
        sys.exit(0)

    if cli_args.config and not os.path.exists(cli_args.config):
        print(f"Config file not found: {cli_args.config}", file=sys.stderr)
        sys.exit(2)

    try:
        root = tk.Tk()
        root.attributes('-topmost', True)
        if is_cli_build:
            # Never show or flash the main window for the CLI build.
            root.withdraw()
        root.resizable(True, True)  # GUI improvement #3: allow horizontal resize too
        # GUI improvement #1: topmost is no longer forced; the app behaves
        # like a normal window and won't block other work.
        app = WindowManager(root, log,
                             override_config_path=cli_args.config,
                             headless=is_cli_build,
                             force_apply=is_cli_build)
        root.mainloop()
    except Exception as e:
        log.exception("Unhandled error")
        if is_cli_build:
            print(f"Error: {e}", file=sys.stderr)
        else:
            messagebox.showerror(TRANSLATIONS['ru']['error'],
                                  f"{TRANSLATIONS['ru']['error']}: {str(e)}")
        sys.exit(1)
