import tkinter as tk
from tkinter import scrolledtext, ttk
import threading
from comcenter_parser import (setup_session, get_laser_printers_database, process_xls_database,
                             parse_printer_compatibility, filter_compatibility_by_stock,
                             parse_cartridges_and_parts, parse_all_cartridges_and_parts,
                             parse_comcenter_products, run_action, CancelFlag)
import datetime

class Tooltip:
    """Класс для создания всплывающих подсказок"""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)
        self.tooltip_window.wm_geometry(f"+{x}+{y}")
        label = tk.Label(self.tooltip_window, text=self.text, background="lightyellow",
                         relief="solid", borderwidth=1, font=("Arial", 10))
        label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

class GUIOutputHandler:
    """Обработчик вывода для GUI с записью в файл и прогресс-баром"""
    def __init__(self, text_widget, progress_bar):
        self.text_widget = text_widget
        self.progress_bar = progress_bar
        self.log_file = "comcenter_parser.log"

    def log(self, message):
        self.text_widget.insert(tk.END, message + "\n")
        self.text_widget.see(tk.END)
        self.text_widget.update()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(self.log_file, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")

    def progress(self, current, total):
        """Обновление прогресс-бара"""
        percentage = (current / total) * 100
        self.progress_bar['value'] = percentage
        self.text_widget.insert(tk.END, f"Прогресс: {current}/{total} ({percentage:.1f}%)\n")
        self.text_widget.see(tk.END)
        self.text_widget.update()

class ComcenterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Comcenter Parser")
        self.root.geometry("600x400")

        # Создаем текстовое поле для логов
        self.log_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, height=15, width=60)
        self.log_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

        # Создаем прогресс-бар
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=500, mode="determinate")
        self.progress_bar.pack(padx=10, pady=5)

        # Создаем обработчик вывода
        self.output_handler = GUIOutputHandler(self.log_area, self.progress_bar)

        # Создаем фрейм для кнопок
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(padx=10, pady=10)

        # Кнопки для действий с подсказками
        self.action_buttons = []
        self.buttons = [
            ("ПАРСИНГ СОВМЕСТИМОСТИ", self.run_action_3_4, "Парсит совместимость всех картриджей и запчастей, затем фильтрует по наличию"),
            ("ПАРСИНГ КАРТРИДЖЕЙ И ЗАПЧАСТЕЙ В НАЛИЧИИ", self.run_action_5, "Парсит данные актуальных картриджей и запчастей"),
            ("ПОЛНЫЙ ПАРСИНГ КАРТРИДЖЕЙ И ЗАПЧАСТЕЙ", self.run_action_6, "Парсит данные всех картриджей и запчастей"),
            ("ПАРСИНГ ПРАЙСА ТОВАРОВ", self.run_action_7, "Парсит данные всех актуальных товаров Comcenter"),
            ("Выход", self.exit, "")
        ]

        # Добавляем кнопки в интерфейс
        for text, command, tooltip in self.buttons:
            btn = tk.Button(self.button_frame, text=text, command=command, width=50)  # Увеличена ширина кнопок
            btn.pack(pady=5)
            if tooltip:
                Tooltip(btn, tooltip)
            if text != "Выход":
                self.action_buttons.append(btn)

        # Кнопка "Отмена"
        self.cancel_button = tk.Button(self.button_frame, text="Отмена", command=self.cancel, width=50, state=tk.DISABLED)
        self.cancel_button.pack(pady=5)
        Tooltip(self.cancel_button, "Прерывает выполнение текущей операции")

        # Флаг отмены
        self.cancel_flag = None

        # Инициализация сессии и запуск начальных действий
        self.session_info = None
        self.setup_session()
        if self.session_info:
            self.run_initial_actions()

    def setup_session(self):
        """Инициализация сессии"""
        self.session_info = setup_session(self.output_handler)
        if not self.session_info:
            self.output_handler.log("Не удалось авторизоваться. Некоторые функции будут недоступны.")

    def enable_buttons(self, enable=True):
        """Активация/деактивация кнопок меню"""
        state = tk.NORMAL if enable else tk.DISABLED
        for btn in self.action_buttons:
            btn.config(state=state)
        self.cancel_button.config(state=tk.DISABLED if enable else tk.NORMAL)

    def reset_progress(self):
        """Сброс прогресс-бара"""
        self.progress_bar['value'] = 0

    def run_in_thread(self, action):
        """Запуск действия в отдельном потоке"""
        self.cancel_flag = CancelFlag()
        self.enable_buttons(False)
        self.reset_progress()

        def wrapper():
            try:
                action()
            finally:
                self.enable_buttons(True)
                self.cancel_flag = None
                self.reset_progress()

        threading.Thread(target=wrapper, daemon=True).start()

    def cancel(self):
        """Отмена текущей операции"""
        if self.cancel_flag:
            self.cancel_flag.cancel()
            self.output_handler.log("Запрос на отмену операции отправлен...")

    def run_initial_actions(self):
        """Автоматический запуск действий 1 и 2 при старте"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.output_handler.log("Загрузка...")
        self.run_in_thread(lambda: self.initial_actions_wrapper(session, headers))

    def initial_actions_wrapper(self, session, headers):
        """Обертка для последовательного выполнения действий 1 и 2"""
        try:
            get_laser_printers_database(session, headers, self.output_handler, CancelFlag())
            if self.cancel_flag and self.cancel_flag.is_cancelled():
                self.output_handler.log("Операция отменена")
                return
            process_xls_database(session, headers, self.output_handler, CancelFlag())
        except Exception as e:
            self.output_handler.log(f"Ошибка при выполнении начальных действий: {e}")

    def run_action_3_4(self):
        """Действие 3 и 4: Парсинг совместимости и фильтрация по наличию"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: self.action_3_4_wrapper(session, headers))

    def action_3_4_wrapper(self, session, headers):
        """Обертка для последовательного выполнения действий 3 и 4"""
        try:
            parse_printer_compatibility(session, headers, self.output_handler, CancelFlag())
            if self.cancel_flag and self.cancel_flag.is_cancelled():
                self.output_handler.log("Операция отменена")
                return
            filter_compatibility_by_stock(self.output_handler, CancelFlag())
        except Exception as e:
            self.output_handler.log(f"Ошибка при парсинге совместимости: {e}")

    def run_action_5(self):
        """Действие 5: Парсинг актуальных картриджей и запчастей"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_cartridges_and_parts(session, headers, self.output_handler, self.cancel_flag))

    def run_action_6(self):
        """Действие 6: Парсинг всех картриджей и запчастей"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_all_cartridges_and_parts(session, headers, self.output_handler, self.cancel_flag))

    def run_action_7(self):
        """Действие 7: Парсинг актуальных товаров Comcenter"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_comcenter_products(session, headers, self.output_handler, self.cancel_flag))

    def exit(self):
        """Выход из приложения"""
        self.output_handler.log("Программа завершена")
        self.root.quit()

def main():
    root = tk.Tk()
    app = ComcenterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()