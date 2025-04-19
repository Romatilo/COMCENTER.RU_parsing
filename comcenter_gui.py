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
            ("ID лазерных принтеров", self.run_action_1, "Сохраняет ID всех лазерных принтеров"),
            ("ID товаров из прайс-листа", self.run_action_2, "Сохраняет ID всех актуальных товаров из xls-прайс-листа"),
            ("Совместимость ВСЕХ", self.run_action_3, "Парсит совместимость всех картриджей и запчастей для принтеров"),
            ("Совместимость АКТУАЛЬНЫХ", self.run_action_4, "Парсит совместимость актуальных картриджей и запчастей для принтеров"),
            ("АКТУАЛЬНЫЕ Картриджи и запчасти", self.run_action_5, "Парсит данные АКТУАЛЬНЫХ картриджей и запчастей"),
            ("ВСЕ картриджи и запчасти", self.run_action_6, "Парсит данные ВСЕХ картриджей и запчастей"),
            ("ВСЕ АКТУАЛЬНЫЕ товары", self.run_action_7, "Парсит данные ВСЕХ АКТУАЛЬНЫХ товаров Comcenter"),
            ("Выход", self.exit, "")
        ]

        # Добавляем кнопки в интерфейс
        for text, command, tooltip in self.buttons:
            btn = tk.Button(self.button_frame, text=text, command=command, width=40)
            btn.pack(pady=5)
            if tooltip:
                Tooltip(btn, tooltip)
            if text != "Выход":
                self.action_buttons.append(btn)

        # Кнопка "Отмена"
        self.cancel_button = tk.Button(self.button_frame, text="Отмена", command=self.cancel, width=40, state=tk.DISABLED)
        self.cancel_button.pack(pady=5)
        Tooltip(self.cancel_button, "Прерывает выполнение текущей операции")

        # Флаг отмены
        self.cancel_flag = None

        # Инициализация сессии
        self.session_info = None
        self.setup_session()

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

    def run_in_thread(self, action, choice):
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

    def run_action_1(self):
        """Действие 1: Получение базы данных лазерных принтеров"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: get_laser_printers_database(session, headers, self.output_handler, self.cancel_flag), "1")

    def run_action_2(self):
        """Действие 2: Получение базы данных из XLS"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: process_xls_database(session, headers, self.output_handler, self.cancel_flag), "2")

    def run_action_3(self):
        """Действие 3: Парсинг совместимости принтеров"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_printer_compatibility(session, headers, self.output_handler, self.cancel_flag), "3")

    def run_action_4(self):
        """Действие 4: Фильтрация совместимости"""
        self.run_in_thread(lambda: filter_compatibility_by_stock(self.output_handler, self.cancel_flag), "4")

    def run_action_5(self):
        """Действие 5: Парсинг актуальных картриджей и запчастей"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_cartridges_and_parts(session, headers, self.output_handler, self.cancel_flag), "5")

    def run_action_6(self):
        """Действие 6: Парсинг всех картриджей и запчастей"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_all_cartridges_and_parts(session, headers, self.output_handler, self.cancel_flag), "6")

    def run_action_7(self):
        """Действие 7: Парсинг актуальных товаров Comcenter"""
        if not self.session_info:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_comcenter_products(session, headers, self.output_handler, self.cancel_flag), "7")

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