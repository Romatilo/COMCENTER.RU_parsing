import tkinter as tk
from tkinter import scrolledtext
import threading
from comcenter_parser import (setup_session, get_laser_printers_database, process_xls_database,
                             parse_printer_compatibility, filter_compatibility_by_stock,
                             parse_cartridges_and_parts, parse_all_cartridges_and_parts,
                             parse_comcenter_products)

class GUIOutputHandler:
    """Обработчик вывода для GUI"""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def log(self, message):
        self.text_widget.insert(tk.END, message + "\n")
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

        # Создаем обработчик вывода
        self.output_handler = GUIOutputHandler(self.log_area)

        # Создаем фрейм для кнопок
        self.button_frame = tk.Frame(self.root)
        self.button_frame.pack(padx=10, pady=10)

        # Кнопки для действий
        self.buttons = [
            ("1. Лазерные принтеры", self.run_action_1),
            ("2. XLS база данных", self.run_action_2),
            ("3. Совместимость принтеров", self.run_action_3),
            ("4. Фильтр совместимости", self.run_action_4),
            ("5. Актуальные картриджи/запчасти", self.run_action_5),
            ("6. Все картриджи/запчасти", self.run_action_6),
            ("7. Актуальные товары", self.run_action_7),
            ("Выход", self.exit)
        ]

        # Добавляем кнопки в интерфейс
        for text, command in self.buttons:
            tk.Button(self.button_frame, text=text, command=command, width=25).pack(pady=5)

        # Инициализация сессии
        self.session_info = None
        self.setup_session()

    def setup_session(self):
        """Инициализация сессии"""
        self.session_info = setup_session(self.output_handler)
        if not self.session_info:
            self.output_handler.log("Не удалось авторизоваться. Некоторые функции будут недоступны.")

    def run_in_thread(self, action):
        """Запуск действия в отдельном потоке"""
        if not self.session_info and action not in [self.run_action_4, self.exit]:
            self.output_handler.log("Сессия не инициализирована. Пожалуйста, перезапустите приложение.")
            return
        threading.Thread(target=action, daemon=True).start()

    def run_action_1(self):
        """Действие 1: Получение базы данных лазерных принтеров"""
        session, headers = self.session_info
        self.run_in_thread(lambda: get_laser_printers_database(session, headers, self.output_handler))

    def run_action_2(self):
        """Действие 2: Получение базы данных из XLS"""
        session, headers = self.session_info
        self.run_in_thread(lambda: process_xls_database(session, headers, self.output_handler))

    def run_action_3(self):
        """Действие 3: Парсинг совместимости принтеров"""
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_printer_compatibility(session, headers, self.output_handler))

    def run_action_4(self):
        """Действие 4: Фильтрация совместимости"""
        self.run_in_thread(lambda: filter_compatibility_by_stock(self.output_handler))

    def run_action_5(self):
        """Действие 5: Парсинг актуальных картриджей и запчастей"""
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_cartridges_and_parts(session, headers, self.output_handler))

    def run_action_6(self):
        """Действие 6: Парсинг всех картриджей и запчастей"""
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_all_cartridges_and_parts(session, headers, self.output_handler))

    def run_action_7(self):
        """Действие 7: Парсинг актуальных товаров Comcenter"""
        session, headers = self.session_info
        self.run_in_thread(lambda: parse_comcenter_products(session, headers, self.output_handler))

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