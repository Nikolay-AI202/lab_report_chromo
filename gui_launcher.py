# gui_launcher.py

import tkinter as tk
import sys
from tkinter import messagebox
from main_wrapped import process_documents  # импортируем обёрнутую функцию

def run_script():
    try:
        process_documents()
        messagebox.showinfo("Успех", "Обработка завершена успешно.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")

def exit_program():
    root.destroy()
    sys.exit()

root = tk.Tk()
root.title("Импорт Word → Excel")
root.geometry("300x200")

# Кнопка запуска обработки
btn = tk.Button(root, text="Запустить обработку", command=run_script, height=2, width=30)
btn.pack(pady=40)

# Кнопка выхода
btn_exit = tk.Button(root, text="Выход", command=exit_program, height=1, width=30)
btn_exit.pack()

root.mainloop()