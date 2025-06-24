import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import subprocess
import threading
import time
import os
import sys

class RobotExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Robot Framework Excel Setup")
        self.root.geometry("1020x600")
        self.root.configure(bg="#f5f5f8")
        self.root.resizable(False, False)

        self.ROBOT_FILE = self.resource_path("test.robot")
        self.COLOR_BG = "#f5f5f8"
        self.COLOR_MAIN = "#7A58BF"
        self.COLOR_ACCENT = "#F2B234"
        self.COLOR_BTN_TEXT = "#ffffff"
        self.COLOR_TEXT = "#333333"
        self.FONT_REGULAR = ("Segoe UI", 13)
        self.FONT_HEADER = ("Segoe UI", 22, "bold")
        self.sleep_prevention_active = False

        self.setup_style()
        self.load_icons()
        self.create_widgets()

        try:
            self.root.call('tk', 'scaling', 1.5)
        except:
            pass

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def setup_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=self.COLOR_BG)
        style.configure("TLabel", background=self.COLOR_BG, foreground=self.COLOR_TEXT, font=self.FONT_REGULAR)
        style.configure("Header.TLabel", font=self.FONT_HEADER, foreground=self.COLOR_MAIN, background=self.COLOR_BG)
        style.configure("TEntry", font=self.FONT_REGULAR, padding=7)
        style.configure("TCombobox", font=self.FONT_REGULAR, padding=7)

    def load_icons(self):
        try:
            self.robot_icon = ImageTk.PhotoImage(Image.open(self.resource_path("robot_icon.png")).resize((36, 36), Image.LANCZOS))
        except:
            self.robot_icon = None
        try:
            self.excel_icon = ImageTk.PhotoImage(Image.open(self.resource_path("excel_icon.png")).resize((26, 26), Image.LANCZOS))
        except:
            self.excel_icon = None
        try:
            self.sheet_icon = ImageTk.PhotoImage(Image.open(self.resource_path("sheet_icon.png")).resize((26, 26), Image.LANCZOS))
        except:
            self.sheet_icon = None

    def create_widgets(self):
        header_frame = ttk.Frame(self.root, style="TFrame")
        header_frame.place(relx=0.5, y=40, anchor="center")
        if self.robot_icon:
            ttk.Label(header_frame, image=self.robot_icon, background=self.COLOR_BG).pack(side="left", padx=(0, 10))
        ttk.Label(header_frame, text="Robot Framework Excel Setup", style="Header.TLabel").pack(side="left")

        outer_frame = ttk.Frame(self.root, padding=30)
        outer_frame.place(relx=0.5, rely=0.42, anchor="center")

        if self.excel_icon:
            ttk.Label(outer_frame, image=self.excel_icon, background=self.COLOR_BG).grid(row=0, column=0, sticky="e", padx=(0, 5), pady=15)
        ttk.Label(outer_frame, text="Excel File:").grid(row=0, column=1, sticky="w", padx=(0, 5), pady=15)
        self.entry_filepath = ttk.Entry(outer_frame, width=50)
        self.entry_filepath.grid(row=0, column=2, padx=5, pady=15)
        ttk.Button(outer_frame, text="Browse", command=self.browse_file, width=12).grid(row=0, column=3, padx=5, pady=15)

        if self.sheet_icon:
            ttk.Label(outer_frame, image=self.sheet_icon, background=self.COLOR_BG).grid(row=1, column=0, sticky="e", padx=(0, 5), pady=15)
        ttk.Label(outer_frame, text="Sheet Name:").grid(row=1, column=1, sticky="w", padx=5, pady=15)
        self.sheet_combobox = ttk.Combobox(outer_frame, width=48, state="readonly")
        self.sheet_combobox.grid(row=1, column=2, columnspan=2, padx=5, pady=15, sticky="w")

        btn_canvas = tk.Canvas(self.root, width=800, height=100, bg=self.COLOR_BG, highlightthickness=0)
        btn_canvas.place(relx=0.5, rely=0.78, anchor="center")
        self.create_rounded_button(btn_canvas, 90, 20, 260, 50, 25, "Update Robot File", self.run_update)
        self.create_rounded_button(btn_canvas, 370, 20, 260, 50, 25, "Run Robot Test", self.run_robot_test_with_nosleep)

    def create_rounded_button(self, canvas, x, y, width, height, radius, text, command,
                              fill="#7A58BF", text_color="#ffffff", hover_fill="#F2B234"):
        points = [
            x + radius, y,
            x + width - radius, y,
            x + width, y,
            x + width, y + radius,
            x + width, y + height - radius,
            x + width, y + height,
            x + width - radius, y + height,
            x + radius, y + height,
            x, y + height,
            x, y + height - radius,
            x, y + radius,
            x, y
        ]
        button = canvas.create_polygon(points, smooth=True, fill=fill, outline=fill)
        label = canvas.create_text(x + width / 2, y + height / 2, text=text, fill=text_color, font=("Segoe UI", 12, "bold"))

        def on_enter(e): canvas.itemconfig(button, fill=hover_fill, outline=hover_fill)
        def on_leave(e): canvas.itemconfig(button, fill=fill, outline=fill)
        def on_click(e): command()

        for tag in [button, label]:
            canvas.tag_bind(tag, "<Enter>", on_enter)
            canvas.tag_bind(tag, "<Leave>", on_leave)
            canvas.tag_bind(tag, "<Button-1>", on_click)

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")], title="Select Excel File")
        if filepath:
            self.entry_filepath.delete(0, tk.END)
            self.entry_filepath.insert(0, filepath)
            self.populate_sheetnames(filepath)

    def populate_sheetnames(self, filepath):
        try:
            sheets = pd.ExcelFile(filepath).sheet_names
            self.sheet_combobox['values'] = sheets
            if sheets:
                self.sheet_combobox.current(0)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file:\n{e}")

    def run_update(self):
        filepath = self.entry_filepath.get().strip()
        sheetname = self.sheet_combobox.get().strip()
        if not filepath:
            messagebox.showwarning("Input Error", "Please select an Excel file.")
            return False
        if not sheetname:
            messagebox.showwarning("Input Error", "Please select a sheet name.")
            return False
        return self.update_robot_library(filepath, sheetname)

    def update_robot_library(self, filepath, sheetname):
        try:
            with open(self.ROBOT_FILE, "r") as file:
                lines = file.readlines()
            updated = False
            for i, line in enumerate(lines):
                if line.strip().startswith("Library") and "ExcelHandler.py" in line:
                    lines[i] = f"Library    ExcelHandler.py    {filepath}    {sheetname}\n"
                    updated = True
                    break
            if not updated:
                lines.insert(0, f"Library    ExcelHandler.py    {filepath}    {sheetname}\n")
            with open(self.ROBOT_FILE, "w") as file:
                file.writelines(lines)
            messagebox.showinfo("Success", f"Updated {self.ROBOT_FILE} successfully!")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update robot file:\n{e}")
            return False

    def prevent_sleep(self):
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001 | 0x00000002)

    def allow_sleep(self):
        ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)

    def keep_awake(self):
        while self.sleep_prevention_active:
            self.prevent_sleep()
            time.sleep(30)
        self.allow_sleep()

    def run_robot_test_with_nosleep(self):
        if not self.run_update():
            return
        self.sleep_prevention_active = True
        threading.Thread(target=self.keep_awake, daemon=True).start()
        try:
            result = subprocess.run(["robot", self.ROBOT_FILE], capture_output=True, text=True)
            if result.returncode == 0:
                messagebox.showinfo("Robot Test Result", "Robot test finished successfully!")
            else:
                messagebox.showerror("Robot Test Error", f"Robot test failed:\n{result.stderr}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to run Robot Framework test:\n{e}")
        finally:
            self.sleep_prevention_active = False


if __name__ == "__main__":
    root = tk.Tk()
    app = RobotExcelApp(root)
    root.mainloop()
