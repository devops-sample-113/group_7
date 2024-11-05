from tkinter import *
import tkinter as tk
import openpyxl
import subprocess
from tkinter import messagebox

def search(id):
    for row in worksheet_students.iter_rows(min_roe=2, values_only=True):
        name, id_, schedule_path = row[:3]
        if id_ == id:
            messagebox.showinfo("成功", "已找到該學生")
            break
        else:
            messagebox.showinfo("錯誤", "未找到該學生")
            break