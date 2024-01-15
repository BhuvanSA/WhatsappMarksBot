import tkinter as tk
from ttkbootstrap import ttk

class RangeSlider(ttk.Frame):
    def __init__(self, parent, min_val=0, max_val=20, **kwargs):
        super().__init__(parent, **kwargs)
        self.min_val = min_val
        self.max_val = max_val
        self.range = self.max_val - self.min_val

        self.canvas = tk.Canvas(self, width=200, height=50, bg='white')
        self.canvas.pack()

        self.left_handle = self.create_handle(self.min_val)
        self.right_handle = self.create_handle(self.max_val)

        self.canvas.tag_bind(self.left_handle, "<B1-Motion>", self.on_drag_left)
        self.canvas.tag_bind(self.right_handle, "<B1-Motion>", self.on_drag_right)

    def create_handle(self, val):
        x = (val - self.min_val) / self.range * 200
        return self.canvas.create_rectangle(x, 0, x + 10, 50, fill='blue')

    def on_drag_left(self, event):
        self.canvas.coords(self.left_handle, event.x, 0, event.x + 10, 50)
        self.update_values()

    def on_drag_right(self, event):
        self.canvas.coords(self.right_handle, event.x, 0, event.x + 10, 50)
        self.update_values()

    def update_values(self):
        left_x = self.canvas.coords(self.left_handle)[0]
        right_x = self.canvas.coords(self.right_handle)[0]
        self.min_val = left_x / 200 * self.range
        self.max_val = right_x / 200 * self.range
        print(f"Range: {self.min_val} - {self.max_val}")