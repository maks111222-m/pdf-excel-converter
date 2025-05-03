import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from pdf2image import convert_from_path
from PIL import Image, ImageTk
import pdfplumber
import pandas as pd
from openpyxl import Workbook
from pandastable import Table
import os

class PDFViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Viewer Named Save Final")
        self.geometry("1200x800")
        self.after(100, self.focus_force)

        self.canvas = tk.Canvas(self, bg="black", cursor="cross")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.page_images = []
        self.tk_img = None
        self.pdf_pages = []
        self.pdf_path = None
        self.page_index = 0
        self.scale = 1.0

        self.selection = None
        self.selecting = False
        self.start_x = self.start_y = 0
        self.help_visible = False

        self.manual_rows, self.auto_rows1, self.auto_rows2 = [], [], []
        self.excel_filename = None
        self.excel_filename = None

        self.bind_keys()
        self.bind_mouse()
        self.load_pdf()

    def bind_keys(self):
        self.bind_all("<Control-MouseWheel>", self.on_zoom)
        self.bind_all("<MouseWheel>", self.on_scroll)
        self.bind_all("<Key>", self.on_key)
        self.bind_all("<Escape>", self.exit_and_save)

    def on_key(self, event):
        keycode_map = {
            72: "h",    # H
            73: "i",    # I
            84: "t",    # T
            37: "left",
            39: "right",
            27: "escape"
        }

        key = keycode_map.get(event.keycode, "").lower()

        if key == "left":
            self.prev_page()
        elif key == "right":
            self.next_page()
        elif key == "escape":
            self.exit_and_save()
        elif key == "h":
            self.toggle_help()
        elif key == "i":
            self.import_table()
        elif key == "t":
            self.import_text()

    def bind_mouse(self):
        self.canvas.bind("<Button-1>", self.start_select)
        self.canvas.bind("<B1-Motion>", self.update_select)
        self.canvas.bind("<ButtonRelease-1>", self.end_select)

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        self.pdf_path = path
        if not path:
            self.destroy()
            return
        self.page_images = convert_from_path(path, dpi=150)
        self.pdf_pages = pdfplumber.open(path).pages
        self.render_page()

    def render_page(self):
        self.canvas.delete("all")
        img = self.page_images[self.page_index]
        w, h = img.size
        img = img.resize((int(w * self.scale), int(h * self.scale)))
        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas.create_image(0, 0, anchor="nw", image=self.tk_img)

        if self.selection:
            self.canvas.create_rectangle(*self.selection, outline="lime", width=2)

        label = f"Page {self.page_index+1}/{len(self.page_images)} | Scale: {self.scale:.2f}"
        self.canvas.create_text(10, 10, anchor="nw", fill="white", font=("Arial", 12), text=label)

        if self.help_visible:
            help_text = "H: подсказка  |  I: импорт таблицы  |  T: импорт текста"
            font = ("Arial", 12, "bold")

            # Отрисуем временно текст, чтобы получить bbox
            temp_id = self.canvas.create_text(10, 30, anchor="nw", text=help_text, font=font)
            bbox = self.canvas.bbox(temp_id)
            self.canvas.delete(temp_id)

            if bbox:
                x0, y0, x1, y1 = bbox
                self.canvas.create_rectangle(
                    x0 - 6, y0 - 4, x1 + 6, y1 + 4,
                    fill="black", outline="", tag="help"
                )
                self.canvas.create_text(10, 30, anchor="nw", fill="red", font=font, text=help_text, tag="help")

    def on_scroll(self, event):
        if event.delta > 0:
            self.prev_page()
        else:
            self.next_page()

    def on_zoom(self, event):
        self.scale *= 1.1 if event.delta > 0 else 0.9
        self.scale = max(0.3, min(self.scale, 5.0))
        self.render_page()

    def prev_page(self, event=None):
        if self.page_index > 0:
            self.page_index -= 1
            self.selection = None
            self.render_page()

    def next_page(self, event=None):
        if self.page_index < len(self.page_images) - 1:
            self.page_index += 1
            self.selection = None
            self.render_page()

    def start_select(self, event):
        self.selecting = True
        self.start_x, self.start_y = event.x, event.y
        self.selection = None

    def update_select(self, event):
        if self.selecting:
            x0, y0 = self.start_x, self.start_y
            x1, y1 = event.x, event.y
            self.selection = (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
            self.render_page()

    def end_select(self, event):
        self.selecting = False

    def toggle_help(self):
        self.help_visible = not self.help_visible
        if not self.help_visible:
            self.canvas.delete("help")
        self.render_page()

    def import_text(self):
        if not self.selection:
            messagebox.showinfo("Импорт текста", "Сначала выделите область.")
            return
        try:
            page = self.pdf_pages[self.page_index]
            w_img, h_img = self.page_images[self.page_index].size
            x0, y0, x1, y1 = self.selection
            bbox = (
                x0 / self.scale * page.width / w_img,
                y0 / self.scale * page.height / h_img,
                x1 / self.scale * page.width / w_img,
                y1 / self.scale * page.height / h_img
            )
            cropped = page.crop(bbox)
            text = cropped.extract_text()
            if text:
                lines = [line.split() for line in text.strip().split("\n")]
                for i in range(1, len(lines)):
                    lines[i].insert(0, "")
                self.manual_rows.extend(lines)           
            else:
                messagebox.showinfo("Пусто", "Не удалось извлечь текст.")
        except Exception as e:
            messagebox.showerror("Ошибка текста", str(e))

    def import_table(self):
        if not self.selection:
            messagebox.showinfo("Импорт таблицы", "Сначала выделите область.")
            return
        try:
            page = self.pdf_pages[self.page_index]
            w_img, h_img = self.page_images[self.page_index].size
            x0, y0, x1, y1 = self.selection
            bbox = (
                x0 / self.scale * page.width / w_img,
                y0 / self.scale * page.height / h_img,
                x1 / self.scale * page.width / w_img,
                y1 / self.scale * page.height / h_img
            )
            cropped = page.crop(bbox)
            tables = cropped.extract_tables()
            if tables:
                for tbl in tables:
                    df = pd.DataFrame(tbl).fillna("")
                    df = self.edit_table(df)
                    target = simpledialog.askstring("Лист", "Auto1 или Auto2?", initialvalue="Auto1")
                    if target and target.lower().startswith("auto2"):
                        self.auto_rows2.extend(df.values.tolist())
                    else:
                        self.auto_rows1.extend(df.values.tolist())
            else:
                messagebox.showinfo("Пусто", "Не найдено таблиц.")
        except Exception as e:
            messagebox.showerror("Ошибка таблицы", str(e))

    def edit_table(self, df):
        top = tk.Toplevel(self)
        top.title("Редактирование таблицы")
        frame = tk.Frame(top)
        frame.pack(fill="both", expand=True)
        table = Table(frame, dataframe=df, showtoolbar=True, showstatusbar=True)
        table.show()

        def on_close():
            top.destroy()

        top.protocol("WM_DELETE_WINDOW", on_close)
        self.wait_window(top)
        return table.model.df.copy()

    def save_excel(self):
        base = os.path.splitext(os.path.basename(self.pdf_path))[0]
        filename = base + ".xlsx"
        counter = 1
        while os.path.exists(filename):
            filename = f"{base}_{counter}.xlsx"
            counter += 1
        with pd.ExcelWriter(filename, engine="openpyxl", mode="w") as writer:
            if self.manual_rows:
                pd.DataFrame(self.manual_rows).to_excel(writer, sheet_name="Manual", index=False, header=False)
            if self.auto_rows1:
                pd.DataFrame(self.auto_rows1).to_excel(writer, sheet_name="Auto1", index=False, header=False)
            if self.auto_rows2:
                pd.DataFrame(self.auto_rows2).to_excel(writer, sheet_name="Auto2", index=False, header=False)
        messagebox.showinfo("Excel сохранён", f"Файл сохранён как: {filename}")


    def exit_and_save(self, event=None):
        try:
            self.save_excel()
        except Exception as e:
            print(f"Ошибка при сохранении Excel: {e}")
        self.destroy()


if __name__ == "__main__":
    app = PDFViewer()
    app.mainloop()