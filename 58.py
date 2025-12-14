import tkinter as tk
from tkinter import ttk

class HighlightDemo:
    def __init__(self, root):
        self.root = root
        self.root.title("Treeview → Code Highlight Demo")
        self.root.geometry("900x500")

        self.setup_ui()

    def setup_ui(self):
        # ---------------- Layout ----------------
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill="both", expand=True)

        # Treeview در چپ
        self.tree = ttk.Treeview(self.frame, columns=("start", "end"), show="headings", height=20)
        self.tree.heading("start", text="Start Line")
        self.tree.heading("end", text="End Line")
        self.tree.column("start", width=80)
        self.tree.column("end", width=80)
        self.tree.pack(side="left", fill="y")

        # Textbox برای نمایش کد
        self.code_text = tk.Text(self.frame, wrap="none", font=("Consolas", 12))
        self.code_text.pack(side="left", fill="both", expand=True)

        # اسکرولبار
        scroll = tk.Scrollbar(self.frame, command=self.code_text.yview)
        scroll.pack(side="right", fill="y")
        self.code_text.configure(yscrollcommand=scroll.set)

        # ---------------- Insert Sample Code ----------------
        sample_code = """
def load_excel():
    print("Excel loaded")

def save_report():
    print("Report saved")

def filter_data():
    print("Filtering data...")

def export_pdf():
    print("Exporting PDF...")

def export_excel():
    print("Export Excel file")
        """
        self.code_text.insert("1.0", sample_code)

        # تعریف تگ هایلایت
        self.code_text.tag_config("highlight", background="yellow")

        # ---------------- Insert items into Treeview ----------------
        # توضیح: start_line, end_line
        rows = [
            ("2", "4"),   # تابع load_excel
            ("6", "8"),   # تابع save_report
            ("10", "12"), # تابع filter_data
            ("14", "16"), # تابع export_pdf
            ("18", "20"), # تابع export_excel
        ]

        for r in rows:
            self.tree.insert("", "end", values=r)

        # ---------------- Bind Hover Event ----------------
        self.tree.bind("<Motion>", self.on_tree_hover)

    # ---------------- Highlight Functions ----------------
    def clear_highlight(self):
        self.code_text.tag_remove("highlight", "1.0", "end")

    def highlight_range(self, start_line, end_line):
        self.clear_highlight()
        self.code_text.tag_add("highlight", f"{start_line}.0", f"{end_line}.0")

    def on_tree_hover(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            start_line, end_line = self.tree.item(item, "values")
            self.highlight_range(start_line, end_line)
        else:
            self.clear_highlight()


# ---------------- Run App ----------------
if __name__ == "__main__":
    root = tk.Tk()
    app = HighlightDemo(root)
    root.mainloop()
