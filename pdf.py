from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from PIL import Image
import tempfile
import os
import webbrowser
from docx2pdf import convert


class ImagetoPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("📄 Image + Word to PDF Converter")
        self.root.geometry("550x680")
        self.root.configure(bg="#f0f4f8")
        self.image_paths = []
        self.word_paths = []
        self.output_pdf_name = tk.StringVar()

        self.selected_file_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, bg="#ffffff", fg="black", relief=tk.GROOVE, bd=2)

        self.initialize_ui()
        self.add_drag_and_drop_support()

    def initialize_ui(self):
        title_label = tk.Label(self.root, text="📄 Image + Word to PDF Converter ❤", font=("Arial", 18, "bold"), bg="#007acc", fg="white", pady=10)
        title_label.pack(fill=tk.X)

        select_file_button = tk.Button(
            self.root, text="📂 Select Images or Word Files", command=self.select_files,
            bg="#28a745", fg="white", activebackground="#1e7e34", font=("Arial", 12, "bold")
        )
        select_file_button.pack(pady=(20, 10), ipadx=10, ipady=5)

        drag_drop_label = tk.Label(self.root, text="or drag and drop files below ↓", font=("Arial", 11), bg="#f0f4f8", fg="gray")
        drag_drop_label.pack()

        self.selected_file_listbox.pack(padx=20, pady=(10, 10), fill=tk.BOTH, expand=True)

        label = tk.Label(self.root, text="Enter Output PDF Name (for images):", font=("Arial", 12), bg="#f0f4f8")
        label.pack(pady=(10, 0))

        pdf_name_entry = tk.Entry(self.root, textvariable=self.output_pdf_name, width=35, font=("Arial", 12))
        pdf_name_entry.pack(pady=(0, 15))

        convert_button = tk.Button(
            self.root, text="💾 Convert to PDF", command=self.convert_all_to_pdf,
            bg="#3A5311", fg="white", activebackground="darkgreen", font=("Arial", 12, "bold")
        )
        convert_button.pack(pady=(10, 10), ipadx=10, ipady=5)

        self.status_label = tk.Label(self.root, text="", font=("Arial", 10, "italic"), fg="green", bg="#f0f4f8")
        self.status_label.pack(pady=10)

    def add_drag_and_drop_support(self):
        self.selected_file_listbox.drop_target_register(DND_FILES)
        self.selected_file_listbox.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for file_path in files:
            self.add_file(file_path)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Images or Word Files",
            filetypes=[("Supported Files", "*.png;*.jpg;*.jpeg;*.docx")]
        )
        for f in files:
            self.add_file(f)

    def add_file(self, path):
        path = str(path)
        if path.lower().endswith((".png", ".jpg", ".jpeg")) and path not in self.image_paths:
            self.image_paths.append(path)
        elif path.lower().endswith(".docx") and path not in self.word_paths:
            self.word_paths.append(path)
        self.update_file_listbox()

    def update_file_listbox(self):
        self.selected_file_listbox.delete(0, tk.END)
        for image_path in self.image_paths:
            self.selected_file_listbox.insert(tk.END, f"[IMG] {os.path.basename(image_path)}")
        for word_path in self.word_paths:
            self.selected_file_listbox.insert(tk.END, f"[DOCX] {os.path.basename(word_path)}")

    def convert_all_to_pdf(self):
        if not self.image_paths and not self.word_paths:
            messagebox.showwarning("No Files Selected", "Please select images or Word documents.")
            return

        if self.image_paths:
            self.convert_images_to_pdf()

        if self.word_paths:
            self.convert_word_docs_to_pdf()

    def convert_images_to_pdf(self):
        output_pdf_name = self.output_pdf_name.get().strip() or "output"
        if not output_pdf_name.endswith(".pdf"):
            output_pdf_name += ".pdf"

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=output_pdf_name,
            title="Save Image PDF As"
        )

        if not save_path:
            return

        try:
            pdf = canvas.Canvas(save_path, pagesize=(612, 792))
            available_width, available_height = 540, 720

            for image_path in self.image_paths:
                img = Image.open(image_path)

                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGB")

                scale_factor = min(available_width / img.width, available_height / img.height)
                new_width = int(img.width * scale_factor)
                new_height = int(img.height * scale_factor)
                x = (612 - new_width) / 2
                y = (792 - new_height) / 2

                with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                    temp_path = tmp.name
                    img = img.resize((new_width, new_height))
                    img.save(temp_path, format="JPEG")

                pdf.drawImage(temp_path, x, y, width=new_width, height=new_height)
                pdf.showPage()
                os.remove(temp_path)

            pdf.save()
            self.status_label.config(text=f"✅ Image PDF saved at:\n{save_path}", fg="green")

        except Exception as e:
            messagebox.showerror("Error", f"Error converting images:\n{e}")
            self.status_label.config(text="❌ Image conversion failed.", fg="red")

    def convert_word_docs_to_pdf(self):
        try:
            for word_path in self.word_paths:
                output_path = os.path.splitext(word_path)[0] + ".pdf"
                convert(word_path, output_path)
                messagebox.showinfo("Success", f"✅ Word converted:\n{output_path}")
                webbrowser.open(os.path.dirname(output_path))

        except Exception as e:
            messagebox.showerror("Error", f"Error converting Word file:\n{e}")


def main():
    root = TkinterDnD.Tk()
    app = ImagetoPDFConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
