from tkinter import Tk, Label, Button, Entry, filedialog, StringVar, ttk, messagebox, Frame
import os
from PIL import Image
import fitz  # PyMuPDF
import zipfile
import shutil
import tempfile

file_path_full = ""
temp_output_path = ""

def get_file_size(path):
    size = os.path.getsize(path)
    return size / 1024  # KB

def compress_office_file(original_path):
    # Buat salinan file
    temp_dir = tempfile.mkdtemp()
    temp_output = os.path.join(temp_dir, os.path.basename(original_path))
    shutil.copyfile(original_path, temp_output)

    # Extract isi file
    extract_dir = os.path.join(temp_dir, "extracted")
    with zipfile.ZipFile(temp_output, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

    # Kompres gambar di dalam file
    media_dir = os.path.join(extract_dir, "word", "media")
    if not os.path.exists(media_dir):
        media_dir = os.path.join(extract_dir, "ppt", "media")
    if os.path.exists(media_dir):
        for img_file in os.listdir(media_dir):
            img_path = os.path.join(media_dir, img_file)
            try:
                img = Image.open(img_path)
                img.save(img_path, optimize=True, quality=60)
            except Exception:
                pass

    # Buat file baru dengan isi hasil ekstrak
    output_path = original_path.replace(".", "_hasil_kompres.")
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as new_zip:
        for root_dir, _, files in os.walk(extract_dir):
            for file in files:
                full_path = os.path.join(root_dir, file)
                relative_path = os.path.relpath(full_path, extract_dir)
                new_zip.write(full_path, relative_path)
    shutil.rmtree(temp_dir)
    return output_path

def kompres_file(path, kualitas):
    ext = os.path.splitext(path)[1].lower()
    output = path.replace(ext, f"_hasil_kompres{ext}")
    try:
        if ext in ['.jpg', '.jpeg', '.png']:
            kualitas_map = {"Tinggi": 80, "Sedang": 50, "Rendah": 30}
            img = Image.open(path)
            img.save(output, quality=kualitas_map[kualitas], optimize=True)
        elif ext == '.pdf':
            doc = fitz.open(path)
            doc.save(output, deflate=True)
        elif ext in ['.docx', '.pptx']:
            output = compress_office_file(path)
        else:
            return None
        return output
    except Exception as e:
        print("Error:", e)
        return None

def pilih_file():
    global file_path_full
    path = filedialog.askopenfilename(filetypes=[("All Supported", "*.jpg *.jpeg *.png *.pdf *.docx *.pptx")])
    if path:
        file_path_full = path
        file_name = os.path.basename(path)
        file_path_var.set(file_name)

def proses_kompresi():
    global temp_output_path
    if not file_path_full:
        messagebox.showwarning("Pilih File", "Silakan pilih file terlebih dahulu.")
        return
    kualitas = kualitas_var.get()
    temp_output_path = kompres_file(file_path_full, kualitas)
    if temp_output_path:
        before_size = get_file_size(file_path_full)
        after_size = get_file_size(temp_output_path)
        preview_var.set(f"📄 Nama: {os.path.basename(temp_output_path)}\nUkuran Asli: {before_size:.2f} KB\nUkuran Kompresi: {after_size:.2f} KB")
        btn_simpan.config(state="normal")
    else:
        messagebox.showerror("Gagal", "Kompresi gagal atau format tidak didukung.")

def simpan_file():
    if not temp_output_path or not os.path.exists(temp_output_path):
        messagebox.showerror("Gagal", "File belum dikompres atau tidak ditemukan.")
        return
    simpan_path = filedialog.asksaveasfilename(defaultextension=os.path.splitext(temp_output_path)[1],
                                               initialfile=os.path.basename(temp_output_path),
                                               filetypes=[("Semua", "*.*")])
    if simpan_path:
        shutil.move(temp_output_path, simpan_path)
        messagebox.showinfo("Berhasil", f"File disimpan ke:\n{simpan_path}")
        preview_var.set("")
        btn_simpan.config(state="disabled")

# UI Setup
root = Tk()
root.title("🛠️ Kompresor File Cerdas")
root.geometry("560x460")
root.configure(bg="#1e1e1e")

file_path_var = StringVar()
kualitas_var = StringVar(value="Sedang")
preview_var = StringVar()

# Header
Label(root, text="🔻 Kompres File Anda", bg="#1e1e1e", fg="white",
      font=("Helvetica", 16, "bold")).pack(pady=15)

frame = Frame(root, bg="#2e2e2e", bd=2, relief="groove")
frame.pack(padx=20, pady=10, fill="both", expand=False)

Label(frame, text="📂 Pilih File", bg="#2e2e2e", fg="#ffc107",
      font=("Helvetica", 11, "bold")).pack(pady=(10, 2))
Entry(frame, textvariable=file_path_var, font=("Helvetica", 10),
      state="readonly", width=50, bd=2, relief="groove").pack()
Button(frame, text="🗂 Browse", command=pilih_file, bg="#ff9800", fg="white",
       font=("Helvetica", 10, "bold"), width=15).pack(pady=6)

Label(frame, text="🎛 Kualitas Kompresi", bg="#2e2e2e", fg="#f44336",
      font=("Helvetica", 11, "bold")).pack(pady=(10, 2))
ttk.Combobox(frame, textvariable=kualitas_var, values=["Tinggi", "Sedang", "Rendah"],
             state="readonly", width=15).pack(pady=4)

Button(frame, text="🚀 Kompres Sekarang", command=proses_kompresi,
       bg="#c62828", fg="white", font=("Helvetica", 10, "bold"), width=30).pack(pady=10)

Label(root, text="📊 Hasil Kompresi", bg="#1e1e1e", fg="#ffeb3b",
      font=("Helvetica", 12, "bold")).pack(pady=(10, 4))
Label(root, textvariable=preview_var, justify="left", bg="#424242", fg="white",
      font=("Courier", 10), width=60, height=5, anchor="nw", bd=2, relief="ridge").pack(pady=5)

btn_simpan = Button(root, text="💾 Simpan File Hasil", command=simpan_file,
                    state="disabled", bg="#4caf50", fg="white",
                    font=("Helvetica", 10, "bold"), width=30)
btn_simpan.pack(pady=12)

Label(root, text="* Mendukung: JPG, PNG, PDF, DOCX, PPTX", font=("Helvetica", 9),
      bg="#1e1e1e", fg="gray").pack(side="bottom", pady=6)

root.mainloop()
