import os
import csv
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

def create_photo_catalog(directory):
    catalog = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.lower().endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp')):
                filepath = os.path.join(root, file)
                stat = os.stat(filepath)
                catalog.append({
                    'filename': file,
                    'path': filepath,
                    'size': stat.st_size,
                    'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                })
    return catalog

def save_to_csv(catalog, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['filename', 'path', 'size', 'modified']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for item in catalog:
            writer.writerow(item)

class PhotoCatalogApp:
    def __init__(self, root):
        self.root = root
        self.root.title("사진대장 생성기")
        self.root.geometry("800x600")

        self.directory = ""

        # 폴더 선택 버튼
        self.select_button = tk.Button(root, text="폴더 선택", command=self.select_directory)
        self.select_button.pack(pady=10)

        # 선택된 폴더 표시
        self.dir_label = tk.Label(root, text="선택된 폴더: 없음")
        self.dir_label.pack()

        # 스캔 버튼
        self.scan_button = tk.Button(root, text="사진 스캔", command=self.scan_photos)
        self.scan_button.pack(pady=10)

        # 결과 표시 (Treeview)
        self.tree = ttk.Treeview(root, columns=('filename', 'path', 'size', 'modified'), show='headings')
        self.tree.heading('filename', text='파일명')
        self.tree.heading('path', text='경로')
        self.tree.heading('size', text='크기 (bytes)')
        self.tree.heading('modified', text='수정 날짜')
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 저장 버튼
        self.save_button = tk.Button(root, text="CSV로 저장", command=self.save_csv)
        self.save_button.pack(pady=10)

        self.catalog = []

    def select_directory(self):
        self.directory = filedialog.askdirectory()
        if self.directory:
            self.dir_label.config(text=f"선택된 폴더: {self.directory}")

    def scan_photos(self):
        if not self.directory:
            messagebox.showerror("오류", "먼저 폴더를 선택하세요.")
            return

        self.catalog = create_photo_catalog(self.directory)

        # Treeview 초기화
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 결과 추가
        for item in self.catalog:
            self.tree.insert('', tk.END, values=(item['filename'], item['path'], item['size'], item['modified']))

        messagebox.showinfo("완료", f"{len(self.catalog)}개의 사진을 찾았습니다.")

    def save_csv(self):
        if not self.catalog:
            messagebox.showerror("오류", "먼저 사진을 스캔하세요.")
            return

        output_file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if output_file:
            save_to_csv(self.catalog, output_file)
            messagebox.showinfo("저장 완료", f"CSV 파일이 {output_file}에 저장되었습니다.")

if __name__ == "__main__":
    root = tk.Tk()
    app = PhotoCatalogApp(root)
    root.mainloop()