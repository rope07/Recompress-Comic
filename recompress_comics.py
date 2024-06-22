import os
import rarfile
import zipfile
import shutil
import threading
import time
from PIL import Image

import tkinter as tk
from tkinter import filedialog, ttk
import queue

rarfile.UNRAR_TOOL = r"C:/Program Files/WinRAR/UnRAR.exe"

class FileUploaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("File Uploader")
        self.root.geometry("600x400")

        self.photo_upload = tk.PhotoImage(file='C:/Users/Pero/Desktop/Python-projekti/Recompress Comics/Icon/upload.png')
        self.photo_upload_resized = self.photo_upload.subsample(12,12)

        self.photo_processing = tk.PhotoImage(file='C:/Users/Pero/Desktop/Python-projekti/Recompress Comics/Icon/engineering.png')
        self.photo_processing_resized = self.photo_processing.subsample(12,12)

        self.upload_button = ttk.Button(root, text="Upload Files", command=self.upload_file, image=self.photo_upload_resized)
        self.upload_button.pack(pady=20)

        self.process_button = ttk.Button(root, text="Process Files", command=self.process_file, image=self.photo_processing_resized)
        self.process_button.pack(pady=10)

        self.clear_button = ttk.Button(root, text="Clear Text", command=self.clear_text)
        self.clear_button.place(relx=0.98, rely=1.0, anchor='se', x=-10, y=-10)

        self.file_path_label = ttk.Label(root, text="No file selected")
        self.file_path_label.pack(pady=10)

        self.scrollbar = ttk.Scrollbar(root)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.message_text = tk.Text(root, height=10, width=50, yscrollcommand=self.scrollbar.set)
        self.message_text.pack(pady=20)
        self.scrollbar.config(command=self.message_text.yview)

        self.file_paths = None
        self.lock = threading.Lock()
        self.file_queue = queue.Queue()

    def upload_file(self):
        self.file_paths = filedialog.askopenfilename(multiple=True)
        if self.file_paths:
            self.file_path_label.config(text="\n".join(self.file_paths))
        else:
            self.file_path_label.config(text="No file selected")

    def _process_file_thread(self):
        while not self.file_queue.empty():
            cbr_file_path = self.file_queue.get()
            cbr_file_name = os.path.splitext(os.path.basename(cbr_file_path))[0]
            try:
                work(cbr_file_path, self.message_text, self.lock)
                self.message_text.insert(tk.END, f"Work completed successfully for {cbr_file_name}.\n")
            except Exception as e:
                self.message_text.insert(tk.END, f"Error during processing {cbr_file_name}: {str(e)}\n")
            self.file_queue.task_done()

    def process_file(self):
        if self.file_paths:
            for file_path in self.file_paths:
                self.file_queue.put(file_path)

            self.message_text.insert(tk.END, "Processing files...\n")
            self.root.update()
            process_thread = threading.Thread(target=self._process_file_thread)
            process_thread.start()
        else:
            self.message_text.insert(tk.END, "Please select a file first.\n")

    def clear_text(self):
        self.message_text.delete('1.0', tk.END)
        self.folder_path_label.config(text="No folders selected")

    def get_file_path(self):
        return self.file_path
    
def delete_original_folder(folder_path):
    try:
        shutil.rmtree(folder_path)
        print(f'Source folder {folder_path} has been deleted.')
    except Exception as e:
        print(f'Error deleting source folder: {e}')

def detect_cbr_compression(cbr_path):
    with open(cbr_path, 'rb') as file:
        file_header = file.read(4)
    if file_header.startswith(b'PK'):
        return 'zip'
    elif file_header.startswith(b'Rar!'):
        return 'rar'
    else:
        return 'unknown'

def extract_cbr_to_folder(cbr_path):
    if not cbr_path.lower().endswith('.cbr'):
        raise ValueError("The file is not a CBR file.")
    
    folder_name = os.path.splitext(os.path.basename(cbr_path))[0] + "-original"    
    output_dir = os.path.join(os.path.dirname(cbr_path), folder_name)

    cbr_compression = detect_cbr_compression(cbr_path)
    
    if cbr_compression == 'rar':
        with rarfile.RarFile(cbr_path) as cbr:
            os.makedirs(output_dir, exist_ok=True)
            cbr.extractall(output_dir)
    elif cbr_compression == 'zip':
        with zipfile.ZipFile(cbr_path, 'r') as zip_ref:
            os.makedirs(output_dir, exist_ok=True)
            zip_ref.extractall(output_dir)
    else:
        raise ValueError(f"Unknown compression type for file: '{cbr_path}'")
    
    print(f"Extracted {cbr_path} to {output_dir}")    
    os.remove(cbr_path)
    print(f"Deleted {cbr_path}")
    return folder_name

def compress_image(image_path, output_path, quality):
    with Image.open(image_path) as img:
        img.save(output_path, "JPEG", optimize=True, quality=quality)

def compress_images_in_folder(folder_path, output_folder_path, quality):
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)
    
    for filename in os.listdir(folder_path):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            image_path = os.path.join(folder_path, filename)
            output_path = os.path.join(output_folder_path, filename)

            #print(f'Compressing {image_path}...')
            compress_image(image_path, output_path, quality)
            #print(f'Saved compressed image to {output_path}')

def has_subfolder(folder_path):
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):
            return True
    return None

def get_subfolder(folder_path):
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):
            return item_path
    return None

def get_folder_size(folder_path):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if os.path.isfile(file_path):
                total_size += os.path.getsize(file_path)
    return total_size

def zip_folder(folder_path, output_path):
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, start=folder_path)
                zipf.write(file_path, arcname)

def compress_folders_in_directory(directory_path, lock):
    for item in os.listdir(directory_path):
        item_path = os.path.join(directory_path, item)
        if os.path.isdir(item_path):
            zip_output_path = os.path.join(directory_path, f"{item}.zip")
            cbr_output_path = os.path.join(directory_path, f"{item}.cbr")
            
            zip_folder(item_path, zip_output_path)
            
            with lock:
                os.rename(zip_output_path, cbr_output_path)
            print(f"Compressed {item_path} into {cbr_output_path}")
            
            shutil.rmtree(item_path)
            print(f"Deleted folder {item_path}")

def work(cbr_file_path, message_text, lock):
    quality = 50

    message_text.insert(tk.END, "Extracting CBR file...\n")
    original_folder_name = extract_cbr_to_folder(cbr_file_path)

    original_folder_path = os.path.join(os.path.dirname(cbr_file_path),original_folder_name)

    if not os.listdir(original_folder_path):
        delete_original_folder(original_folder_path)

    output_folder_name = original_folder_name.replace('-original', '')
    output_folder_path = os.path.join(os.path.dirname(cbr_file_path), output_folder_name)

    message_text.insert(tk.END, "Compressing images...\n")
    if has_subfolder(original_folder_path):
        original_subfolder_path = get_subfolder(original_folder_path)
        compress_images_in_folder(original_subfolder_path, output_folder_path, quality)
    else:
        compress_images_in_folder(original_folder_path, output_folder_path, quality)

    original_folder_size_in_bytes = get_folder_size(original_folder_path)
    original_folder_size_in_MB = original_folder_size_in_bytes / (1024*1024)
    output_folder_size_in_bytes = get_folder_size(output_folder_path)
    output_folder_size_in_MB = output_folder_size_in_bytes / (1024*1024)

    compress_rate = (1-(output_folder_size_in_bytes/original_folder_size_in_bytes)) * 100

    message_text.insert(tk.END, f"Folder size before compressing: {original_folder_size_in_MB:.2f} MB\n")
    message_text.insert(tk.END, f"Folder size after compressing: {output_folder_size_in_MB:.2f} MB\n")
    message_text.insert(tk.END, f"Compressing saved {compress_rate:.2f}% of space\n")

    time.sleep(1)
    delete_original_folder(original_folder_path)

    message_text.insert(tk.END, "Compressing folder to CBR file...\n")
    directory_path = os.path.dirname(cbr_file_path)    
    compress_folders_in_directory(directory_path, lock)

    message_text.insert(tk.END, "Process completed.\n")


if __name__ == "__main__":
    root = tk.Tk()
    app = FileUploaderApp(root)
    root.mainloop()