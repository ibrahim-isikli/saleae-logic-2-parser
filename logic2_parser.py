import pandas as pd
import os
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox

def process_excel(file_path, threshold):
    try:
        # Dosyayı yükle
        df = pd.read_csv(file_path)

        # İşlem 1: Zaman farkına göre boş satır ekleme
        rows = []
        prev_time = None

        for _, row in df.iterrows():
            if prev_time is not None:
                time_diff = row['start_time'] - prev_time
                if time_diff > threshold:
                    # Boş satır ekle
                    rows.append([None, None, None, None, None])
            rows.append(row.tolist())
            prev_time = row['start_time']

        df_with_blank_rows = pd.DataFrame(rows, columns=df.columns)

        # İşlem 2: Byte verilerini hücre hücre yerleştirme
        processed_data = []
        current_row = []

        for _, row in df_with_blank_rows.iterrows():
            if pd.isnull(row['start_time']):  # Boş satır varsa
                if current_row:
                    processed_data.append(current_row)
                current_row = []
            else:
                current_row.append(row['data'])

        if current_row:  # Son satırları da ekle
            processed_data.append(current_row)

        # Yeni DataFrame oluştur
        max_len = max(len(row) for row in processed_data)
        parsed_df = pd.DataFrame(
            [row + [None] * (max_len - len(row)) for row in processed_data]
        )

        # Yeni dosyayı kaydet
        output_file = os.path.splitext(file_path)[0] + "_parsed.csv"
        parsed_df.to_csv(output_file, index=False, header=False)

        messagebox.showinfo("Başarılı", f"İşlem tamamlandı! Dosya kaydedildi:\n{output_file}")
    except Exception as e:
        messagebox.showerror("Hata", f"İşlem sırasında bir hata oluştu:\n{str(e)}")

def select_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV Dosyaları", "*.csv")],
        title="CSV Dosyasını Seçin"
    )
    if file_path:
        entry_file_path.delete(0, "end")
        entry_file_path.insert(0, file_path)

def run_process():
    file_path = entry_file_path.get().strip()
    threshold_str = entry_threshold.get().strip()

    if not file_path:
        messagebox.showwarning("Uyarı", "Lütfen bir dosya seçin!")
        return

    try:
        threshold = float(threshold_str)
        process_excel(file_path, threshold)
    except ValueError:
        messagebox.showerror("Hata", "Lütfen geçerli bir threshold (eşik) değeri girin!")

# GUI oluşturma
root = Tk()
root.title("Excel İşlem Aracı")

# Dosya seçimi
Label(root, text="CSV Dosyası:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_file_path = Entry(root, width=50)
entry_file_path.grid(row=0, column=1, padx=10, pady=10)
Button(root, text="Gözat", command=select_file).grid(row=0, column=2, padx=10, pady=10)

# Threshold girişi
Label(root, text="Threshold (eşik) değeri:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
entry_threshold = Entry(root, width=10)
entry_threshold.grid(row=1, column=1, padx=10, pady=10, sticky="w")

# Çalıştır butonu
Button(root, text="İşlemi Başlat", command=run_process).grid(row=2, column=0, columnspan=3, pady=20)

# GUI'yi başlat
root.mainloop()
