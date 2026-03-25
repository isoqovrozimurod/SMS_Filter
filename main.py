import pandas as pd
import re
import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox

UZB_CODES = [
    "20", "33", "50", "55", "77", "88",
    "90", "91", "93", "94", "95", "97", "98", "99"
]

def normalize_number(num: str):
    num = re.sub(r"\D", "", num)

    if len(num) == 9 and num[:2] in UZB_CODES:
        return "998" + num

    if len(num) == 12 and num.startswith("998") and num[3:5] in UZB_CODES:
        return num

    if len(num) == 13 and num.startswith("998"):
        num = num[1:]
        if num[3:5] in UZB_CODES:
            return num

    return None  # ❌ chet el va noto‘g‘ri raqamlar

def format_number(num, plus, compact, spaced):
    if plus:
        num = "+" + num

    if spaced:
        return f"{num[:4]} {num[4:6]} {num[6:9]} {num[9:11]} {num[11:13]}"
    if compact:
        return num.replace(" ", "")
    return num

def start_process():
    path = filedialog.askopenfilename(
        title="Excel faylni tanlang",
        filetypes=[("Excel", "*.xlsx *.xls")]
    )
    if not path:
        return

    df = pd.read_excel(path, header=None)

    numbers = []
    for col in df.columns:
        for val in df[col]:
            found = re.findall(r"\d{7,15}", str(val))
            for f in found:
                n = normalize_number(f)
                if n:
                    numbers.append(n)

    numbers = list(dict.fromkeys(numbers))

    if not numbers:
        messagebox.showerror("Xatolik", "Mos O‘zbekiston raqamlari topilmadi")
        return

    plus = plus_var.get()
    compact = compact_var.get()
    spaced = spaced_var.get()

    if compact and spaced:
        messagebox.showwarning("Diqqat", "Faqat bittasini tanlang!")
        return

    formatted = [
        format_number(n, plus, compact, spaced)
        for n in numbers
    ]

    bosh = bosh_entry.get().strip()
    oxir = oxir_entry.get().strip()

    try:
        chunk = int(chunk_entry.get())
    except:
        chunk = 700

    save_path = os.path.join(
        os.path.dirname(path),
        "natija.xlsx"
    )

    writer = pd.ExcelWriter(save_path, engine="openpyxl")

    pd.DataFrame({
        "T/R": range(1, len(formatted) + 1),
        "Telefon": formatted
    }).to_excel(writer, sheet_name="All", index=False)

    idx = 1
    real_chunk = chunk - (1 if bosh else 0) - (1 if oxir else 0)
    real_chunk = max(1, real_chunk)

    for i in range(0, len(formatted), real_chunk):
        part = formatted[i:i + real_chunk]
        full = []
        if bosh:
            full.append(bosh)
        full.extend(part)
        if oxir:
            full.append(oxir)

        pd.DataFrame({
            "T/R": range(1, len(full) + 1),
            "Telefon": full
        }).to_excel(writer, sheet_name=str(idx), index=False)

        idx += 1

    writer.close()

    messagebox.showinfo(
        "Tayyor ✅",
        f"Jami: {len(formatted)} ta\nVaraq: {idx - 1} ta"
    )

    try:
        os.startfile(save_path)
    except:
        subprocess.run(["open", save_path], check=False)

# ================= UI =================

root = tk.Tk()
root.title("📞 Telefon raqam filter")

plus_var = tk.BooleanVar(value=True)
compact_var = tk.BooleanVar()
spaced_var = tk.BooleanVar(value=True)

tk.Checkbutton(root, text="+ belgisi qo‘shilsin", variable=plus_var).pack(anchor="w")
tk.Checkbutton(root, text="Raqamlar yopishtirilsin", variable=compact_var).pack(anchor="w")
tk.Checkbutton(root, text="Raqamlar orasiga bo‘sh joy", variable=spaced_var).pack(anchor="w")

tk.Label(root, text="Boshiga qo‘shiladigan raqam (test):").pack()
bosh_entry = tk.Entry(root, width=30)
bosh_entry.pack()

tk.Label(root, text="Oxiriga qo‘shiladigan raqam (test):").pack()
oxir_entry = tk.Entry(root, width=30)
oxir_entry.pack()

tk.Label(root, text="Har bir varaqda nechta raqam:").pack()
chunk_entry = tk.Entry(root, width=10)
chunk_entry.insert(0, "700")
chunk_entry.pack()

tk.Button(
    root,
    text="🚀 Boshlash",
    command=start_process,
    bg="#4CAF50",
    fg="white",
    width=20
).pack(pady=10)

root.mainloop()
