import json
import os
import re
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkFont
import webbrowser

JSON_FILE = "baza_crteza.json"
DRAWINGS_FOLDER = "crtezi"

FIELDS = [
    "IDENTBROJ", "CRTEZBROJ", "NAZIVDELA", "TEHNPODACI",
    "KATALBROJ", "FORMAT", "ARHIVA", "KOMENTAR",
    "OBJEKAT1", "OBJEKAT2", "OBJEKAT3", "OBJEKAT4",
    "KATALOG", "MAGSIFRA"
]

# --- Load JSON ---
try:
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
except FileNotFoundError:
    messagebox.showerror("Greška", f"{JSON_FILE} nije pronađen!")
    data = []
except json.JSONDecodeError:
    messagebox.showerror("Greška", f"{JSON_FILE} nije validan JSON fajl!")
    data = []

current_index = 0
search_results = []


# --- Tkinter setup ---
root = tk.Tk()
root.title("Baza Crteza")
root.geometry("1400x800")

# Fonts and styles
default_font = tkFont.nametofont("TkDefaultFont")
default_font.configure(family="Arial", size=12)
root.option_add("*Font", default_font)

style = ttk.Style()
style.configure("TButton", font=("Arial", 11, "bold"), padding=(10, 5))


# --- Search frame ---
search_frame = ttk.Frame(root, padding=10, relief="solid", borderwidth=1)
search_frame.pack(fill=tk.X, padx=10, pady=10)

search_field_var = tk.StringVar(value="CRTEZBROJ")
search_value_var = tk.StringVar()

inner_frame = ttk.Frame(search_frame)
inner_frame.pack()

ttk.Label(inner_frame, text="Traži u:").pack(side=tk.LEFT, padx=5)
search_field_menu = ttk.Combobox(inner_frame, textvariable=search_field_var, values=FIELDS, width=15)
search_field_menu.pack(side=tk.LEFT, padx=5)
ttk.Label(inner_frame, text="Vrednost:").pack(side=tk.LEFT, padx=5)
search_value_entry = ttk.Entry(inner_frame, textvariable=search_value_var, width=30, justify='center')
search_value_entry.pack(side=tk.LEFT, padx=5, pady=5, ipadx=5, ipady=8)

def normalize(s):
    if not s:
        return ""
    return re.sub(r"\W+", "", str(s).lower())

def load_record(idx):
    global current_index
    if not data:
        return
    rec = data[idx]
    for f in FIELDS:
        entry = entries.get(f)
        if entry:
            entry.delete(0, tk.END)
            val = rec.get(f)
            entry.insert(0, "" if val is None else str(val))
    current_index = idx
    record_number_var.set(f"{idx+1}/{len(data)}")

def do_search():
    global search_results, search_index
    key = search_field_var.get()
    val = re.sub(r"\W+", "", search_value_var.get().strip().lower())
    search_results = [
        i for i, rec in enumerate(data)
        if val in normalize(rec.get(key))
    ]
    if search_results:
        search_index = 0
        load_record(search_results[search_index])
    else:
        messagebox.showinfo("Pretraga", "Ništa nije pronađeno!")

def next_result():
    global search_index
    if not search_results:
        return
    search_index += 1
    if search_index >= len(search_results):
        search_index = 0
    load_record(search_results[search_index])

ttk.Button(inner_frame, text="Traži", command=do_search).pack(side=tk.LEFT, padx=5)
ttk.Button(inner_frame, text="Sledeći", command=next_result).pack(side=tk.LEFT, padx=5)


# --- Record Frame ---
record_frame = ttk.Frame(root, padding=20)
record_frame.pack(fill=tk.BOTH, expand=True)

entries = {}

def make_entry(parent, label_text, field_name, width=30):
    frame = ttk.Frame(parent)
    frame.pack(fill=tk.X, pady=4)
    ttk.Label(frame, text=label_text, width=15, anchor="w").pack(side=tk.LEFT)
    e = ttk.Entry(frame, width=width)
    e.pack(side=tk.LEFT, fill=tk.X, ipady=2)
    entries[field_name] = e
    return e


# --- Section 1: Basic Info ---
section1 = ttk.Frame(record_frame)
section1.pack(fill=tk.X, pady=10)

id_crtez_frame = ttk.Frame(section1)
id_crtez_frame.pack(fill=tk.X, pady=5)

# ID Broj
ttk.Label(id_crtez_frame, text="ID Broj:", width=15, anchor="w").pack(side=tk.LEFT)
entries["IDENTBROJ"] = ttk.Entry(id_crtez_frame, width=25)
entries["IDENTBROJ"].pack(side=tk.LEFT, padx=(0, 15), ipady=4)

# Broj crteža
ttk.Label(id_crtez_frame, text="Broj crteža:", width=10, anchor="w").pack(side=tk.LEFT)
entries["CRTEZBROJ"] = ttk.Entry(id_crtez_frame, width=30)
entries["CRTEZBROJ"].pack(side=tk.LEFT, padx=(0, 10), ipady=4)

# Open button
open_btn = ttk.Button(id_crtez_frame, text="Open Drawing")
open_btn.pack(side=tk.LEFT)

make_entry(section1, "Naziv dela:", "NAZIVDELA", width=68)
make_entry(section1, "Tehnički podaci:", "TEHNPODACI", width=68)


# --- Section 2: Catalog Info ---
section2 = ttk.Frame(record_frame, padding=10, relief="solid")
section2.pack(fill=tk.X, pady=10)

make_entry(section2, "Kataloški broj:", "KATALBROJ", width=25)
make_entry(section2, "Format:", "FORMAT", width=10)
make_entry(section2, "Arhiva:", "ARHIVA", width=10)


# --- Section 3: Split - Left entries / Right buttons ---
section3 = ttk.Frame(record_frame, padding=10, relief="solid")
section3.pack(fill=tk.BOTH, expand=True, pady=10)

left_frame = ttk.Frame(section3)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)

make_entry(left_frame, "Komentar:", "KOMENTAR", width=50)
make_entry(left_frame, "Objekat1:", "OBJEKAT1", width=50)
make_entry(left_frame, "Objekat2:", "OBJEKAT2", width=50)
make_entry(left_frame, "Objekat3:", "OBJEKAT3", width=50)
make_entry(left_frame, "Objekat4:", "OBJEKAT4", width=50)
make_entry(left_frame, "Katalog:", "KATALOG", width=50)
make_entry(left_frame, "Magsifra:", "MAGSIFRA", width=50)

right_frame = ttk.Frame(section3)
right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=15, pady=5)

record_number_var = tk.StringVar()

def add_record():
    global current_index, search_results
    # Clear all entry fields
    for f in FIELDS:
        entries[f].delete(0, tk.END)
    current_index = len(data)  # new record index (at end)
    record_number_var.set(f"{current_index+1}/{len(data)+1}")
    search_results = []  # clear search

def save_record():
    global current_index
    if current_index is None:
        return
    if current_index >= len(data):
        # New record, append
        rec = {}
        data.append(rec)
    else:
        rec = data[current_index]

    for f in FIELDS:
        val = entries[f].get().strip()
        # Convert empty string to None (null in JSON)
        if val == "":
            rec[f] = None
        else:
            # Convert IDENTBROJ to int if possible
            if f == "IDENTBROJ":
                try:
                    rec[f] = int(val)
                except ValueError:
                    rec[f] = val  # fallback if not a number
            else:
                rec[f] = val

    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    messagebox.showinfo("Save", "Record saved!")

def delete_record():
    global current_index, search_results
    if not data or current_index is None:
        return
    confirm = messagebox.askyesno("Delete", "Are you sure you want to delete this record?")
    if not confirm:
        return

    # Remove from data
    del data[current_index]
    with open(JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    messagebox.showinfo("Delete", "Record deleted!")

    # Reset current index
    if len(data) == 0:
        current_index = None
        for e in entries.values():
            e.delete(0, tk.END)
        record_number_var.set("0/0")
    else:
        if current_index >= len(data):
            current_index = len(data) - 1
        load_record(current_index)

    search_results = []  # clear search results
    
    # Reset current index
    if current_index >= len(data):
        current_index = len(data) - 1

    # Clear search results to prevent invalid indices
    search_results = []
    
    # Load next available record
    if data:
        load_record(current_index)
    else:
        # Clear all entries if no records left
        for e in entries.values():
            e.delete(0, tk.END)
        record_number_var.set("0/0")

add_button = ttk.Button(right_frame, text="Add New Record")
add_button.pack(pady=5, fill=tk.X)
save_button = ttk.Button(right_frame, text="Save Record")
save_button.pack(pady=5, fill=tk.X)
delete_button = ttk.Button(right_frame, text="Delete Record")
delete_button.pack(pady=5, fill=tk.X)

add_button.config(command=add_record)
save_button.config(command=save_record)
delete_button.config(command=delete_record)


# --- Navigation Frame ---
nav_frame = ttk.Frame(root, padding=15)
nav_frame.pack(fill=tk.X)

def first_record():
    global current_index
    current_index = 0
    load_record(current_index)

def prev_record():
    global current_index
    if current_index > 0:
        current_index -= 1
        load_record(current_index)

def goto_record():
    global current_index
    try:
        idx = int(record_number_entry.get()) - 1
        if 0 <= idx < len(data):
            current_index = idx
            load_record(current_index)
    except ValueError:
        pass

def next_record_nav():
    global current_index
    if current_index < len(data)-1:
        current_index += 1
        load_record(current_index)

def last_record():
    global current_index
    current_index = len(data)-1
    load_record(current_index)

ttk.Button(nav_frame, text="<< Prvi", command=first_record).pack(side=tk.LEFT, padx=5)
ttk.Button(nav_frame, text="<< Prethodni", command=prev_record).pack(side=tk.LEFT, padx=5)
ttk.Label(nav_frame, text="Unos:").pack(side=tk.LEFT, padx=5)
record_number_entry = ttk.Entry(nav_frame, width=12, textvariable=record_number_var, justify='center')
record_number_entry.pack(side=tk.LEFT, ipady=5)
ttk.Button(nav_frame, text="Idi", command=goto_record).pack(side=tk.LEFT, padx=5)
ttk.Button(nav_frame, text="Sledeći >>", command=next_record_nav).pack(side=tk.LEFT, padx=5)
ttk.Button(nav_frame, text="Poslednji >>", command=last_record).pack(side=tk.LEFT, padx=5)


def open_drawing():
    crtez = entries["CRTEZBROJ"].get().strip()
    if not crtez:
        return
    filename = crtez.replace("/", "-").replace("\\", "-") + ".jpg"
    path = os.path.join(DRAWINGS_FOLDER, filename)
    if os.path.exists(path):
        webbrowser.open(path)
    else:
        messagebox.showinfo("Otvori crtež", f"Fajl nije pronađen: {path}")

open_btn.config(command=open_drawing)



# --- Start ---
if data:
    load_record(0)

root.mainloop()