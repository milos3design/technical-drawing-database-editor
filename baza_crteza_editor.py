import os
import re
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkFont
import webbrowser
import pandas as pd

EXCEL_FILE = "BAZACRTEZA.xlsx"
DRAWINGS_FOLDER = "crtezi"

FIELDS = [
    "IDENTBROJ", "CRTEZBROJ", "NAZIVDELA", "TEHNPODACI",
    "KATALBROJ", "FORMAT", "ARHIVA", "KOMENTAR",
    "OBJEKAT1", "OBJEKAT2", "OBJEKAT3", "OBJEKAT4",
    "KATALOG", "MAGSIFRA"
]

# --- Load Excel ---
try:
    df = pd.read_excel(EXCEL_FILE)
    # Replace NaN with None for consistency
    df = df.where(pd.notna(df), None)
    data = df.to_dict('records')
except FileNotFoundError:
    messagebox.showerror("Greška", f"{EXCEL_FILE} nije pronađen!")
    data = []
    df = pd.DataFrame(columns=FIELDS)
except Exception as e:
    messagebox.showerror("Greška", f"Greška pri učitavanju {EXCEL_FILE}: {str(e)}")
    data = []
    df = pd.DataFrame(columns=FIELDS)

current_index = None
search_results = []
search_index = 0


# --- Tkinter setup ---
root = tk.Tk()
root.title("Baza Crteza Pomoćne mehanizacije - Editor")
root.geometry("1400x800")

record_number_var = tk.StringVar()

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

def save_to_excel():
    """Helper function to save DataFrame to Excel with error handling"""
    try:
        df.to_excel(EXCEL_FILE, index=False)
        return True
    except PermissionError:
        messagebox.showerror(
            "Greška - Fajl je zaključan", 
            f"Ne mogu da sačuvam '{EXCEL_FILE}'!\n\n"
            f"Fajl je verovatno otvoren u Excel programu.\n"
            f"Molim vas zatvorite fajl u Excel-u i pokušajte ponovo."
        )
        return False
    except Exception as e:
        messagebox.showerror("Greška pri čuvanju", f"Greška: {str(e)}")
        return False

search_counter_var = tk.StringVar(value="")

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
        search_counter_var.set(f"{search_index+1}/{len(search_results)}")
    else:
        search_index = 0
        search_counter_var.set("")
        messagebox.showinfo("Pretraga", "Ništa nije pronađeno!")

def next_result():
    global search_index
    if not search_results:
        return
    search_index += 1
    if search_index >= len(search_results):
        search_index = 0
    load_record(search_results[search_index])
    search_counter_var.set(f"{search_index+1}/{len(search_results)}")

ttk.Button(inner_frame, text="Traži", command=do_search).pack(side=tk.LEFT, padx=5)
ttk.Button(inner_frame, text="Sledeći", command=next_result).pack(side=tk.LEFT, padx=5)
ttk.Label(inner_frame, textvariable=search_counter_var, width=8, anchor="center").pack(side=tk.LEFT, padx=5)



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

make_entry(section1, "Naziv dela:", "NAZIVDELA", width=68)
make_entry(section1, "Tehnički podaci:", "TEHNPODACI", width=68)

# Open button
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

open_btn = ttk.Button(id_crtez_frame, text="Otvori Crtež")
open_btn.pack(side=tk.LEFT)
open_btn.config(command=open_drawing)


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

def add_record():
    global current_index, search_results
    # Clear all entry fields
    for f in FIELDS:
        entries[f].delete(0, tk.END)

    if not data:
        current_index = 0  # First record
    else:
        current_index = len(data)  # New record at end
    
    record_number_var.set(f"{current_index+1}/{len(data)+1}")
    search_results = []

def save_record():
    global current_index, data, df

    if current_index is None:
        return
    
    if current_index >= len(data):
        # New record at the end
        rec = {}
        data.append(rec)
    else:
        # Existing record
        rec = data[current_index]
    
    ident_str = entries["IDENTBROJ"].get().strip()
    if ident_str == "":
        rec["IDENTBROJ"] = None
    else:
        try:
            rec["IDENTBROJ"] = int(ident_str)
        except ValueError:
            messagebox.showerror("Greška", "ID Broj mora biti broj!")
            return

    for f in FIELDS:
        if f == "IDENTBROJ":
            continue
        val = entries[f].get().strip()
        rec[f] = None if val == "" else val

    # Convert data back to DataFrame and save to Excel
    df = pd.DataFrame(data)
    if save_to_excel():
        messagebox.showinfo("Sačuvaj", "Unos sačuvan!")

def delete_record():
    global current_index, search_results, data, df
    if not data or current_index is None:
        return
    
    crtez_broj = data[current_index].get("CRTEZBROJ", "")
    if crtez_broj:
        message = f"Da li ste sigurni da želite da izbrišete crtež {crtez_broj}?"
    else:
        message = "Da li ste sigurni da želite da izbrišete ovaj unos?"

    confirm = messagebox.askyesno("BRISANJE", message)
    if not confirm:
        return

    # Remove from data
    rec = data[current_index].copy()
    del data[current_index]
    
    # Convert data back to DataFrame and save to Excel
    df = pd.DataFrame(data)
    if not save_to_excel():
        # Restore the deleted record if save failed
        data.insert(current_index, rec)
        return
    
    messagebox.showinfo("Brisanje", "Unos izbrisan!")

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
   
save_button = ttk.Button(right_frame, text="Sačuvaj Unos")
save_button.pack(pady=5, fill=tk.X)
add_button = ttk.Button(right_frame, text="Dodaj Novi Unos")
add_button.pack(pady=5, fill=tk.X)
delete_button = ttk.Button(right_frame, text="Izbriši Unos")
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
    if not data or current_index is None:
        return
    if current_index > 0:
        current_index -= 1
        load_record(current_index)

def goto_record():
    global current_index
    if not data:
        return
    try:
        idx = int(record_number_entry.get()) - 1
        if 0 <= idx < len(data):
            current_index = idx
            load_record(current_index)
    except ValueError:
        pass

def next_record_nav():
    global current_index
    if not data or current_index is None:
        return

    if current_index < len(data)-1:
        current_index += 1
        load_record(current_index)

def last_record():
    global current_index
    if not data:
        return
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


# --- Start ---
if data:
    current_index = 0
    load_record(0)
else:
    record_number_var.set("0/0")

root.mainloop()