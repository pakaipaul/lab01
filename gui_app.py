import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd

def show_about():
    messagebox.showinfo("Despre", "Aceasta este o aplicație GUI simplă. Aici poți adăuga informații despre proiect sau companie.")

def submit_contact():
    nume = entry_nume.get()
    email = entry_email.get()
    mesaj = text_mesaj.get("1.0", tk.END)
    messagebox.showinfo("Contact", f"Mesaj trimis!\nNume: {nume}\nEmail: {email}\nMesaj: {mesaj}")

def search():
    query = entry_search.get()
    if query.strip() == "":
        messagebox.showwarning("Caută", "Te rog introdu un termen de căutare.")
    else:
        try:
            # Încarcă fișierul Excel (presupunem că există 'data.xlsx' în același director)
            df = pd.read_excel('data.xlsx')
            # Caută în toate coloanele, insensibil la majuscule
            mask = df.apply(lambda row: row.astype(str).str.contains(query, case=False, na=False).any(), axis=1)
            results = df[mask]
            if results.empty:
                messagebox.showinfo("Rezultate căutare", f"Nu s-au găsit rezultate pentru '{query}'.")
            else:
                # Afișează primele rezultate (trunchiază dacă e prea lung)
                result_text = results.head(10).to_string()  # Primele 10 rânduri
                if len(result_text) > 1000:
                    result_text = result_text[:1000] + "..."
                messagebox.showinfo("Rezultate căutare", f"Rezultatele pentru '{query}':\n\n{result_text}")
        except FileNotFoundError:
            messagebox.showerror("Eroare", "Fișierul 'data.xlsx' nu a fost găsit. Te rog adaugă un fișier Excel în directorul proiectului.")
        except Exception as e:
            messagebox.showerror("Eroare", f"Eroare la căutare: {str(e)}")

# Creare fereastră principală
root = tk.Tk()
root.title("Aplicație GUI Simplă")
root.geometry("400x400")

# Notebook pentru tab-uri
notebook = ttk.Notebook(root)
notebook.pack(fill='both', expand=True)

# Tab Despre
frame_despre = ttk.Frame(notebook)
notebook.add(frame_despre, text='Despre')
label_despre = tk.Label(frame_despre, text="Despre\n\nAceasta este o pagină de prezentare.\nAici poți adăuga informații despre proiect sau companie.", justify=tk.LEFT)
label_despre.pack(pady=20)
btn_despre = tk.Button(frame_despre, text="Află mai mult", command=show_about)
btn_despre.pack()

# Tab Contact
frame_contact = ttk.Frame(notebook)
notebook.add(frame_contact, text='Contact')
tk.Label(frame_contact, text="Nume:").pack()
entry_nume = tk.Entry(frame_contact)
entry_nume.pack()
tk.Label(frame_contact, text="Email:").pack()
entry_email = tk.Entry(frame_contact)
entry_email.pack()
tk.Label(frame_contact, text="Mesaj:").pack()
text_mesaj = tk.Text(frame_contact, height=5)
text_mesaj.pack()
btn_contact = tk.Button(frame_contact, text="Trimite", command=submit_contact)
btn_contact.pack(pady=10)

# Tab Caută
frame_cauta = ttk.Frame(notebook)
notebook.add(frame_cauta, text='Caută')
tk.Label(frame_cauta, text="Caută ceva:").pack()
entry_search = tk.Entry(frame_cauta)
entry_search.pack()
btn_search = tk.Button(frame_cauta, text="Caută", command=search)
btn_search.pack(pady=10)

# Rulare aplicație
root.mainloop()