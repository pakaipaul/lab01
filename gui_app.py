import os
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import win32com.client

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


def get_outlook_folder(path):
    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    path = path.strip()
    if path == "" or path.lower() == "inbox":
        return namespace.GetDefaultFolder(6)
    parts = path.split("\\")
    folder = None
    try:
        if parts[0].lower() in {"inbox", "sent items", "drafts", "deleted items", "outbox"}:
            default_map = {"inbox": 6, "sent items": 5, "drafts": 16, "deleted items": 3, "outbox": 4}
            folder = namespace.GetDefaultFolder(default_map[parts[0].lower()])
            parts = parts[1:]
        else:
            folder = namespace.Folders[parts[0]]
            parts = parts[1:]
        for part in parts:
            folder = folder.Folders[part]
    except Exception:
        raise ValueError("Nu am putut găsi folderul Outlook specificat. Verifică numele și structura folderului.")
    return folder


def search_outlook_file():
    folder_path = entry_outlook_folder.get().strip()
    attachment_name = entry_outlook_file.get().strip()
    query = entry_outlook_query.get().strip()
    if query == "":
        messagebox.showwarning("Caută Outlook", "Te rog introdu un termen de căutare.")
        return
    try:
        folder = get_outlook_folder(folder_path)
        messages = folder.Items
        temp_dir = os.path.join(os.getcwd(), 'temp_outlook')
        os.makedirs(temp_dir, exist_ok=True)
        found = []
        for msg in messages:
            try:
                for attachment in msg.Attachments:
                    if attachment_name and attachment_name.lower() not in attachment.FileName.lower():
                        continue
                    filename = attachment.FileName
                    if not filename.lower().endswith(('.xlsx', '.xls')):
                        continue
                    path = os.path.join(temp_dir, filename)
                    attachment.SaveAsFile(path)
                    df = pd.read_excel(path)
                    mask = df.apply(lambda row: row.astype(str).str.contains(query, case=False, na=False).any(), axis=1)
                    if mask.any():
                        found.append((msg.Subject, filename, df[mask].head(5).to_string(index=False)))
            except Exception:
                continue
        if not found:
            messagebox.showinfo("Rezultate Outlook", f"Nu s-au găsit rezultate pentru '{query}' în folderul Outlook selectat.")
            return
        result_text = ''
        for subject, filename, preview in found:
            result_text += f"Mesaj: {subject}\nFișier: {filename}\n{preview}\n---\n"
        if len(result_text) > 1500:
            result_text = result_text[:1500] + "..."
        messagebox.showinfo("Rezultate Outlook", result_text)
    except ValueError as ve:
        messagebox.showerror("Eroare Outlook", str(ve))
    except Exception as e:
        messagebox.showerror("Eroare Outlook", f"Eroare la căutarea în Outlook: {str(e)}")

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

# Tab Outlook
frame_outlook = ttk.Frame(notebook)
notebook.add(frame_outlook, text='Outlook')
tk.Label(frame_outlook, text="Folder Outlook (ex: Inbox sau Cont Outlook\\Subfolder):").pack(padx=5, pady=(10,0), anchor='w')
entry_outlook_folder = tk.Entry(frame_outlook)
entry_outlook_folder.pack(fill='x', padx=5)
tk.Label(frame_outlook, text="Nume fișier atașat (opțional):").pack(padx=5, pady=(10,0), anchor='w')
entry_outlook_file = tk.Entry(frame_outlook)
entry_outlook_file.pack(fill='x', padx=5)
tk.Label(frame_outlook, text="Termen căutare în fișier:").pack(padx=5, pady=(10,0), anchor='w')
entry_outlook_query = tk.Entry(frame_outlook)
entry_outlook_query.pack(fill='x', padx=5)
btn_outlook_search = tk.Button(frame_outlook, text="Caută în Outlook", command=search_outlook_file)
btn_outlook_search.pack(pady=10)

# Rulare aplicație
root.mainloop()