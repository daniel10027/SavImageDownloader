import os
import pandas as pd
import requests
from PIL import Image, ImageTk
from io import BytesIO
import tkinter as tk
from tkinter import ttk, filedialog

# Fonction pour créer le dossier "images" s'il n'existe pas
def create_images_folder():
    if not os.path.exists("images"):
        os.makedirs("images")

# Fonction pour télécharger et redimensionner une image
def download_and_resize_image(url, ip):
    response = requests.get(url)
    img = Image.open(BytesIO(response.content))
    img = img.resize((413, 532))
    img.save(os.path.join("images", f"{ip}.jpg"))
    return img

# Fonction pour démarrer le processus
def start_process():
    filepath = filedialog.askopenfilename()
    if not filepath:
        return
    
    df = pd.read_excel(filepath)
    total_images = len(df)
    progress_step = 100 / total_images
    
    create_images_folder()
    
    for index, row in df.iterrows():
        url = row['PHOTO']
        ip = row['IP']
        
        try:
            img = download_and_resize_image(url, ip)
        except Exception as e:
            print(f"Erreur lors du traitement de l'image {ip}: {e}")
            continue
        
        progress_var.set((index + 1) * progress_step)
        progress_label.config(text=f"Téléchargement de l'image {index + 1}/{total_images} complété.")
        
        img_tk = ImageTk.PhotoImage(img)
        image_label.config(image=img_tk)
        image_label.image = img_tk
        
        root.update_idletasks()

# Création de la fenêtre principale
root = tk.Tk()
root.title("Téléchargeur d'Images")
root.geometry("600x400")

# Barre de progression
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100)
progress_bar.pack(pady=20)

progress_label = tk.Label(root, text="Progression du téléchargement")
progress_label.pack(pady=5)

# Label pour afficher l'image téléchargée
image_label = tk.Label(root)
image_label.pack(pady=20)

# Bouton pour démarrer le processus
start_button = tk.Button(root, text="Sélectionner le fichier Excel et démarrer", command=start_process)
start_button.pack(pady=20)

# Boucle principale de Tkinter
root.mainloop()