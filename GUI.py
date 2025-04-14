import pickle
import tkinter as tk
from tkinter import messagebox
import win32com.client
import pythoncom
from PIL import Image
import base64
from io import BytesIO

def speak(text):
    pythoncom.CoInitialize()
    speak = win32com.client.Dispatch("SAPI.SpVoice")
    speak.Speak(text)

def image_to_base64(image_path):
    img = Image.open(image_path)
    buffered = BytesIO()
    img.save(buffered, format="JPEG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return img_str

def load_background_image(image_path, root):
    img_str = image_to_base64(image_path)
    background = tk.Label(root, image=None)  # Placeholder for image
    background.place(relwidth=1, relheight=1)  # Adjust for full screen

def classify_email(msg, model, cv):
    data = [msg]
    vect = cv.transform(data).toarray()
    prediction = model.predict(vect)
    result = prediction[0]
    return result

def on_predict_click():
    msg = email_entry.get()
    if msg == "":
        messagebox.showwarning("Input Error", "Please enter a text!")
    else:
        result = classify_email(msg, model, cv)
        if result == 1:
            result_label.config(text="This is a SPAM EMAIL !!!", fg="red")
            speak("This is a SPAM EMAIL")
        else:
            result_label.config(text="This is a HAM Email", fg="green")
            speak("This is a HAM EMAIL")

# Load model and vectorizer
model = pickle.load(open("spam.pkl", "rb"))
cv = pickle.load(open("vectorizer.pkl", "rb"))

# Set up the GUI window
root = tk.Tk()
root.title("Email Spam Classification")
root.geometry("600x400")

# Set background image (optional)
load_background_image("email1.jpg", root)

# Add title
title_label = tk.Label(root, text="Email Spam Classification", font=("Arial", 16))
title_label.pack(pady=20)

# Add email input field
email_entry = tk.Entry(root, width=50, font=("Arial", 12))
email_entry.pack(pady=10)

# Add prediction button
predict_button = tk.Button(root, text="Predict", font=("Arial", 12), command=on_predict_click)
predict_button.pack(pady=10)

# Add result label
result_label = tk.Label(root, text="", font=("Arial", 14))
result_label.pack(pady=20)

# Run the GUI
root.mainloop()
