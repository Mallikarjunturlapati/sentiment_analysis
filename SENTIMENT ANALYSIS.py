import tkinter as tk
from tkinter import messagebox, filedialog
from textblob import TextBlob
import fitz  # PyMuPDF
from docx import Document

def analyze_sentiment(text):
    blob = TextBlob(text)
    sentiment = blob.sentiment.polarity
    
    if sentiment > 0:
        result = "SENTIMENT ANALYSIS:--> Positive <--"
        text_entry.config(bg="lightgreen")
    elif sentiment < 0:
        result = "SENTIMENT ANALYSIS:--> Negative <--"
        text_entry.config(bg="lightcoral")
    else:
        result = "SENTIMENT ANALYSIS:--> Neutral <--"
        text_entry.config(bg="lightgray")

    messagebox.showinfo("SENTIMENT ANALYSIS REPORT IS:", f"{result} (Score: {sentiment:.3f})")

def analyze_text_input():
    text = text_entry.get("1.0", tk.END).strip()
    if len(text) < 2:
        messagebox.showwarning("Warning", "Please enter at least 5 characters.")
        return
    analyze_sentiment(text)

def clear_text():
    text_entry.delete("1.0", tk.END)
    text_entry.config(bg="white")  # Reset background color

def open_document():
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"),
                                                       ("PDF Files", "*.pdf"),
                                                       ("Word Files", "*.docx")])
    if file_path:
        text = ""
        if file_path.endswith('.txt'):
            with open(file_path, 'r', encoding='utf-8') as file:
                text = file.read()
        elif file_path.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif file_path.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        
        text_entry.delete("1.0", tk.END)  # Clear existing text
        text_entry.insert(tk.END, text)  # Insert the content of the file
        analyze_sentiment(text)  # Analyze the sentiment of the document

def extract_text_from_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as pdf_document:
        for page in pdf_document:
            text += page.get_text()
    return text

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

# Create the main window
window = tk.Tk()
window.title("Sentiment Analysis Tool")

# Create a text box for input
text_entry = tk.Text(window, height=20, width=50)
text_entry.pack()

# Create buttons for analyzing text and opening a document
text_button = tk.Button(window, text="ANALYZE TEXT", command=analyze_text_input)
text_button.pack()

document_button = tk.Button(window, text="OPEN DOCUMENT", command=open_document)
document_button.pack()

# Create a button to clear the text box
clear_button = tk.Button(window, text="CLEAR", command=clear_text)
clear_button.pack()

# Run the application
window.mainloop()
