import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import matplotlib.pyplot as plt
import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from collections import Counter
import string
import re

# Descargar recursos de nltk si no están descargados
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('averaged_perceptron_tagger')

# Obtener conjunciones y preposiciones
stop_words = set(stopwords.words('spanish'))
stop_words.update({'cc', 'cs', 'in', 'rb', 'rbr', 'rbs'})

def read_docx(file_path):
    try:
        doc = Document(file_path)
        text = ""
        for para in doc.paragraphs:
            text += para.text + " "
        return text
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")
        return None

def clean_text(text):
    # Remover todo excepto letras y espacios
    text = re.sub(r'[^a-záéíóúüñ\s]', '', text, flags=re.IGNORECASE)
    return text

def analyze_texts(text1, text2):
    # Limpiar los textos
    text1 = clean_text(text1)
    text2 = clean_text(text2)

    # Tokenizar el texto
    tokens1 = word_tokenize(text1.lower())
    tokens2 = word_tokenize(text2.lower())

    # Contar las ocurrencias de cada palabra en cada texto
    counter1 = Counter(tokens1)
    counter2 = Counter(tokens2)

    # Filtrar palabras comunes
    common_words = set(counter1.keys()).intersection(set(counter2.keys()))

    # Contar las ocurrencias de cada palabra común en ambos textos
    common_word_counts = {word: counter1[word] + counter2[word] for word in common_words}

    # Filtrar palabras según las partes del discurso
    filtered_common_words = {}
    for word, count in common_word_counts.items():
        if word not in stop_words:  # Excluir conjunciones, preposiciones y adverbios
            pos = nltk.pos_tag([word])[0][1]  # Obtener parte del discurso
            if pos.startswith('N') or pos.startswith('J'):  # Sustantivo o Adjetivo
                filtered_common_words[word] = count

    return filtered_common_words

def load_file(text_area):
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    if file_path:
        text = read_docx(file_path)
        if text:
            text_area.delete("1.0", tk.END)  # Limpiar el área de texto antes de cargar el nuevo contenido
            text_area.insert(tk.END, text)
    else:
        messagebox.showinfo("Información", "No se seleccionó ningún archivo")

def analyze_and_visualize():
    text1 = text_area1.get("1.0", "end-1c")
    text2 = text_area2.get("1.0", "end-1c")

    if not text1 or not text2:
        messagebox.showwarning("Advertencia", "Por favor, carga ambos archivos antes de analizar")
        return

    # Obtener palabras comunes filtradas por partes del discurso
    common_words_count = analyze_texts(text1, text2)

    messagebox.showinfo("Resultado", f"Número de palabras comunes: {sum(common_words_count.values())}")

    # Obtener las 50 palabras más comunes
    common_words_count = dict(sorted(common_words_count.items(), key=lambda item: item[1], reverse=True)[:50])

    # Crear gráfico
    labels = list(common_words_count.keys())
    values = list(common_words_count.values())
    plt.bar(labels, values)
    plt.xlabel('Palabras')
    plt.ylabel('Frecuencia exacta')
    plt.title('Palabras más comunes')
    plt.xticks(rotation=60, ha='right')  # Rotar los labels en un ángulo de 60 grados

    # Mostrar la frecuencia exacta en el gráfico
    for i in range(len(labels)):
        plt.text(i, values[i], str(values[i]), ha='center', va='bottom')

    plt.tight_layout()
    plt.show()

    return common_words_count

# Crear ventana principal
root = tk.Tk()
root.title("Comparación de Textos")

# Crear marcos
frame1 = tk.Frame(root)
frame1.pack(side=tk.TOP)
frame2 = tk.Frame(root)
frame2.pack(side=tk.BOTTOM)

# Crear áreas de texto
text_area1 = tk.Text(frame1, height=10, width=50)
text_area1.pack(side=tk.LEFT)
text_area2 = tk.Text(frame1, height=10, width=50)
text_area2.pack(side=tk.RIGHT)

# Botón para cargar archivos
load_button1 = tk.Button(frame2, text="Cargar Archivo 1", command=lambda: load_file(text_area1))
load_button1.pack(side=tk.LEFT)
load_button2 = tk.Button(frame2, text="Cargar Archivo 2", command=lambda: load_file(text_area2))
load_button2.pack(side=tk.LEFT)

# Botón para analizar y visualizar
analyze_button = tk.Button(frame2, text="Analizar y Visualizar", command=analyze_and_visualize)
analyze_button.pack(side=tk.RIGHT)

# Ejecutar la aplicación
root.mainloop()
