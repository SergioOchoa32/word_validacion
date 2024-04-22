import os.path
import docx
import nltk

# Función para extraer el texto de un documento Word
def extraer_texto_docx(docx_file):
    if os.path.exists(docx_file):  # Verificar si el archivo existe
        doc = docx.Document(docx_file)
        texto = ""
        for para in doc.paragraphs:
            texto += para.text + "\n"
        return texto
    else:
        print("El documento no existe. Asegurate de que el nombre del documento sea correcto.")
        return None

# Función para contar el número de palabras en el texto
def contar_palabras(texto):
    palabras = texto.split()
    return len(palabras)

# Función para contar el número de líneas de texto en el documento
def contar_lineas(texto):
    lineas = texto.split('\n')
    return len(lineas)

# Función para mostrar palabras de una longitud específica y contar su frecuencia
def mostrar_palabras_cortas(texto, longitud=3):
    palabras = texto.split()
    palabras_cortas = [palabra for palabra in palabras if len(palabra) == longitud]
    frecuencia_palabras_cortas = nltk.FreqDist(palabras_cortas)
    print("Palabras de", longitud, "caracteres:")
    for palabra, frecuencia in frecuencia_palabras_cortas.items():
        print(palabra, "-", frecuencia)

# Nombre del archivo .docx
docx_file = input("Introduce el nombre del archivo .docx: ")

# Extraer texto del documento Word si existe
texto_documento = extraer_texto_docx(docx_file)

if texto_documento:  # Si se extrajo el texto con éxito
    # Guardar el texto extraído en un archivo de texto
    with open("texto_documento.txt", "w", encoding="utf-8") as archivo:
        archivo.write(texto_documento)

    # Contar el número de palabras en el texto
    num_palabras = contar_palabras(texto_documento)
    print("Número de palabras en el documento:", num_palabras)

    # Contar el número de líneas de texto en el documento
    num_lineas = contar_lineas(texto_documento)
    print("Número de líneas de texto en el documento:", num_lineas)

    # Mostrar palabras de 3 o 4 caracteres y contar su frecuencia
    mostrar_palabras_cortas(texto_documento, longitud=3)
    mostrar_palabras_cortas(texto_documento, longitud=4)

    # Cargar el texto del archivo
    archivo_nombre = "texto_documento.txt"
    with open(archivo_nombre, "r", encoding="utf-8") as archivo:
        texto = archivo.read()

    print("----------------------------------------------------------------------")

    # Cargar palabras funcionales en español de NLTK
    nltk.download('stopwords')
    palabras_funcionales = nltk.corpus.stopwords.words("spanish")
    for palabra_funcional in palabras_funcionales:
        print(palabra_funcional)

    print("----------------------------------------------------------------------")

    # Tokenizar el texto y eliminar palabras funcionales
    tokens = nltk.word_tokenize(texto, "spanish")
    tokens_limpios = [token for token in tokens if token.lower() not in palabras_funcionales]

    # Imprimir algunos detalles sobre los tokens
    print(tokens_limpios)
    print("Número total de tokens:", len(tokens))
    print("Número de tokens limpios:", len(tokens_limpios))

    # Crear un objeto Text de NLTK y calcular la distribución de frecuencia
    texto_limpio_nltk = nltk.Text(tokens_limpios)
    distribucion_limpia = nltk.FreqDist(texto_limpio_nltk)

    # Graficar las 40 palabras más comunes
    distribucion_limpia.plot(40)