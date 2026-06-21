import re
import unicodedata
from pathlib import Path

import nltk
import pandas as pd
from nltk.tokenize import sent_tokenize

# Download once
nltk.download('punkt')

# ==========================================
# FOLDERS
# ==========================================

DATA_DIR = Path("data")
OUTPUT_FILE = "Dr.Jangala.xlsx"

ENGLISH_FILE = DATA_DIR / "dr.jangala.txt"
MANIPURI_FILE = DATA_DIR / "Manipuri.dr.jangala.txt"


# ==========================================
# ENGLISH CLEANING
# ==========================================

def clean_english(text):

    text = re.sub(r'^\{"format":"html","output":"', '', text)
    text = re.sub(r'"\}$', '', text)

    text = text.replace("\\n", " ")
    text = text.replace("\\r", " ")
    text = text.replace("\\t", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ==========================================
# MANIPURI CLEANING
# ==========================================

def clean_manipuri(text):

    text = unicodedata.normalize("NFC", text)

    text = text.replace("\ufeff", "")
    text = text.replace("\\n", " ")
    text = text.replace("\\r", " ")
    text = text.replace("\\t", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ==========================================
# MANIPURI TOKENIZER
# ==========================================

def manipuri_tokenize(text):

    text = clean_manipuri(text)

    sentences = re.findall(r'[^꯫]+꯫', text)

    return [
        sentence.strip()
        for sentence in sentences
        if sentence.strip()
    ]


# ==========================================
# CHECK FILES
# ==========================================

if not DATA_DIR.exists():
    raise FileNotFoundError(
        "Create a folder named 'data'."
    )

if not ENGLISH_FILE.exists():
    raise FileNotFoundError(
        "english.txt not found inside data folder."
    )

if not MANIPURI_FILE.exists():
    raise FileNotFoundError(
        "manipuri.txt not found inside data folder."
    )


# ==========================================
# READ ENGLISH
# ==========================================

with open(
    ENGLISH_FILE,
    "r",
    encoding="utf-8"
) as f:

    english_text = f.read()

english_text = clean_english(english_text)

english_sentences = sent_tokenize(english_text)

english_sentences = [
    s.strip()
    for s in english_sentences
    if len(s.strip()) > 3
]


# ==========================================
# READ MANIPURI
# ==========================================

with open(
    MANIPURI_FILE,
    "r",
    encoding="utf-8"
) as f:

    manipuri_text = f.read()

manipuri_sentences = manipuri_tokenize(
    manipuri_text
)


# ==========================================
# MAKE SAME LENGTH
# ==========================================

max_len = max(
    len(english_sentences),
    len(manipuri_sentences)
)

english_sentences.extend(
    [""] * (max_len - len(english_sentences))
)

manipuri_sentences.extend(
    [""] * (max_len - len(manipuri_sentences))
)


# ==========================================
# CREATE DATAFRAME
# ==========================================

df = pd.DataFrame({
    "English": english_sentences,
    "Manipuri": manipuri_sentences
})


# ==========================================
# SAVE EXCEL
# ==========================================

df.to_excel(
    OUTPUT_FILE,
    index=False,
    header=False

)

print("English sentences :", len(english_sentences))
print("Manipuri sentences:", len(manipuri_sentences))
print("Saved:", OUTPUT_FILE)