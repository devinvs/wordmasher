import pandas as pd
from docx import Document
from python_docx_replace import docx_replace
import re

def wordmash(replace_df):
    for (i, row) in replace_df.iterrows():
        doc = Document("template.docx")
        for col in replace_df.columns:
            search_replace(doc, col, row[col])

        doc.save(f"project_{i}.docx")
        break

def search_replace(doc, search, replace):
    regex = re.compile(search)
    docx_replace(doc, **{search: replace})

if __name__ == "__main__":
    db = pd.read_excel("database.xlsx")
    wordmash(db)
