import pandas as pd
from docx import Document
from python_docx_replace import docx_replace

def wordmash(replace_df):
    for (i, row) in replace_df.iterrows():
        doc = Document("template.docx")
        for col in replace_df.columns:
            docx_replace(doc, **{col: row[col]})

        doc.save(f"project_{i+1}.docx")

if __name__ == "__main__":
    db = pd.read_excel("database.xlsx")
    wordmash(db)
