import pandas as pd
import joblib
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def LoadClassifier():
    model = joblib.load("classifier_model.pkl")
    vectorizer = joblib.load("vectorizer_model.pkl")
    return model, vectorizer

def PredictJuryStatus(docketText, model, vectorizer):
    # Standardize text to match training
    docketText = str(docketText).lower().strip()
    vec = vectorizer.transform([docketText])
    return model.predict(vec)[0]

def Main():
    Tk().withdraw()
    filePath = askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
    
    if not filePath:
        print("❌ No file selected.")
        return

    df = pd.read_excel(filePath)

    if "Docket Text" not in df.columns:
        print("❌ Column 'Docket Text' not found in the Excel file.")
        return

    model, vectorizer = LoadClassifier()

    if "Status" not in df.columns:
        print("ℹ️ 'Status' column not found. Creating it using the trained model...")
        df["Status"] = df["Docket Text"].apply(lambda text: PredictJuryStatus(text, model, vectorizer))
    else:
        print("ℹ️ 'Status' column already exists. Predictions will be stored in 'PredictedStatus'...")
        df["PredictedStatus"] = df["Docket Text"].apply(lambda text: PredictJuryStatus(text, model, vectorizer))

    outputFile = "prediction.xlsx"
    df.to_excel(outputFile, index=False)
    print(f"✅ Done. File saved as: {outputFile}")

if __name__ == "__main__":
    Main()