import os
import pandas as pd
from ingestion import pick_and_copy_file
from cleaning import load_bank_data
from processing import process_bank_transactions, enrich_with_apartments

if __name__ == "__main__":
    # 1. הגדרת התיקיות והנתיבים
    raw_directory = r"C:\Users\danie\PyProjects\vaad_project\data\raw"
    processed_directory = r"C:\Users\danie\PyProjects\vaad_project\data\processed"
    apartments_csv_path = r"C:\Users\danie\PyProjects\vaad_project\data\reference\apartments.csv"
    categories_csv_path = r"C:\Users\danie\PyProjects\vaad_project\data\reference\categories.csv" # הנתיב החדש לקטגוריות!
    
    print("Starting Vaad Project Pipeline...")
    
    # שלב א': בחירת קובץ מההורדות והעתקה חכמה
    copied_file_path = pick_and_copy_file(raw_directory)
    
    if copied_file_path:
        try:
            # שלב ב': טעינה וניקוי
            raw_df = load_bank_data(copied_file_path)
            print("✅ Module 1: Data loaded and cleaned.")
            
            # שלב ג': עיבוד מתקדם (כולל קריאת קובץ הקטגוריות החדש!)
            processed_df = process_bank_transactions(raw_df, categories_csv_path)
            print("✅ Module 2: Data processed and categorized using external rules.")
            
            # שלב ד': הצלבת דירות להכנסות
            processed_df = enrich_with_apartments(processed_df, apartments_csv_path)
            print("✅ Module 3: Data enriched with apartment numbers.")
            
            # שלב ה': חילוץ תאריכים לשם הקובץ
            dates = pd.to_datetime(processed_df['Date'], dayfirst=True, errors='coerce')
            if dates.notna().any():
                min_date = dates.min().strftime('%d-%m-%Y')
                max_date = dates.max().strftime('%d-%m-%Y')
                date_suffix = f"{min_date}_to_{max_date}"
            else:
                date_suffix = "unknown_dates"
                
            credit_csv_path = os.path.join(processed_directory, f"credit_transactions_{date_suffix}.csv")
            debit_csv_path = os.path.join(processed_directory, f"debit_transactions_{date_suffix}.csv")
            
            # שלב ו': פיצול הנתונים וסידור סופי
            
            # --- טיפול בטבלת זכות (הכנסות) ---
            credit_df = processed_df[processed_df['Amount'] > 0].copy()
            # העברת עמודת מס' דירה מיד אחרי קטגוריה בקובץ זכות
            if 'Apartment_Number' in credit_df.columns and 'Category' in credit_df.columns:
                cols = list(credit_df.columns)
                cols.insert(cols.index('Category') + 1, cols.pop(cols.index('Apartment_Number')))
                credit_df = credit_df[cols]
            
            # --- טיפול בטבלת חובה (הוצאות) ---
            debit_df = processed_df[processed_df['Amount'] < 0].copy()
            # מחיקת עמודת מספר דירה מטבלת ההוצאות
            if 'Apartment_Number' in debit_df.columns:
                debit_df = debit_df.drop(columns=['Apartment_Number'])
                
            # שלב ז': שמירת הקבצים המפוצלים
            credit_df.to_csv(credit_csv_path, index=False, encoding='utf-8-sig')
            debit_df.to_csv(debit_csv_path, index=False, encoding='utf-8-sig')
            
            print(f"\n🎉 Success! Files split and saved:")
            print(f"💰 Credit file: {os.path.basename(credit_csv_path)} ({len(credit_df)} rows)")
            print(f"💸 Debit file: {os.path.basename(debit_csv_path)} ({len(debit_df)} rows)")
            
        except Exception as e:
            print(f"\n❌ An unexpected error occurred in main pipeline: {e}")
    else:
        print("\nPipeline stopped because no file was selected.")