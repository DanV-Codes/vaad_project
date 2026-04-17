import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import shutil

def load_bank_data(file_path):
    """
    Loads bank HTML data, dynamically finds headers, and cleans the dataframe.
    """
    tables = pd.read_html(file_path, encoding='utf-8')
    df = max(tables, key=len)
    
    # 1. Dynamically find the header row
    header_row_index = -1
    for index, row in df.iterrows():
        row_str = ' '.join(row.dropna().astype(str))
        if 'תאריך' in row_str and 'תיאור' in row_str:
            header_row_index = index
            break
            
    if header_row_index != -1:
        df.columns = df.iloc[header_row_index]
        df = df.iloc[header_row_index + 1:].reset_index(drop=True)
    else:
        print("⚠️ Warning: Could not automatically find the header row.")

    # 2. Clean up column names
    df.columns = df.columns.astype(str).str.strip()
    
    # 3. Filter out "תנועות היום" from the Remark column (הערה)
    if 'הערה' in df.columns:
        df = df[~df['הערה'].astype(str).str.contains("תנועות היום", na=False)]
        
    # 4. Drop any completely empty rows
    df = df.dropna(how='all')
    
    return df

def pick_and_copy_file(raw_dir):
    """
    Opens a window to select a file, reads its dates, and copies it to the 
    raw directory with a name containing the date range.
    Returns the path to the newly copied file.
    """
    root = tk.Tk()
    root.withdraw()
    
    source_file_path = filedialog.askopenfilename(
        title="בחר את קובץ הבנק שהורדת",
        filetypes=[("Excel files", "*.xls"), ("All files", "*.*")]
    )
    
    if not source_file_path:
        print("⚠️ No file was selected.")
        return None
        
    print(f"Selected file: {source_file_path}")
    
    try:
        # 1. טעינה זמנית של הנתונים כדי להציץ בתאריכים
        print("Reading dates from the file...")
        df_temp = load_bank_data(source_file_path)
        
        # 2. המרת עמודת התאריך לפורמט תאריך אמיתי
        dates = pd.to_datetime(df_temp['תאריך'], dayfirst=True, errors='coerce')
        
        # 3. חילוץ תאריך התחלה וסיום ובניית השם החדש
        if dates.notna().any():
            min_date = dates.min().strftime('%d-%m-%Y')
            max_date = dates.max().strftime('%d-%m-%Y')
            new_filename = f"bank_data_{min_date}_to_{max_date}.xls"
        else:
            new_filename = "bank_data_unknown_dates.xls"
            
        dest_file_path = os.path.join(raw_dir, new_filename)
        
        # 4. הגנה מפני דריסה: אם כבר יש קובץ כזה בדיוק, נוסיף לו מספר
        counter = 1
        while os.path.exists(dest_file_path):
            name_without_ext = new_filename.replace('.xls', '')
            dest_file_path = os.path.join(raw_dir, f"{name_without_ext}_{counter}.xls")
            counter += 1
            
        # 5. העתקת הקובץ עם השם החכם!
        shutil.copy2(source_file_path, dest_file_path)
        print(f"✅ File successfully copied and renamed to: {os.path.basename(dest_file_path)}")
        
        return dest_file_path

    except Exception as e:
        print(f"❌ Error while trying to rename and copy the file: {e}")
        return None
    """
    Opens a window to select a file, and copies it to the raw directory.
    Returns the path to the newly copied file.
    """
    # Create the window but keep it hidden
    root = tk.Tk()
    root.withdraw()
    
    # Open the file picker
    source_file_path = filedialog.askopenfilename(
        title="בחר את קובץ הבנק שהורדת",
        filetypes=[("Excel files", "*.xls"), ("All files", "*.*")]
    )
    
    if not source_file_path:
        print("⚠️ No file was selected.")
        return None
        
    print(f"Selected file: {source_file_path}")
    
    # Extract just the file name (e.g., "bank_export.xls")
    file_name = os.path.basename(source_file_path)
    
    # Create the exact path where we want to save it in the raw folder
    dest_file_path = os.path.join(raw_dir, file_name)
    
    # Copy the file!
    shutil.copy2(source_file_path, dest_file_path)
    print(f"✅ File successfully copied to: {dest_file_path}")
    
    return dest_file_path


# --- Execution Block ---
if __name__ == "__main__":
    # הגדרת התיקייה שאליה אנחנו רוצים להעתיק את הקובץ (ולא קובץ ספציפי)
    raw_directory = r"C:\Users\danie\PyProjects\vaad_project\data\raw"
    
    # הנתיב שבו נשמור את קובץ הבדיקה 
    test_csv_path = r"C:\Users\danie\PyProjects\vaad_project\output\test.csv"
    
    print("Opening file picker...")
    
    # 1. הפעלת החלון הקופץ והעתקת הקובץ
    copied_file_path = pick_and_copy_file(raw_directory)
    
    # רק אם נבחר קובץ (ולא לחצו Cancel), נמשיך הלאה
    if copied_file_path:
        print(f"\nAttempting to load data from: {copied_file_path}")
        
        try:
            # 2. טעינת הנתונים מהקובץ החדש שהועתק
            raw_df = load_bank_data(copied_file_path)
            print("\n✅ File loaded successfully!")
            
            # --- יצירת קובץ ה-CSV לבדיקה ---
            print(f"\nSaving a 3-row test file to: {test_csv_path}")
            raw_df.head(3).to_csv(test_csv_path, index=False, encoding='utf-8-sig')
            
            print("✅ Test CSV created! Go check output/test.csv in VS Code or Excel.")
            
        except Exception as e:
            print(f"\n❌ An unexpected error occurred: {e}")