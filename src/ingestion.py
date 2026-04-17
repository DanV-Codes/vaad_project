import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog
import shutil
from cleaning import load_bank_data  

def pick_and_copy_file(raw_dir):
    """
    Opens a window to select a file, reads its dates, and copies it to the 
    raw directory with a name containing the date range.
    Overwrites the file if it already exists.
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
        print("Reading dates from the file...")
        df_temp = load_bank_data(source_file_path)
        
        dates = pd.to_datetime(df_temp['תאריך'], dayfirst=True, errors='coerce')
        
        if dates.notna().any():
            min_date = dates.min().strftime('%d-%m-%Y')
            max_date = dates.max().strftime('%d-%m-%Y')
            new_filename = f"bank_data_{min_date}_to_{max_date}.xls"
        else:
            new_filename = "bank_data_unknown_dates.xls"
            
        dest_file_path = os.path.join(raw_dir, new_filename)
        
        # מחקנו את מנגנון ההגנה! עכשיו הוא פשוט מעתיק ודורס אוטומטית אם קיים
        shutil.copy2(source_file_path, dest_file_path)
        print(f"✅ File successfully copied (overwritten if existed) to: {os.path.basename(dest_file_path)}")
        
        return dest_file_path

    except Exception as e:
        print(f"❌ Error while trying to rename and copy the file: {e}")
        return None