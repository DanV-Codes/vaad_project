import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import os
import tkinter as tk
from tkinter import scrolledtext

def show_summary_popup(month_num, updated_count, already_filled_apts, newly_updated_apts, unidentified_df, amount_exceptions_df):
    """
    חלון קופץ ידידותי למשתמש עם סיכום הרצת הנתונים.
    """
    root = tk.Tk()
    root.title(f"דוח סיכום גבייה - חודש {month_num}")
    root.geometry("850x650")
    root.configure(bg="#f4f4f4")

    def close_app():
        root.quit()
        root.destroy()

    title_lbl = tk.Label(root, text=f"📊 סיכום הרצת נתונים באקסל - חודש {month_num}", font=("Arial", 18, "bold"), bg="#f4f4f4")
    title_lbl.pack(pady=15)

    stats_frame = tk.Frame(root, bg="#f4f4f4")
    stats_frame.pack(fill=tk.X, padx=20)

    success_text = f"✅ עודכנו בהצלחה: {updated_count} דירות"
    if newly_updated_apts:
        success_text += f" ({sorted(newly_updated_apts)})"
    tk.Label(stats_frame, text=success_text, font=("Arial", 14), fg="green", bg="#f4f4f4").pack(anchor="e", pady=2)

    if already_filled_apts:
        tk.Label(stats_frame, text=f"ℹ️ דולגו (כבר שילמו באקסל): {len(already_filled_apts)} דירות ({sorted(already_filled_apts)})", font=("Arial", 14), fg="#b8860b", bg="#f4f4f4").pack(anchor="e", pady=2)

    text_area = scrolledtext.ScrolledText(root, wrap=tk.NONE, font=("Courier", 11), bg="#2b2b2b", fg="#a9b7c6")
    text_area.pack(pady=15, padx=20, fill=tk.BOTH, expand=True)

    if not unidentified_df.empty:
        text_area.insert(tk.END, f"❌ רשימה שחורה (דירה לא זוהתה): {len(unidentified_df)} שורות\n")
        text_area.insert(tk.END, "="*70 + "\n")
        display_cols = [c for c in ['Date', 'Amount', 'Original_Desc', 'Entity_Name'] if c in unidentified_df.columns]
        text_area.insert(tk.END, unidentified_df[display_cols].to_string(index=False) + "\n\n\n")

    if not amount_exceptions_df.empty:
        text_area.insert(tk.END, f"⚠️ חריגי תשלום (סכום לא תואם לחוקים): {len(amount_exceptions_df)} שורות\n")
        text_area.insert(tk.END, "="*70 + "\n")
        display_cols = [c for c in ['Apartment_Number', 'Amount', 'Date', 'Entity_Name'] if c in amount_exceptions_df.columns]
        text_area.insert(tk.END, amount_exceptions_df[display_cols].to_string(index=False) + "\n")

    if unidentified_df.empty and amount_exceptions_df.empty:
        text_area.insert(tk.END, "\n🎉 אין חריגים או שגיאות! הכל עבר חלק.\n")

    text_area.config(state=tk.DISABLED)

    close_btn = tk.Button(root, text="סגור דוח והמשך", command=close_app, font=("Arial", 14, "bold"), bg="#4CAF50", fg="white", width=20, pady=10)
    close_btn.pack(pady=15)

    root.mainloop()


def update_master_excel(df, master_path):
    if df.empty:
        print("⚠️ No transactions to process.")
        return

    # סינון דירות לא מזוהות (רשימה שחורה)
    numeric_apt = pd.to_numeric(df['Apartment_Number'], errors='coerce')
    unidentified_mask = numeric_apt.isna()
    unidentified_df = df[unidentified_mask].copy()
    
    identified_df = df[~unidentified_mask].copy()
    identified_df['Apt_Numeric'] = numeric_apt[~unidentified_mask]

    print(f"Opening Master Excel: {os.path.basename(master_path)}...")
    try:
        wb = openpyxl.load_workbook(master_path)
    except Exception as e:
        print(f"❌ Error loading Excel file: {e}")
        return
        
    sheet_name = "גביה 2026"
    if sheet_name not in wb.sheetnames:
        print(f"❌ Error: Sheet '{sheet_name}' not found!")
        return
        
    ws = wb[sheet_name]
    
    valid_dates = pd.to_datetime(df['Date'], dayfirst=True, errors='coerce').dropna()
    if valid_dates.empty:
        print("❌ Error: Could not find valid dates.")
        return
        
    month_num = valid_dates.iloc[0].month
    target_col = month_num + 4 

    # ==========================================
    # לוגיקת העדכון והחרגות (דירות 12 ו-40)
    # ==========================================
    updated_count = 0
    newly_updated_apts = []
    already_filled_apts = []
    exceptions_idx = []

    for idx, row in identified_df.iterrows():
        apt_num = int(row['Apt_Numeric'])
        amount = row['Amount']
        target_row = apt_num + 1
        
        current_val = ws.cell(row=target_row, column=target_col).value
        curr_str = str(current_val).strip() if current_val is not None else ""
        if curr_str.endswith(".0"): curr_str = curr_str[:-2]

        if apt_num in [12, 40]:
            if amount == 320:
                if curr_str == "30":
                    ws.cell(row=target_row, column=target_col).value = 350 
                    updated_count += 1
                    newly_updated_apts.append(apt_num)
                elif curr_str == "350":
                    already_filled_apts.append(apt_num)
                else:
                    exceptions_idx.append(idx) 
                    
            elif amount == 350: 
                if curr_str == "" or curr_str == "30":
                    ws.cell(row=target_row, column=target_col).value = 350
                    updated_count += 1
                    newly_updated_apts.append(apt_num)
                elif curr_str == "350":
                    already_filled_apts.append(apt_num)
                else:
                    exceptions_idx.append(idx)
            else:
                exceptions_idx.append(idx)

        else:
            if amount == 350:
                if curr_str == "":
                    ws.cell(row=target_row, column=target_col).value = 350
                    updated_count += 1
                    newly_updated_apts.append(apt_num)
                elif curr_str == "350":
                    already_filled_apts.append(apt_num)
                else:
                    exceptions_idx.append(idx) 
            else:
                exceptions_idx.append(idx) 

    amount_exceptions_df = identified_df.loc[exceptions_idx].copy()

    # ==========================================
    # יצירת ושמירת קובץ ה-CSV של החריגים
    # ==========================================
    # נוסיף עמודה שמסבירה מה סוג החריגה כדי שיהיה קל לסנן בדאשבורד
    unidentified_save = unidentified_df.copy()
    if not unidentified_save.empty:
        unidentified_save['Exception_Reason'] = 'דירה לא זוהתה'
        
    amount_exceptions_save = amount_exceptions_df.copy()
    if not amount_exceptions_save.empty:
        amount_exceptions_save['Exception_Reason'] = 'סכום חריג / תא מלא'
        
    # איחוד כל החריגים לטבלה אחת
    all_exceptions = pd.concat([unidentified_save, amount_exceptions_save], ignore_index=True)
    
    if not all_exceptions.empty:
        # ניווט לתיקיית data/exceptions (בהנחה ש-master_path הוא בשורש הפרויקט)
        project_root = os.path.dirname(master_path)
        exceptions_dir = os.path.join(project_root, "data", "exceptions")
        
        # יצירת התיקייה אם היא לא קיימת
        os.makedirs(exceptions_dir, exist_ok=True) 
        
        exceptions_file = os.path.join(exceptions_dir, f"exceptions_month_{month_num}.csv")
        all_exceptions.to_csv(exceptions_file, index=False, encoding='utf-8-sig')
        print(f"\n📁 קובץ חריגים נוצר ונשמר בהצלחה בנתיב: {exceptions_file}")


    # יצירת לשונית גיבוי גולמית באקסל
    backup_name = f"פירוט_גביה_{month_num}"
    if backup_name in wb.sheetnames: del wb[backup_name] 
    ws_backup = wb.create_sheet(title=backup_name)
    for r in dataframe_to_rows(df, index=False, header=True): ws_backup.append(r)
    
    try:
        wb.save(master_path)
    except PermissionError:
        print(f"\n❌ שגיאה: הקובץ '{os.path.basename(master_path)}' פתוח באקסל. סגור אותו ונסה שוב.")
        return

    # הקפצת דוח הסיכום
    show_summary_popup(
        month_num=month_num,
        updated_count=updated_count,
        already_filled_apts=already_filled_apts,
        newly_updated_apts=newly_updated_apts,
        unidentified_df=unidentified_df,
        amount_exceptions_df=amount_exceptions_df
    )