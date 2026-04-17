import pandas as pd
import os

def process_bank_transactions(df, categories_csv_path=None):
    """
    Processes the raw bank dataframe: calculates standard amount, 
    categorizes transactions using an external CSV file, 
    and extracts entity names.
    """
    processed_df = pd.DataFrame()
    processed_df['Date'] = df['תאריך']
    
    credit_col = 'בזכות' if 'בזכות' in df.columns else 'זכות'
    debit_col = 'בחובה' if 'בחובה' in df.columns else 'חובה'
    
    credit = pd.to_numeric(df[credit_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    debit = pd.to_numeric(df[debit_col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    processed_df['Amount'] = credit - debit
    
    # --- קריאת חוקי הקטגוריות מהקובץ החיצוני ---
    rules = []
    if categories_csv_path and os.path.exists(categories_csv_path):
        rules_df = pd.read_csv(categories_csv_path, encoding='utf-8')
        # הופכים את הטבלה לרשימה של מילונים שקל לחפש בה
        rules = rules_df.to_dict('records')
    else:
        print("⚠️ Warning: Categories CSV not found. Proceeding without rules.")

    categories = []
    entity_names = []
    
    for index, row in df.iterrows():
        desc = str(row.get('תיאור', ''))
        ext_desc = str(row.get('תאור מורחב', '')) 
        
        category = 'לא ידוע'
        name = 'לא ידוע'
        
        # 1. לולאה שעוברת על כל החוקים מקובץ ה-CSV
        for rule in rules:
            keyword = str(rule.get('Keyword', '')).strip()
            column_to_search = str(rule.get('Column_To_Search', '')).strip()
            
            # אם החוק אומר לחפש בתיאור הרגיל
            if column_to_search == 'Original_Desc' and keyword in desc:
                category = str(rule.get('Category', '')).strip()
                name = str(rule.get('Entity_Name', '')).strip()
                break
                
            # אם החוק אומר לחפש בתיאור המורחב
            elif column_to_search == 'Original_Ext_Desc' and keyword in ext_desc:
                category = str(rule.get('Category', '')).strip()
                name = str(rule.get('Entity_Name', '')).strip()
                break
                
        # 2. טיפול בהעברות בנקאיות (רק אם לא מצאנו קטגוריה אחרת ב-CSV)
        if category == 'לא ידוע':
            if 'העברה מאת' in ext_desc:
                category = 'הכנסה מדיירים'
                try:
                    name = ext_desc.split(',')[0].split(':')[1].strip()    
                except IndexError:
                    name = 'שגיאה בחילוץ שם'
                    
            elif 'העברה אל' in ext_desc:
                category = 'הוצאה (העברה)'
                try:
                    name = ext_desc.split(',')[0].split(':')[1].strip()
                except IndexError:
                    name = 'שגיאה בחילוץ שם'
                    
        categories.append(category)
        entity_names.append(name)
        
    processed_df['Category'] = categories
    processed_df['Entity_Name'] = entity_names
    processed_df['Original_Desc'] = df['תיאור']
    processed_df['Original_Ext_Desc'] = df['תאור מורחב']
    
    return processed_df

def enrich_with_apartments(df, apartments_csv_path):
    """
    Reads the apartments.csv file, handles the '|' separated names,
    and smartly searches for these names within the transaction text.
    Ensures apartment numbers are integers.
    """
    if not os.path.exists(apartments_csv_path):
        print(f"⚠️ Warning: Apartments file not found at {apartments_csv_path}")
        df['Apartment_Number'] = 'חסר קובץ'
        return df
        
    apt_df = pd.read_csv(apartments_csv_path, encoding='utf-8')
    
    # במקום מילון, ניצור רשימה של זוגות: (שם, מספר דירה)
    search_list = []
    
    for index, row in apt_df.iterrows():
        # המרת מספר הדירה למספר שלם (Integer) כדי להימנע מ- .0
        try:
            apt_num = int(row['apartment_number'])
        except (ValueError, TypeError):
            # במקרה שמישהו כתב טקסט במקום מספר בקובץ הדירות
            apt_num = str(row['apartment_number']) 
            
        payer_names_str = str(row['payer_names'])
        names = payer_names_str.split('|')
        
        for name in names:
            clean_name = name.strip()
            if clean_name:
                search_list.append((clean_name, apt_num))
                
    # פונקציית חיפוש חכמה שתעבור על כל שורה בטראנזקציות
    def find_apartment_in_text(row):
        # נחבר את התיאור המורחב והשם שחילצנו לטקסט אחד ארוך
        text_to_search = str(row.get('Entity_Name', '')) + " " + str(row.get('Original_Ext_Desc', ''))
        
        # נחפש אם אחד מהשמות ברשימה שלנו נמצא בתוך הטקסט
        for name, apt in search_list:
            if name in text_to_search:
                return apt
                
        return 'לא זוהה'
        
    # הפעלת החיפוש החכם על כל השורות שלנו
    df['Apartment_Number'] = df.apply(find_apartment_in_text, axis=1)
    
    # סידור העמודות (נכניס את מספר הדירה אחרי השם)
    if 'Apartment_Number' in df.columns and 'Entity_Name' in df.columns:
        cols = list(df.columns)
        cols.insert(cols.index('Entity_Name') + 1, cols.pop(cols.index('Apartment_Number')))
        df = df[cols]
        
    return df