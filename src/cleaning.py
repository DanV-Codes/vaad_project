import pandas as pd

def load_bank_data(file_path):
    """
    Loads bank HTML data, dynamically finds headers, and cleans the dataframe.
    """
    tables = pd.read_html(file_path, encoding='utf-8')
    df = max(tables, key=len)
    
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

    df.columns = df.columns.astype(str).str.strip()
    
    if 'הערה' in df.columns:
        df = df[~df['הערה'].astype(str).str.contains("תנועות היום", na=False)]
        
    df = df.dropna(how='all')
    
    return df