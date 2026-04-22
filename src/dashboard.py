import streamlit as st
import pandas as pd
import os
import openpyxl

# הגדרות עמוד
st.set_page_config(page_title='ניהול ועד - האגמית 7', layout='wide', page_icon='🏢')

# --- קבועים ---
DEBT_LIMIT = 350
COMMITTEE_APTS = [19, 33, 26, 38]
TOTAL_PAYING_APTS = 60 - len(COMMITTEE_APTS)
MONTHS_HEB = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר']
MASTER_FILE = 'האגמית7_כספים_2026.xlsx'
EXCEPTIONS_DIR = os.path.join('data', 'exceptions')

st.title('📊 מערכת ניהול גבייה - האגמית 7')

# --- פונקציות טעינה ושמירה חכמה ---
def load_excel_data():
    try:
        # טעינה של הנתונים בלבד ללא עיצובים לתוך DataFrame
        df = pd.read_excel(MASTER_FILE, sheet_name='גביה 2026', usecols='B,E:P', nrows=60)
        df.columns = ['דירה'] + MONTHS_HEB
        for month in MONTHS_HEB:
            df[month] = pd.to_numeric(df[month], errors='coerce').fillna(0).astype(int)
        df['דירה'] = df['דירה'].astype(int)
        return df
    except Exception as e:
        st.error(f"שגיאה בקריאת הקובץ: {e}")
        return None

def save_to_excel_with_format(updated_rows):
    """שומר את השינויים תוך שמירה על פורמט ונוסחאות באקסל"""
    try:
        # טעינת הקובץ הקיים בעזרת openpyxl (שומר על עיצוב)
        wb = openpyxl.load_workbook(MASTER_FILE)
        ws = wb['גביה 2026']
        
        # מיפוי עמודות: עמודה B היא 2, עמודות E-P הן 5-16
        month_to_col = {name: 5 + i for i, name in enumerate(MONTHS_HEB)}
        
        # מעבר על כל הדירות שעודכנו
        for _, row in updated_rows.iterrows():
            apt_num = int(row['דירה'])
            
            # חיפוש השורה הנכונה באקסל (מחפשים בעמודה B, החל משורה 2)
            target_row = None
            for r in range(2, 65): # סריקה של הטווח הרלוונטי
                if ws.cell(row=r, column=2).value == apt_num:
                    target_row = r
                    break
            
            if target_row:
                # עדכון הערכים של החודשים בלבד
                for month_name, col_idx in month_to_col.items():
                    # כותבים את הערך החדש כתא מספר
                    ws.cell(row=target_row, column=col_idx).value = int(row[month_name])
        
        wb.save(MASTER_FILE)
        return True
    except Exception as e:
        st.error(f"שגיאה בשמירה לאקסל: {e}")
        return False

# אתחול Session State
if 'master_df' not in st.session_state or st.session_state.master_df is None:
    st.session_state.master_df = load_excel_data()

# הגנה מפני קריסה אם הקובץ לא נטען
if st.session_state.master_df is None:
    st.stop()

# --- תפריט צדדי ---
st.sidebar.header('⚙️ הגדרות')
selected_month_name = st.sidebar.selectbox('בחר חודש לעבודה:', MONTHS_HEB, index=2)
selected_month_idx = MONTHS_HEB.index(selected_month_name)
report_months = st.sidebar.multiselect('כלול בהודעת הווטסאפ:', MONTHS_HEB, default=[selected_month_name])

# --- לוגיקת נתונים ---
paying_df = st.session_state.master_df[~st.session_state.master_df['דירה'].isin(COMMITTEE_APTS)].copy()

# --- 1. הודעת ווטסאפ ---
st.subheader('📱 הודעה מוכנה לווטסאפ')
debt_lines = []
for _, row in paying_df.iterrows():
    apt_debts = []
    for month in report_months:
        if row[month] < DEBT_LIMIT:
            apt_debts.append(f'{month} (חוב {int(DEBT_LIMIT - row[month])})')
    if apt_debts:
        debt_lines.append(f"• דירה {int(row['דירה'])} - {', '.join(apt_debts)}")

st.text_area('העתק מכאן:', value="היי, להלן רשימת החובות:\n" + "\n".join(debt_lines), height=120)

# --- 2. רשימת חריגים (CSV) ---
st.markdown('---')
st.subheader(f'⚠️ חריגים - {selected_month_name}')
exc_path = os.path.join(EXCEPTIONS_DIR, f'exceptions_month_{selected_month_idx+1}.csv')

if os.path.exists(exc_path):
    exc_df = pd.read_csv(exc_path)
    edited_exc = st.data_editor(exc_df, num_rows="dynamic", use_container_width=True, key="exc_ed", hide_index=True)
    if st.button('💾 שמור שינויים בחריגים'):
        edited_exc.to_csv(exc_path, index=False)
        st.success("קובץ החריגים עודכן.")
        st.rerun()
else:
    st.info(f"לא נמצא קובץ חריגים לחודש {selected_month_name}.")

# --- 3. טבלת גבייה (Editable Excel) ---
st.markdown('---')
st.subheader('📅 עריכת נתוני גבייה (אקסל מאסטר)')

# הכנת שורת סיכום
stats = {'דירה': 'אחוז גבייה %'}
for month in MONTHS_HEB:
    paid_count = (paying_df[month] >= DEBT_LIMIT).sum()
    stats[month] = round((paid_count / TOTAL_PAYING_APTS) * 100, 1)

# עיבוד לתצוגה
display_cols = MONTHS_HEB[::-1] + ['דירה']
view_df = paying_df.copy()
view_df['דירה'] = view_df['דירה'].astype(str)
final_display_df = pd.concat([view_df, pd.DataFrame([stats])], ignore_index=True)

# פונקציית עיצוב (ללא אפסים מיותרים)
def format_func(val):
    if isinstance(val, float): return f"{val:.1f}%"
    try: return f"{int(float(val))}"
    except: return str(val)

# עורך הטבלה
edited_table = st.data_editor(
    final_display_df[display_cols],
    use_container_width=True,
    height=500,
    hide_index=True,
    key="master_ed"
)

if st.button('💾 שמור שינויים באקסל המאסטר'):
    # מחלצים רק את השורות של הדירות (בלי שורת האחוזים)
    rows_to_save = edited_table.iloc[:-1].copy()
    rows_to_save['דירה'] = rows_to_save['דירה'].astype(int)
    
    if save_to_excel_with_format(rows_to_save):
        st.success("הנתונים נשמרו! הפורמט והנוסחאות באקסל נשמרו.")
        st.session_state.master_df = load_excel_data() # טעינה מחדש לזיכרון
        st.rerun()