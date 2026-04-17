import streamlit as st
import pandas as pd

# הגדרת סך הדירות בבניין (למשל מ-1 עד 20)
TOTAL_APARTMENTS = list(range(1, 21))

# הגדרות עיצוב לדף
st.set_page_config(page_title="דאשבורד ועד בית", layout="wide")
st.title("סטטוס תשלומי ועד בית 🏢")

# --- שלב 1: קריאת הנתונים ---
# כאן אתה תשים את הנתיב לקובץ ה-CSV האמיתי שלך מתוך data_processed
# לדוגמה: df = pd.read_csv('../data_processed/incomes_04_2026.csv')

# לצורך ההדגמה, אני יוצר נתונים פיקטיביים שייראו כמו ה-CSV שלך:
data = {
    'תאריך תשלום': ['01-04-2026', '02-04-2026', '05-04-2026'],
    'מספר דירה': [3, 7, 12],
    'סכום': [300, 300, 300],
    'שם משלם': ['ישראל ישראלי', 'כהן', 'לוי']
}
df = pd.DataFrame(data)

# --- שלב 2: עיבוד ובדיקת הסטטוס ---
# מוציאים רשימה ייחודית של הדירות ששילמו
paid_apartments = df['מספר דירה'].unique().tolist()
# מסננים את הדירות שלא שילמו
unpaid_apartments = [apt for apt in TOTAL_APARTMENTS if apt not in paid_apartments]

# --- שלב 3: התצוגה הוויזואלית ---
st.markdown("---")

# חלוקה ל-2 עמודות במסך
col1, col2 = st.columns(2)

with col1:
    st.success(f"שילמו ({len(paid_apartments)} דירות) ✅")
    # הצגת טבלה מסודרת של מי ששילם
    st.dataframe(df[['מספר דירה', 'שם משלם', 'סכום', 'תאריך תשלום']], hide_index=True)

with col2:
    st.error(f"טרם שילמו ({len(unpaid_apartments)} דירות) ❌")
    # הצגת הדירות שלא שילמו כרשימה (או כתגיות)
    for apt in unpaid_apartments:
        st.write(f"**דירה {apt}**")