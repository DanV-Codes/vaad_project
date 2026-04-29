import streamlit as st
import pandas as pd
import datetime
import os

# הגדרות עמוד (חייב להיות הפקודה הראשונה!)
st.set_page_config(page_title='ניהול ועד - האגמית 7', layout='wide', page_icon='🏢')

# יבוא עמודי המשנה
from dash_credits import render_credits
from dash_debit import render_debit
from dash_summary import render_summary

# --- קבועים ---
COMMITTEE_APTS = [19, 33, 26, 38]
MONTHS_HEB = ['ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני', 'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר']
MASTER_FILE = 'האגמית7_כספים_2026.xlsx'

def load_data(sheet_name):
    try:
        return pd.read_excel(MASTER_FILE, sheet_name=sheet_name)
    except Exception as e:
        st.error(f"שגיאה בטעינת הקובץ: {e}")
        return None

# טעינת נתונים ל-Session State כדי שיהיו זמינים בכל הקבצים
if 'master_df' not in st.session_state:
    st.session_state.master_df = load_data('גביה 2026')
if 'expenses_df' not in st.session_state:
    st.session_state.expenses_df = load_data('הוצאות 2026')

st.title('📊 מערכת ניהול ועד בית - האגמית 7')

# --- תפריט ניווט צדדי ---
st.sidebar.title("ניווט")
menu = st.sidebar.radio("בחר עמוד:", ["ניהול גביה 💰", "הוצאות וספקים 💸", "סיכום לווטסאפ 📸"])

current_month_index = datetime.datetime.now().month - 1
selected_month_name = st.sidebar.selectbox('בחר חודש לעבודה:', MONTHS_HEB, index=current_month_index)

st.sidebar.divider()
st.sidebar.info("המערכת מורכבת כעת מ-4 קבצים נפרדים לעבודה מסודרת.")

# --- הפנייה לקבצי המשנה ---
if menu == "ניהול גביה 💰":
    render_credits(selected_month_name, MONTHS_HEB, MASTER_FILE, COMMITTEE_APTS, DEBT_LIMIT=350, PETTY_LIMIT=500)
elif menu == "הוצאות וספקים 💸":
    render_debit(selected_month_name, MONTHS_HEB)
elif menu == "סיכום לווטסאפ 📸":
    render_summary(selected_month_name, MONTHS_HEB)