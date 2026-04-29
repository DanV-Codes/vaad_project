import streamlit as st

def render_debit(selected_month_name, MONTHS_HEB):
    st.header(f"💸 הוצאות ספקים - חודש {selected_month_name}")
    
    df_exp = st.session_state.expenses_df
    if df_exp is not None:
        # כאן תכניס את טבלת ההוצאות שעשינו
        exp_display_cols = ['ספק'] + MONTHS_HEB
        if 'הערות' in df_exp.columns:
            exp_display_cols.append('הערות')
            
        edited_expenses = st.data_editor(
            df_exp[[col for col in exp_display_cols if col in df_exp.columns]],
            use_container_width=True,
            hide_index=True
        )
        if st.button('💾 שמור שינויים בהוצאות'):
            st.success("הוצאות נשמרו! (יש להוסיף את לוגיקת השמירה לאקסל)")