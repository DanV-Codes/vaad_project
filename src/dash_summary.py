import streamlit as st
import pandas as pd

def render_summary(selected_month_name, MONTHS_HEB):
    st.markdown("<h2 style='text-align: center;'>📊 סיכום מצב קופה - ועד בית האגמית 7</h2>", unsafe_allow_html=True)
    st.divider()
    
    df_col = st.session_state.master_df
    df_exp = st.session_state.expenses_df
    
    # --- 1. אחוזי גבייה ---
    st.subheader("📈 מצב גבייה")
    
    # חישוב חודשים עד לחודש הנבחר (כולל)
    current_month_idx = MONTHS_HEB.index(selected_month_name)
    months_to_show = MONTHS_HEB[:current_month_idx+1]
    
    # שורת מדדים (אחוזים)
    cols = st.columns(len(months_to_show) + 1) # +1 עבור קופה קטנה
    
    # קופה קטנה (הנחה: 60 דירות חייבות)
    if 'קופה קטנה' in df_col.columns:
        petty_paid = df_col[df_col['קופה קטנה'] > 0].shape[0]
        petty_pct = (petty_paid / 60) * 100
        cols[0].metric("קופה קטנה", f"{petty_paid}/60", f"{petty_pct:.0f}%", delta_color="off")
    
    # חודשים שוטפים (הנחה: 56 דירות חייבות - ללא ועד)
    for i, month in enumerate(months_to_show):
        if month in df_col.columns:
            paid_apts = df_col[df_col[month] > 0].shape[0]
            pct = (paid_apts / 56) * 100
            # שימוש ב-Metric מובנה ומרשים
            cols[i+1].metric(f"חודש {month}", f"{paid_apts}/56", f"{pct:.0f}%", delta_color="off")
            
    st.divider()
    
    # --- 2. סיכום הוצאות חודשי ---
    st.subheader(f"💸 ריכוז הוצאות - חודש {selected_month_name}")
    
    if df_exp is not None and 'ספק' in df_exp.columns:
        summary_exp = pd.DataFrame()
        summary_exp['ספק'] = df_exp['ספק']
        
        # הוצאה לחודש הנוכחי
        summary_exp['הוצאה החודש'] = df_exp[selected_month_name] if selected_month_name in df_exp.columns else 0
        
        # סה"כ מתחילת השנה (סכימה של כל החודשים מתוך months_to_show)
        summary_exp['סה"כ מתחילת השנה'] = df_exp[[m for m in months_to_show if m in df_exp.columns]].sum(axis=1)
        
        # הערות
        summary_exp['הערות'] = df_exp['הערות'] if 'הערות' in df_exp.columns else ""
        
        # נסנן רק שורות שיש בהן הוצאה מתחילת השנה (כדי להסתיר ספקים ריקים)
        summary_exp = summary_exp[summary_exp['סה"כ מתחילת השנה'] > 0]
        
        # עיצוב הסכומים שיראו עם שקלים ופסיקים, והפיכת 0 לקו ריק למראה נקי
        summary_exp['הוצאה החודש'] = summary_exp['הוצאה החודש'].apply(lambda x: f"₪{x:,.2f}" if x > 0 else "-")
        summary_exp['סה"כ מתחילת השנה'] = summary_exp['סה"כ מתחילת השנה'].apply(lambda x: f"₪{x:,.2f}" if x > 0 else "-")
        
        # הצגת הטבלה סטטית ללא אפשרות עריכה (מושלם לצילום מסך)
        st.dataframe(summary_exp, hide_index=True, use_container_width=True)
        
        # חישוב סה"כ הוצאות החודש
        total_month = df_exp[selected_month_name].sum() if selected_month_name in df_exp.columns else 0
        st.markdown(f"**סה״כ הוצאות לחודש {selected_month_name}: ₪{total_month:,.2f}**")
        
    st.caption("💡 טיפ: ניתן לצלם חלק זה של המסך ולשלוח ישירות לקבוצת הווטסאפ של הבניין.")