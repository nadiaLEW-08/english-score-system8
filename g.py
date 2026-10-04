import streamlit as st
import pandas as pd
from PIL import Image
import os
from datetime import datetime

# 1. 頁面基本設定
st.set_page_config(page_title="信班英文考卷登記系統", layout="wide")

SUBJECT = "英文"
TOTAL_SEATS = 36  # 座號預設 1 到 36 號
excel_file = "scores.xlsx"

# ---------------------------------------------------------
# Excel 公告狀態讀取與儲存函式
# ---------------------------------------------------------
def load_announcement():
    """從 Excel 讀取英文科目的發布公告"""
    ann = {"has_new": False, "scope": "尚未發布", "date": ""}
    if os.path.exists(excel_file):
        try:
            df_ann = pd.read_excel(excel_file, sheet_name="Announcements")
            row = df_ann[df_ann["科目"] == SUBJECT]
            if not row.empty:
                r = row.iloc[0]
                ann = {
                    "has_new": bool(r["has_new"]),
                    "scope": str(r["scope"]),
                    "date": str(r["date"]) if pd.notna(r["date"]) and str(r["date"]) != "nan" else ""
                }
        except Exception:
            pass
    return ann

def save_announcement():
    """將當前英文科目的發布狀態寫入 Excel"""
    all_sheets = pd.read_excel(excel_file, sheet_name=None) if os.path.exists(excel_file) else {}
    
    data = [{
        "科目": SUBJECT,
        "has_new": st.session_state.announcement["has_new"],
        "scope": st.session_state.announcement["scope"],
        "date": st.session_state.announcement["date"]
    }]
    all_sheets["Announcements"] = pd.DataFrame(data)
    
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        for sheet, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

# 2. Session State 初始化
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "current_page" not in st.session_state:
    st.session_state.current_page = "home"

if "announcement" not in st.session_state:
    st.session_state.announcement = load_announcement()

# ---------------------------------------------------------
# 對話框設定（新增 / 編輯 / 刪除）
# ---------------------------------------------------------
@st.dialog("➕ 新增英文考卷項目")
def add_exam_dialog():
    with st.form("add_form"):
        new_scope = st.text_input("考試範圍（例如：L3 單字小考 / Review 1）")
        new_date = st.date_input("考試日期", datetime.now())
        if st.form_submit_button("確認新增發布", type="primary"):
            st.session_state.announcement = {
                "has_new": True,
                "scope": new_scope if new_scope else "未指定範圍",
                "date": str(new_date)
            }
            save_announcement()
            st.rerun()

@st.dialog("✏️ 編輯考卷項目")
def edit_exam_dialog():
    cur = st.session_state.announcement
    with st.form("edit_form"):
        new_scope = st.text_input("修改考試範圍", value="" if cur["scope"] == "尚未發布" else cur["scope"])
        new_date = st.date_input("修改考試日期", datetime.now())
        if st.form_submit_button("儲存修改", type="primary"):
            st.session_state.announcement = {
                "has_new": True,
                "scope": new_scope if new_scope else "未指定範圍",
                "date": str(new_date)
            }
            save_announcement()
            st.rerun()

@st.dialog("🗑️ 刪除考卷項目")
def delete_exam_dialog():
    st.write("確定要刪除/重置目前發布的英文考試項目嗎？")
    if st.button("確認刪除", type="primary"):
        st.session_state.announcement = {
            "has_new": False,
            "scope": "尚未發布",
            "date": ""
        }
        save_announcement()
        st.rerun()

# ---------------------------------------------------------
# 階段 1：身份驗證階段（無側邊欄）
# ---------------------------------------------------------
if st.session_state.user_role is None:
    st.title("👋 歡迎使用信班英文考卷登記系統")
    st.write("請選擇您的身分進入系統：")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🎓 我是學生", use_container_width=True):
            st.session_state.user_role = "student"
            st.rerun()
    with col2:
        if st.button("👨‍🏫 我是教師", use_container_width=True):
            st.session_state.user_role = "teacher_pending"
            st.rerun()
    st.stop()

if st.session_state.user_role == "teacher_pending":
    st.title("👨‍🏫 教師身分驗證")
    pwd = st.text_input("🔑 請輸入教師驗證碼：", type="password")
    if st.button("確認驗證", type="primary"):
        if pwd == "0112":
            st.session_state.user_role = "teacher"
            st.rerun()
        else:
            st.error("❌ 驗證碼錯誤，請重新輸入！")
    
    if st.button("返回選擇身份"):
        st.session_state.user_role = None
        st.rerun()
    st.stop()

# ---------------------------------------------------------
# 階段 2：初始化 Excel
# ---------------------------------------------------------
if not os.path.exists(excel_file):
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        df_init = pd.DataFrame([{"座號": seat} for seat in range(1, TOTAL_SEATS + 1)])
        df_init.to_excel(writer, sheet_name=SUBJECT, index=False)
        
        df_ann_init = pd.DataFrame([{
            "科目": SUBJECT, "has_new": False, "scope": "尚未發布", "date": ""
        }])
        df_ann_init.to_excel(writer, sheet_name="Announcements", index=False)

# 左側導覽列
with st.sidebar:
    st.title("⚙️ 系統選單")
    if st.session_state.user_role == "student":
        st.info("👤 身分：**學生**")
    else:
        st.success("👨‍🏫 身分：**教師**")

    if st.button("🏠 主頁面", use_container_width=True):
        st.session_state.current_page = "home"
        st.session_state.announcement = load_announcement()
        st.rerun()

    if st.button("🔄 切換/登出身份", use_container_width=True):
        st.session_state.user_role = None
        st.rerun()

    # 教師側欄：僅顯示新增項目
    if st.session_state.user_role == "teacher":
        st.divider()
        st.write("📢 **教師管理**")
        if st.button("➕ 新增考卷項目", type="primary", use_container_width=True):
            add_exam_dialog()

# ==================== 主頁面 (Home) ====================
st.session_state.announcement = load_announcement()
ann = st.session_state.announcement

if st.session_state.current_page == "home":
    st.title("🔤 信班英文考卷登記系統")
    st.divider()

    if ann["scope"] == "尚未發布":
        st.warning("📢 目前暫無發布的英文考卷項目。")
    else:
        # 類 Google Classroom 視窗卡片
        with st.container(border=True):
            col_title, col_action = st.columns([8, 1])
            
            with col_title:
                has_red = "🔴 " if (ann["has_new"] and st.session_state.user_role == "student") else ""
                st.subheader(f"{has_red}📚 英文考卷：{ann['scope']}")
                date_text = f"📅 考試日期：{ann['date']}" if ann['date'] else "📅 考試日期：未指定"
                st.caption(date_text)

            # 教師視角：視窗右上角三個點 (⋮) 選單按鈕
            with col_action:
                if st.session_state.user_role == "teacher":
                    with st.popover("⋮"):
                        if st.button("✏️ 編輯", key="menu_edit", use_container_width=True):
                            edit_exam_dialog()
                        if st.button("🗑️ 刪除", key="menu_delete", use_container_width=True):
                            delete_exam_dialog()

            st.write("")
            if st.button("🔍 查看詳情與登記", type="primary", key="enter_detail"):
                st.session_state.current_page = "detail"
                st.session_state.announcement["has_new"] = False
                save_announcement()
                st.rerun()

# ==================== 細項頁面 (Detail) ====================
elif st.session_state.current_page == "detail":
    st.title("📖 英文考卷詳細資訊")
    st.caption(f"當前項目：**{ann['scope']}** （日期：{ann['date']}）")
    st.divider()

    # --- 教師細項頁面：看成績 + 未繳紅字 + 匯出 Excel ---
    if st.session_state.user_role == "teacher":
        scope_col = ann['scope']

        # 讀取現有成績檔計算未繳交號碼
        if os.path.exists(excel_file):
            all_sheets = pd.read_excel(excel_file, sheet_name=None)
            if SUBJECT in all_sheets:
                df_english = all_sheets[SUBJECT]
                
                if scope_col in df_english.columns:
                    missing_seats = df_english[df_english[scope_col].isna() | (df_english[scope_col] == "")]["座號"].tolist()
                    if missing_seats:
                        missing_str = "、".join([f"{int(s)}號" for s in missing_seats])
                        st.markdown(f"🚨 **未繳交座號（共 {len(missing_seats)} 人）：** <span style='color:red; font-size: 20px; font-weight: bold;'>{missing_str}</span>", unsafe_allow_html=True)
                    else:
                        st.success("🎉 全班座號皆已完成登記！")

        col_tbl, col_exp = st.columns([3, 1])
        with col_tbl:
            st.subheader("📊 全班成績登記表")
        with col_exp:
            # 匯出 (下載) Excel 成績表
            if os.path.exists(excel_file):
                with open(excel_file, "rb") as f:
                    st.download_button(
                        label="📥 匯出全班成績 Excel 檔",
                        data=f,
                        file_name=f"信班英文成績表_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

        if os.path.exists(excel_file):
            all_sheets = pd.read_excel(excel_file, sheet_name=None)
            if SUBJECT in all_sheets:
                st.dataframe(all_sheets[SUBJECT], use_container_width=True)

        st.divider()
        st.subheader("📤 匯入覆蓋 Excel 成績表")
        uploaded_excel = st.file_uploader("選擇 Excel 檔案 (.xlsx) 覆蓋", type=["xlsx"])
        if uploaded_excel is not None:
            if st.button("確認匯入覆蓋", type="primary"):
                try:
                    uploaded_sheets = pd.read_excel(uploaded_excel, sheet_name=None)
                    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                        for sheet, df in uploaded_sheets.items():
                            df.to_excel(writer, sheet_name=sheet, index=False)
                    st.success("✅ 匯入成功！")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ 匯入失敗：{e}")

    # --- 學生細項頁面：登記與上傳 ---
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("1. 選擇座號與登記分數")
            seat_options = [f"{seat} 號" for seat in range(1, TOTAL_SEATS + 1)]
            selected_option = st.selectbox("請選擇您的座號", options=seat_options)
            selected_seat = int(selected_option.replace(" 號", ""))
            
            manual_score = st.number_input("登記分數", min_value=0, max_value=100, value=0)
            uploaded_file = st.file_uploader("上傳考卷照片證明", type=["png", "jpg", "jpeg"])

        with col2:
            st.subheader("2. 考卷照片預覽")
            if uploaded_file is not None:
                image = Image.open(uploaded_file)
                st.image(image, caption="已上傳之考卷證明", use_container_width=True)

        st.divider()

        if st.button("💾 儲存成績至 Excel", type="primary", use_container_width=True):
            if uploaded_file is None:
                st.warning("⚠️ 請上傳考卷照片作為證明！")
            else:
                all_sheets = pd.read_excel(excel_file, sheet_name=None)
                
                if SUBJECT in all_sheets:
                    df_sub = all_sheets[SUBJECT]
                else:
                    df_sub = pd.DataFrame([{"座號": seat} for seat in range(1, TOTAL_SEATS + 1)])
                
                scope_col = ann['scope']
                if scope_col not in df_sub.columns:
                    df_sub[scope_col] = None

                idx_list = df_sub.index[df_sub["座號"] == selected_seat].tolist()
                if idx_list:
                    row_idx = idx_list[0]
                    df_sub.at[row_idx, scope_col] = manual_score

                df_sub = df_sub.sort_values(by="座號").reset_index(drop=True)
                all_sheets[SUBJECT] = df_sub

                with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
                    for sheet, df in all_sheets.items():
                        df.to_excel(writer, sheet_name=sheet, index=False)

                st.success(f"已成功儲存 {selected_seat} 號的【英文 - {scope_col}】成績！")