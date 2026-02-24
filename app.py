import streamlit as st
import time
import random
import pandas as pd
from datetime import datetime, date

# --- ページ設定 ---
st.set_page_config(page_title="Logistics AI Hub", page_icon="📦", layout="wide")

# --- 定数・初期設定 ---
STAFF_LIST = [
    {"id": "staff_a", "name": "山田純也", "role": "リーダー", "color": "blue"},
    {"id": "staff_b", "name": "スタッフB", "role": "スタッフ", "color": "green"},
    {"id": "staff_c", "name": "スタッフC", "role": "スタッフ", "color": "purple"},
    {"id": "all", "name": "全員の日報を見る", "role": "管理者ビュー", "color": "gray"}
]

HOURS = [str(i).zfill(2) for i in range(24)]
MINUTES = [str(i * 5).zfill(2) for i in range(12)]
CATEGORIES = ["現場", "デスクワーク", "会議", "その他"]
WEATHER_OPTIONS = ["晴", "曇", "雨", "雪"]

# --- セッションステート管理 ---
if 'view' not in st.session_state:
    st.session_state.view = 'staff_select'
if 'selected_staff' not in st.session_state:
    st.session_state.selected_staff = None
if 'reports' not in st.session_state:
    st.session_state.reports = []
if 'analysis_result' not in st.session_state:
    st.session_state.analysis_result = None
if 'current_tasks' not in st.session_state:
    st.session_state.current_tasks = [{"sh": "09", "sm": "00", "category": "デスクワーク", "content": ""}]

# --- Excel解析ロジック ---
def parse_logistics_excel(file):
    """添付された特定のExcelフォーマットを解析する"""
    try:
        # 全シート読み込み
        xl = pd.ExcelFile(file)
        new_reports = []
        
        for sheet_name in xl.sheet_names:
            # 「〇日」という名前のシートのみ対象（または全シート）
            df = xl.parse(sheet_name, header=None)
            
            # 安全対策：行や列が極端に少ないシート（空シートなど）はスキップ
            if len(df) < 4 or df.shape[1] < 7:
                continue
            
            # 1. 基本情報の抽出
            # 日付: 4行目(インデックス3) B列(1)
            raw_date = df.iloc[3, 1]
            report_date = str(raw_date.date()) if hasattr(raw_date, 'date') else str(raw_date)
            
            # 天候: 4行目 F列(5)
            weather = str(df.iloc[3, 5]) if not pd.isna(df.iloc[3, 5]) else "晴"
            
            # 報告者: 4行目 G列(6)
            reporter = str(df.iloc[3, 6]) if not pd.isna(df.iloc[3, 6]) else "不明"
            
            # 2. 業務内容の抽出 (6行目〜28行目付近のD列)
            tasks = []
            max_task_row = min(28, len(df)) # 安全対策：実際の行数に合わせてストップ
            for i in range(5, max_task_row):
                content = df.iloc[i, 3] if df.shape[1] > 3 else ""
                category = df.iloc[i, 13] if df.shape[1] > 13 else "現場"
                if pd.notna(content) and str(content).strip() != "" and "午前" not in str(content) and "午後" not in str(content):
                    tasks.append({
                        "time": "設定なし", 
                        "category": str(category) if pd.notna(category) else "現場",
                        "content": str(content)
                    })

            # 3. 特記事項の抽出 (34行目付近)
            note = ""
            max_note_row = min(40, len(df)) # ファイルの行数か40の小さい方までにする
            if len(df) > 33: # そもそも34行目以降が存在する場合のみ実行
                for i in range(33, max_note_row): 
                    val = df.iloc[i, 3] if df.shape[1] > 3 else ""
                    if pd.notna(val):
                        note += str(val) + "\n"

            # 4. 実績数値（特記事項内のテキストから抽出を試みるモックロジック）
            import re
            case_match = re.search(r'(\d+)ケース', note)
            item_match = re.search(r'(\d+)件', note)
            
            new_reports.append({
                "id": int(time.time()) + random.randint(0, 1000),
                "name": reporter,
                "date": report_date,
                "weather": weather,
                "tasks": tasks,
                "note": note.strip(),
                "metrics": {
                    "inbound": 0,
                    "replenishment_count": int(item_match.group(1)) if item_match else 0,
                    "replenishment_cases": int(case_match.group(1)) if case_match else 0
                }
            })
        return new_reports
    except Exception as e:
        st.error(f"Excel解析エラー: {e}")
        return []

# --- ヘルパー関数 ---
def change_view(new_view):
    st.session_state.view = new_view

def select_staff(staff):
    st.session_state.selected_staff = staff
    change_view('dashboard')

def add_task():
    st.session_state.current_tasks.append({"sh": "09", "sm": "00", "category": "現場", "content": ""})

def generate_mock_analysis(reports):
    if not reports: return None
    return {
        "alerts": [
            {"id": 101, "type": "system", "text": "【システム】さいまるV4への移行に伴う操作の戸惑いが散見されます。操作マニュアルの再周知が必要です。", "comment": "", "feedback": None},
            {"id": 102, "type": "operation", "text": "【現場負荷】スギ薬局向けの緊急補充（61件/364ケース等）が多発しており、ステージングエリアが圧迫されています。", "comment": "", "feedback": None}
        ],
        "connections": [
            {"id": 201, "title": "改善事例の水平展開", "text": "山田さんが改善発表大会で提示した「生成AIによる可視化」を他センターへも展開する準備を推奨します。", "comment": "", "feedback": None}
        ],
        "stats": {
            "fieldWorkRatio": "70%", "deskWorkRatio": "25%", "meetingRatio": "5%",
            "trendComment": f"直近{len(reports)}件のデータを分析。改善活動への意欲が高く、現場の細かな課題（SSエラー等）も即時記録されています。"
        }
    }

# --- UI コンポーネント ---

def render_navigation():
    cols = st.columns([3, 2, 2, 2, 2])
    with cols[0]: st.markdown("### 📦 Logistics AI Hub")
    staff_name = st.session_state.selected_staff["name"] if st.session_state.selected_staff else ""
    if staff_name and staff_name != "全員の日報を見る":
        with cols[1]: st.info(f"👤 **{staff_name}**")
    with cols[2]:
        if st.button("スタッフ選択", use_container_width=True): change_view('staff_select'); st.rerun()
    with cols[3]:
        if st.button("日報入力", use_container_width=True): change_view('form'); st.rerun()
    with cols[4]:
        if st.button("一覧ダッシュボード", use_container_width=True): change_view('dashboard'); st.rerun()
    
    # サイドバーにExcelアップローダーを配置
    with st.sidebar:
        st.markdown("### 📥 Excel一括取込")
        st.caption("指定の日報フォーマット(xlsx)を選択してください")
        uploaded_file = st.file_uploader("ファイルを選択", type=["xlsx"])
        if uploaded_file:
            if st.button("データを取り込む", use_container_width=True, type="primary"):
                with st.spinner("解析中..."):
                    new_data = parse_logistics_excel(uploaded_file)
                    if new_data:
                        st.session_state.reports = new_data + st.session_state.reports
                        st.success(f"{len(new_data)}件のシートを取り込みました")
                        time.sleep(1)
                        st.rerun()
    st.markdown("---")

def render_staff_selection():
    st.markdown("<h2 style='text-align: center;'>物流センター日報AIハブ</h2>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    for i, staff in enumerate(STAFF_LIST):
        col = col1 if i % 2 == 0 else col2
        with col:
            if st.button(f"{staff['role']} \n\n ### {staff['name']}", key=f"st_{staff['id']}", use_container_width=True):
                select_staff(staff); st.rerun()

def render_form():
    st.subheader("📝 日報作成")
    with st.form("daily_report_form"):
        c1, c2, c3 = st.columns(3)
        with c1: name = st.text_input("報告者", value=st.session_state.selected_staff["name"] if st.session_state.selected_staff else "")
        with c2: report_date = st.date_input("日付", value=date.today())
        with c3: weather = st.selectbox("天候", WEATHER_OPTIONS)
        
        st.markdown("##### 業務報告")
        new_tasks = []
        for i, task in enumerate(st.session_state.current_tasks):
            tc1, tc2, tc3, tc4 = st.columns([1, 1, 1.5, 5])
            with tc1: sh = st.selectbox("開始時", HOURS, index=HOURS.index(task["sh"]), key=f"sh_{i}")
            with tc2: sm = st.selectbox("分", MINUTES, index=MINUTES.index(task["sm"]), key=f"sm_{i}")
            with tc3: cat = st.selectbox("分類", CATEGORIES, index=CATEGORIES.index(task["category"]), key=f"cat_{i}")
            with tc4: cont = st.text_input("内容", value=task["content"], key=f"cont_{i}")
            new_tasks.append({"sh": sh, "sm": sm, "category": cat, "content": cont})
        
        if st.form_submit_button("＋ 行追加"):
            st.session_state.current_tasks = new_tasks
            st.session_state.current_tasks.append({"sh": "09", "sm": "00", "category": "現場", "content": ""})
            st.rerun()

        note = st.text_area("特記事項", height=150)
        submitted = st.form_submit_button("🚀 日報を提出", type="primary", use_container_width=True)

    if submitted:
        report = {
            "id": int(time.time()), "name": name, "date": str(report_date), "weather": weather,
            "tasks": [{"time": f"{t['sh']}:{t['sm']}", "category": t["category"], "content": t["content"]} for t in new_tasks if t["content"]],
            "note": note, "metrics": {"inbound": 0, "replenishment_count": 0, "replenishment_cases": 0}
        }
        st.session_state.reports.insert(0, report)
        change_view('dashboard'); st.rerun()

def render_dashboard():
    col1, col2 = st.columns([3, 1])
    with col1: st.subheader("📋 日報一覧")
    with col2:
        if st.button("💡 AI分析を実行", type="primary", use_container_width=True):
            st.session_state.analysis_result = generate_mock_analysis(st.session_state.reports)
            change_view('analysis'); st.rerun()

    filtered = st.session_state.reports
    if st.session_state.selected_staff and st.session_state.selected_staff["id"] != "all":
        filtered = [r for r in filtered if r["name"] == st.session_state.selected_staff["name"]]

    for r in filtered:
        with st.container(border=True):
            st.markdown(f"#### {r['name']} ({r['date']})")
            if r['metrics']['replenishment_cases'] > 0:
                st.info(f"実績: {r['metrics']['replenishment_count']}件 / {r['metrics']['replenishment_cases']}ケース")
            st.write(r['note'])
            with st.expander("詳細タイムライン"):
                for t in r["tasks"]: st.write(f"- {t['time']} [{t['category']}] {t['content']}")

def render_analysis():
    res = st.session_state.analysis_result
    st.header("💡 AI分析レポート")
    c_main, c_side = st.columns([2, 1])
    with c_main:
        st.subheader("🚨 抽出された課題")
        for item in res["alerts"]:
            with st.container(border=True):
                st.write(f"**{item['text']}**")
                st.text_area("対応メモ", key=f"memo_{item['id']}")
    with c_side:
        st.subheader("📊 サマリー")
        st.info(res["stats"]["trendComment"])
        for k, v in [("現場", "fieldWorkRatio"), ("デスク", "deskWorkRatio")]:
            st.write(f"{k}: {res['stats'][v]}")
            st.progress(int(res['stats'][v].replace('%', '')))

# --- メイン実行 ---
if st.session_state.view != 'staff_select': render_navigation()
if st.session_state.view == 'staff_select': render_staff_selection()
elif st.session_state.view == 'form': render_form()
elif st.session_state.view == 'dashboard': render_dashboard()
elif st.session_state.view == 'analysis': render_analysis()
