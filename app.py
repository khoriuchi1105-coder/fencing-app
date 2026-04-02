import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import gspread

# --- ページ設定 ---
st.set_page_config(
    page_title="フェンシング パフォーマンス分析",
    page_icon="🤺",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
h1, h2, h3 { color: #1e3d59; font-family: 'Helvetica Neue', sans-serif; }
.stButton>button { border-radius: 5px; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# --- 認証機能 (堅牢化版) ---
def check_password():
    """正しいパスワードが入力された場合にTrueを返す"""

    def password_entered():
        """入力されたパスワードが正しいかチェックする"""
        pw = st.session_state.get("password_input", "")
        if not pw:
            # 何も入力されていない場合はエラー判定も更新も行わない
            return
            
        if pw == st.secrets["password"]:
            st.session_state["password_correct"] = True
            if "password_input" in st.session_state:
                st.session_state["password_input"] = "" # 平文パスワードをクリア
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct"):
        return True

    # 未認証または間違っている場合
    st.title("🔒 認証が必要です")
    st.text_input(
        "パスワードを入力してください", 
        type="password", 
        on_change=password_entered, 
        key="password_input"
    )
    
    if st.session_state.get("password_correct") is False:
        st.error("😕 パスワードが正しくありません")
        
    return False

# --- 定数 ---
DEFAULT_DATA_FILE = "fencing_data.xlsx"
TEMPLATE_FILE = "template.xlsx"
COLS = ['大会名', '大会年月', '試合番号', '予選/本戦', '対戦相手', 'ピリオド', 'イベント時間（秒）',
        'イベント種別', '得点者', '得点の型', '得点エリア', '無効打突（誰）', '勝敗']

# 配色設定
COLOR_MAP_PLAYER = {'自分': '#2E86C1', '相手': '#C0392B'}
COLOR_MAP_TYPE = {
    '攻撃': '#AED6F1',     # 薄い青
    'リポスト': '#F9E79F',   # 薄い黄
    'カウンター': '#A9DFBF', # 薄い緑
    '反則': '#F5B7B1',    # 薄い赤
    'なし': '#E5E7E9'     # グレー
}

# --- データ読み込み・保存 ---
@st.cache_data(ttl=1) # リアルタイム反映のためにキャッシュの保持を1秒にする
def load_data(file_source):
    """
    file_source: ファイルパス(str) または BytesIO オブジェクト
    """
    # 1. Google Sheets の設定がある場合はそちらを優先
    if "gcp_service_account" in st.secrets and "google_sheets" in st.secrets:
        try:
            client = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
            sheet_url = st.secrets["google_sheets"]["url"]
            sheet = client.open_by_url(sheet_url).sheet1
            data = sheet.get_all_values()
            
            if data:
                headers = data[0]
                rows = data[1:]
                df = pd.DataFrame(rows, columns=headers)
            else:
                df = pd.DataFrame(columns=COLS)
            
            # 以降の整形処理
            new_columns = []
            seen = set()
            for col in df.columns:
                base = str(col).strip()
                name = base
                i = 1
                while name in seen:
                    name = f"{base}_{i}"
                    i += 1
                new_columns.append(name)
                seen.add(name)
            df.columns = new_columns
            
            for col in COLS:
                if col not in df.columns:
                    df[col] = ""

            if '得点の型' in df.columns:
                df['得点の型'] = df['得点の型'].fillna('未定義')
            if 'イベント時間（秒）' in df.columns:
                df['イベント時間（秒）'] = pd.to_numeric(df['イベント時間（秒）'], errors='coerce').fillna(0)
            if '試合番号' in df.columns:
                df['試合番号'] = pd.to_numeric(df['試合番号'], errors='coerce').fillna(0).astype(int)
            if 'ピリオド' in df.columns:
                df['ピリオド'] = pd.to_numeric(df['ピリオド'], errors='coerce').fillna(1).astype(int)
            
            return df
        except Exception as e:
            st.warning(f"Googleスプレッドシート連携エラー (ローカルExcelを使用します): {e}")

    # 2. 既存のローカルファイル（Excel）読み込み（フォールバック）
    if isinstance(file_source, str) and not os.path.exists(file_source):
        return pd.DataFrame(columns=COLS)
    try:
        df = pd.read_excel(file_source, sheet_name=0)
        # 列名の空白を除去し一意化
        new_columns = []
        seen = set()
        for col in df.columns:
            base = str(col).strip()
            name = base
            i = 1
            while name in seen:
                name = f"{base}_{i}"
                i += 1
            new_columns.append(name)
            seen.add(name)
        df.columns = new_columns
        # 不足している列を追加
        for col in COLS:
            if col not in df.columns:
                df[col] = ""

        if '得点の型' in df.columns:
            df['得点の型'] = df['得点の型'].fillna('未定義')
        if 'イベント時間（秒）' in df.columns:
            df['イベント時間（秒）'] = pd.to_numeric(df['イベント時間（秒）'], errors='coerce').fillna(0)
        if '試合番号' in df.columns:
            df['試合番号'] = pd.to_numeric(df['試合番号'], errors='coerce').fillna(0).astype(int)
        if 'ピリオド' in df.columns:
            df['ピリオド'] = pd.to_numeric(df['ピリオド'], errors='coerce').fillna(1).astype(int)
        return df
    except Exception as e:
        st.error(f"データ読み込みエラー: {e}")
        return pd.DataFrame(columns=COLS)

def save_to_excel(df, file_path):
    # 1. Google Sheets の設定がある場合はスプレッドシートを更新
    if "gcp_service_account" in st.secrets and "google_sheets" in st.secrets:
        try:
            client = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
            sheet_url = st.secrets["google_sheets"]["url"]
            sheet = client.open_by_url(sheet_url).sheet1
            
            # シートの中身をクリアして上書き
            sheet.clear()
            # 空の値を空白文字に変換してからリストのリストに追加
            sheet.update([df.columns.values.tolist()] + df.fillna("").astype(str).values.tolist())
            
            st.cache_data.clear()
            return True
        except Exception as e:
            st.error(f"スプレッドシートへの保存エラー (権限やURL設定を確認してください): {e}")
            return False

    # 2. 既存のローカルファイルへの保存
    try:
        df.to_excel(file_path, index=False)
        st.cache_data.clear()
        return True
    except Exception as e:
        st.error(f"保存エラー: {e}")
        return False

def get_next_match_number(df, tournament_name):
    """指定大会内の試合番号の最大値+1を返す"""
    if df.empty or '大会名' not in df.columns or '試合番号' not in df.columns:
        return 1
    t_df = df[df['大会名'].astype(str) == str(tournament_name)]
    if t_df.empty:
        return 1
    return int(t_df['試合番号'].max()) + 1

# --- メイン ---
def main():
    st.title("🤺 フェンシング ゲーム分析ダッシュボード")

    with st.expander("📖 使い方（クリックで展開）"):
        st.markdown("""
        1. 左側サイドバーで「分析モード」を選択。
        2. 「⚡ 効率入力」で試合データを記録。
        3. 「📊 分析ダッシュボード」で結果を確認。
        4. 「📝 データ編集」でExcel形式の書き出しが可能。
        """)

    # --- サイドバー: データ管理 ---
    st.sidebar.header("📁 データ管理")
    
    # テンプレートダウンロード
    if os.path.exists(TEMPLATE_FILE):
        with open(TEMPLATE_FILE, "rb") as f:
            st.sidebar.download_button(
                label="📥 テンプレートをダウンロード",
                data=f,
                file_name="fencing_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ファイルアップローダー
    uploaded_file = st.sidebar.file_uploader("Excelファイルをアップロード", type=["xlsx"])
    
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        data_path = None
        st.sidebar.info("💡 アップロードされたデータを表示中。")
    else:
        df = load_data(DEFAULT_DATA_FILE)
        data_path = DEFAULT_DATA_FILE
        if not os.path.exists(DEFAULT_DATA_FILE):
            st.sidebar.warning("⚠️ データをアップロードしてください。")

    # --- セッション状態の初期化 ---
    defaults = {
        'last_tournament': '', 'last_tournament_date': '', 'last_match_num': 1,
        'last_opponent': '', 'last_period': 1, 'last_round': '予選',
        'last_winloss': '-', 'last_time': 0,
        'selected_attack': '攻撃', 'selected_area': '1',
        'pending_scorer': None, 'pending_event_type': None,
        'pending_invalid_who': None,
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

    # --- サイドバー: 比較モード選択 ---
    st.sidebar.header("📊 分析モード")
    analysis_mode = st.sidebar.radio(
        "比較モードを選択",
        ["全試合比較", "勝敗の比較", "予選本戦比較", "大会ごとの比較", "試合ごとの比較"]
    )

    t_list = sorted(df['大会名'].dropna().astype(str).unique().tolist())
    filter_tournaments = []
    filter_match = "すべて"
    filter_period = "すべて"
    filtered_df = df.copy()

    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 絞り込み条件")

    if analysis_mode == "大会ごとの比較":
        filter_tournaments = st.sidebar.multiselect("大会を選択 (最大4つ)", t_list, max_selections=4)
        if not filter_tournaments:
            st.sidebar.info("大会を選択してください。")
            filtered_df = pd.DataFrame(columns=df.columns)
        else:
            filtered_df = filtered_df[filtered_df['大会名'].astype(str).isin(filter_tournaments)]

    elif analysis_mode == "試合ごとの比較":
        sel_t = st.sidebar.selectbox("大会", ["すべて"] + t_list)
        m_df_for_m_selection = df.copy()
        if sel_t != "すべて":
            m_df_for_m_selection = m_df_for_m_selection[m_df_for_m_selection['大会名'].astype(str) == sel_t]
        
        m_list = sorted(m_df_for_m_selection['試合番号'].dropna().astype(int).unique().tolist())
        if m_list:
            filter_match = st.sidebar.selectbox("試合を選択", [str(x) for x in m_list])
            if sel_t != "すべて":
                filtered_df = filtered_df[filtered_df['大会名'].astype(str) == sel_t]
            try:
                filtered_df = filtered_df[filtered_df['試合番号'] == int(filter_match)]
            except ValueError:
                filtered_df = pd.DataFrame(columns=df.columns)
        else:
            st.sidebar.info("データがありません。")
            filtered_df = pd.DataFrame(columns=df.columns)

    filter_period = st.sidebar.selectbox("ピリオド", ["すべて", "1", "2", "3"])
    if filter_period != "すべて":
        filtered_df = filtered_df[filtered_df['ピリオド'] == int(filter_period)]

    # --- タブ ---
    tab1, tab2, tab3 = st.tabs(["📊 分析ダッシュボード", "⚡ 効率入力", "📝 データ編集"])

    # =========================================================
    with tab1:
        st.header(f"📊 {analysis_mode}")

        if filtered_df.empty:
            st.warning("条件に一致するデータがありません。")
        else:
            # セクション1: 全体集計
            col1, col2, col3 = st.columns(3)
            pts_df = filtered_df[filtered_df['イベント種別'] == '得点']
            my_pts = len(pts_df[pts_df['得点者'] == '自分'])
            op_pts = len(pts_df[pts_df['得点者'] == '相手'])
            with col1: st.metric("自分の得点", my_pts)
            with col2: st.metric("相手の得点", op_pts)
            with col3:
                total_pts = my_pts + op_pts
                wr = (my_pts / total_pts * 100) if total_pts > 0 else 0
                st.metric("得点シェア", f"{wr:.1f}%")

            st.markdown("---")

            # セクション2: モード別比較
            analysis_targets = []
            if analysis_mode == "勝敗の比較":
                st.markdown("## ⚔️ 勝敗詳細比較 (W vs L)")
                analysis_targets = [
                    ("🏆 勝ち試合 (W)", filtered_df[filtered_df['勝敗'] == 'W']),
                    ("💔 負け試合 (L)", filtered_df[filtered_df['勝敗'] == 'L'])
                ]
            elif analysis_mode == "予選本戦比較":
                st.markdown("## 🤺 予選 vs 本戦 比較")
                analysis_targets = [
                    ("🟦 予選", filtered_df[filtered_df['予選/本戦'] == '予選']),
                    ("🟥 本戦", filtered_df[filtered_df['予選/本戦'] == '本戦'])
                ]
            elif analysis_mode == "大会ごとの比較":
                st.markdown("## 🏆 大会別比較")
                if filter_tournaments:
                    for t in filter_tournaments:
                        analysis_targets.append((f"🏅 {t}", filtered_df[filtered_df['大会名'].astype(str) == str(t)]))
            elif analysis_mode == "試合ごとの比較":
                st.subheader("📈 試合全体の流れ (スコアフロー)")
                flow_df = filtered_df[filtered_df['イベント種別'].isin(['得点', '無効'])].copy()
                if not flow_df.empty:
                    # 累積計算
                    flow_df = flow_df.sort_values(['ピリオド', 'イベント時間（秒）'])
                    flow_df['自分_点'] = flow_df['得点者'].apply(lambda x: 1 if x == '自分' else 0).cumsum()
                    flow_df['相手_点'] = flow_df['得点者'].apply(lambda x: 1 if x == '相手' else 0).cumsum()
                    
                    # ピリオド分割表示
                    periods = sorted(flow_df['ピリオド'].unique())
                    for p in periods:
                        p_df = flow_df[flow_df['ピリオド'] == p]
                        if p_df.empty: continue
                        
                        st.markdown(f"##### ピリオド {p}")
                        fig = go.Figure()
                        # 自分
                        fig.add_trace(go.Scatter(x=p_df['イベント時間（秒）'], y=p_df['自分_点'],
                                              name='自分', line=dict(color=COLOR_MAP_PLAYER['自分'], width=3),
                                              mode='lines+markers', line_shape='hv'))
                        # 相手
                        fig.add_trace(go.Scatter(x=p_df['イベント時間（秒）'], y=p_df['相手_点'],
                                              name='相手', line=dict(color=COLOR_MAP_PLAYER['相手'], width=3),
                                              mode='lines+markers', line_shape='hv'))
                        
                        # 無効打突
                        i_df = p_df[p_df['イベント種別'] == '無効']
                        if not i_df.empty:
                            for who, color in [("自分", COLOR_MAP_PLAYER['自分']), ("相手", COLOR_MAP_PLAYER['相手'])]:
                                w_i = i_df[i_df['無効打突（誰）'] == who]
                                if not w_i.empty:
                                    fig.add_trace(go.Scatter(x=w_i['イベント時間（秒）'], y=w_i[f'{who}_点'].tolist(),
                                                          mode='markers', name=f'無効({who})',
                                                          marker=dict(symbol='x', size=12, color=color, line=dict(width=1, color='white')),
                                                          hovertemplate=f'無効打突({who})<br>時間: %{{x}}秒'))
                        
                        fig.update_layout(xaxis_title="時間 (秒)", yaxis_title="累積スコア", height=300,
                                         margin=dict(t=10, b=10, l=10, r=10), hovermode='x unified')
                        st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("---")
                analysis_targets = [("この試合の分析", filtered_df)]
            else:
                analysis_targets = [("統計結果", filtered_df)]

            # セクション3: 並列描画
            if analysis_targets:
                # 1. 成功率
                st.subheader("🎯 成功率の比較 (最重要指標)")
                n_targets = len(analysis_targets)
                cols_succ = st.columns(min(n_targets, 4))
                atk_list = ["攻撃", "リポスト", "カウンター", "反則"]
                
                for idx, (lbl, sub) in enumerate(analysis_targets):
                    with cols_succ[idx % 4]:
                        st.markdown(f"#### {lbl}")
                        s_pts = sub[sub['イベント種別'] == '得点']
                        if s_pts.empty:
                            st.caption("データなし")
                            continue
                        
                        m_c = len(s_pts[s_pts['得点者'] == '自分'])
                        o_c = len(s_pts[s_pts['得点者'] == '相手'])
                        st.caption(f"計: 自{m_c} - 相{o_c}")
                        
                        sm_cols = st.columns(2)
                        for aidx, atype in enumerate(atk_list):
                            a_df = s_pts[s_pts['得点の型'] == atype]
                            am = len(a_df[a_df['得点者'] == '自分'])
                            ao = len(a_df[a_df['得点者'] == '相手'])
                            tot = am + ao
                            rate = (am / tot * 100) if tot > 0 else 0
                            with sm_cols[aidx % 2]:
                                st.metric(atype, f"{rate:.1f}%", f"{am}/{tot}")

                st.markdown("---")
                # 2. 円グラフ
                st.subheader("📊 得点の型 比較")
                cols_pie = st.columns(min(n_targets, 4))
                for idx, (lbl, sub) in enumerate(analysis_targets):
                    with cols_pie[idx % 4]:
                        st.markdown(f"**{lbl}**")
                        s_pts = sub[(sub['イベント種別'] == '得点') & (sub['得点者'] == '自分')]
                        if not s_pts.empty:
                            v_c = s_pts['得点の型'].value_counts()
                            # 表示順を指定通りに並べ替え
                            ordered_vt = ["攻撃", "リポスト", "カウンター", "反則"]
                            v_c = v_c.reindex(ordered_vt).fillna(0)
                            v_c = v_c[v_c > 0] # 値が1以上のものだけ表示
                            
                            f_pie = px.pie(values=v_c.values, names=v_c.index, hole=0.4,
                                          color=v_c.index, color_discrete_map=COLOR_MAP_TYPE,
                                          category_orders={"得点の型": ordered_vt})
                            f_pie.update_traces(sort=False) # 自動ソートを無効化
                            f_pie.update_layout(showlegend=(n_targets <= 2), height=250, margin=dict(t=10, b=10))
                            st.plotly_chart(f_pie, use_container_width=True)
                        else: st.caption("データなし")

                st.markdown("---")
                # 3. ピットマップ (縦ならび)
                st.subheader("🤺 視覚的ピットマップ 比較")
                for lbl, sub in analysis_targets:
                    st.markdown(f"**{lbl}**")
                    s_pts = sub[sub['イベント種別'] == '得点'].copy()
                    if s_pts.empty:
                        st.caption("データなし")
                        continue
                        
                    s_pts['得点エリア'] = s_pts['得点エリア'].astype(str)
                    counts = s_pts.groupby(['得点エリア', '得点者']).size().unstack(fill_value=0)
                    stats = counts.to_dict('index')
                    
                    p_cols = st.columns(5)
                    for k in range(1, 6):
                        area = str(k); area_dict = stats.get(area, {})
                        m = area_dict.get('自分', 0); o = area_dict.get('相手', 0)
                        tot = m + o; mp = (m/tot*100) if tot > 0 else 0; op = (o/tot*100) if tot > 0 else 0
                        with p_cols[k-1]:
                            st.markdown(f"""
                            <div style="border: 1px solid #ddd; padding: 5px; text-align: center; background: #fafafa; border-radius: 5px;">
                                <div style="font-size: 0.7rem; color: #666;">エリア{area}</div>
                                <div style="display: flex; justify-content: space-around; font-weight: bold; font-size: 0.9rem;">
                                    <span style="color: #2E86C1;">{m}</span>
                                    <span style="color: #C0392B;">{o}</span>
                                </div>
                                <div style="height: 6px; background: #eee; margin-top: 3px; display: flex;">
                                    <div style="width: {mp}%; background: #2E86C1;"></div>
                                    <div style="width: {op}%; background: #C0392B;"></div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                    st.write("")

    # =========================================================
    with tab2:
        st.header("⚡ 効率入力モード")
        # 簡易フォーム
        with st.expander("🏟️ 試合情報", expanded=True):
            ci1, ci2, ci3 = st.columns(3)
            with ci1: t_in = st.text_input("大会名", value=st.session_state.last_tournament)
            with ci2: d_in = st.text_input("大会年月", value=st.session_state.last_tournament_date)
            with ci3: o_in = st.text_input("対戦相手", value=st.session_state.last_opponent)
            
            ci4, ci5, ci6 = st.columns(3)
            with ci4: p_in = st.selectbox("ピリオド", [1,2,3], index=st.session_state.last_period-1)
            with ci5: r_in = st.selectbox("予選/本戦", ["予選", "本戦"], index=0 if st.session_state.last_round=="予選" else 1)
            with ci6: m_in = st.number_input("試合番号", value=int(st.session_state.last_match_num), step=1)
            
            st.session_state.last_tournament = t_in
            st.session_state.last_tournament_date = d_in
            st.session_state.last_opponent = o_in
            st.session_state.last_period = p_in
            st.session_state.last_round = r_in
            st.session_state.last_match_num = m_in

        # イベント記録
        st.markdown("---")
        st.subheader("⚡ クイック記録")
        
        # エリア選択
        st.markdown("**1. エリアを選択**")
        acols = st.columns(5)
        for i in range(1, 6):
            if acols[i-1].button(f"{'✔' if st.session_state.selected_area==str(i) else ''} {i}", key=f"a_{i}"):
                st.session_state.selected_area = str(i); st.rerun()
        
        # 得点者・種別
        st.markdown("**2. イベント**")
        e1, e2, e3, e4 = st.columns(4)
        if e1.button("🔵 自 得点"): st.session_state.pending_scorer='自分'; st.session_state.pending_event_type='得点'; st.session_state.pending_invalid_who='なし'
        if e2.button("🔴 相 得点"): st.session_state.pending_scorer='相手'; st.session_state.pending_event_type='得点'; st.session_state.pending_invalid_who='なし'
        if e3.button("⚪ 無効(自)"): st.session_state.pending_scorer='なし'; st.session_state.pending_event_type='無効'; st.session_state.pending_invalid_who='自分'
        if e4.button("🔘 無効(相)"): st.session_state.pending_scorer='なし'; st.session_state.pending_event_type='無効'; st.session_state.pending_invalid_who='相手'
        
        if st.session_state.pending_event_type:
            st.info(f"選択中: {st.session_state.pending_event_type} ({st.session_state.pending_scorer}/{st.session_state.pending_invalid_who})")
            
            st.markdown("**3. 得点の型**")
            atk_cols = st.columns(4)
            for at in ["攻撃", "リポスト", "カウンター", "反則"]:
                if atk_cols[["攻撃", "リポスト", "カウンター", "反則"].index(at)].button(f"{'✔' if st.session_state.selected_attack==at else ''} {at}", key=f"at_{at}"):
                    st.session_state.selected_attack = at; st.rerun()
            
            st.markdown("**4. 保存**")
            sc1, sc2 = st.columns([3, 1])
            with sc1: t_val = st.number_input("時間(秒)", value=int(st.session_state.last_time+5), step=1)
            with sc2:
                if st.button("💾 保存", use_container_width=True):
                    new_row = {
                        '大会名': st.session_state.last_tournament, '大会年月': st.session_state.last_tournament_date,
                        '試合番号': int(st.session_state.last_match_num), '対戦相手': st.session_state.last_opponent,
                        'ピリオド': int(st.session_state.last_period), '予選/本戦': st.session_state.last_round,
                        'イベント時間（秒）': t_val, 'イベント種別': st.session_state.pending_event_type,
                        '得点者': st.session_state.pending_scorer, '得点の型': st.session_state.selected_attack if st.session_state.pending_event_type=='得点' else 'なし',
                        '得点エリア': st.session_state.selected_area, '無効打突（誰）': st.session_state.pending_invalid_who, '勝敗': ''
                    }
                    updated = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    if data_path:
                        if save_to_excel(updated, data_path):
                            st.session_state.last_time = t_val; st.rerun()

    # =========================================================
    with tab3:
        st.header("🔍 データ管理")
        edited = st.data_editor(df, num_rows="dynamic", use_container_width=True)
        if st.button("💾 変更を保存 (クラウド/ローカル)"):
            if data_path: save_to_excel(edited, data_path); st.success("保存完了")
            
        st.markdown("---")
        # ダウンロードボタンの追加
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            edited.to_excel(writer, index=False)
        st.download_button(
            label="⬇️ 現在のデータをExcelとしてダウンロード",
            data=output.getvalue(),
            file_name="fencing_data_current.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    if check_password():
        main()
