"""
勤怠記録自動入力アプリ - Streamlit GUI版
"""
import streamlit as st
import pandas as pd
import os
import logging
from pathlib import Path
from src.ocr_extractor import OCRExtractor
from src.excel_extractor import ExcelExtractor
from src.data_validator import DataValidator
from src.recoru_client import RecoruClient
from src.utils import load_config, calculate_work_hours, build_date_from_components, normalize_time

# ログ設定
os.makedirs('logs', exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log', encoding='utf-8'),
        logging.StreamHandler(),  # コンソールにも出力
    ]
)
logger = logging.getLogger(__name__)

# ページ設定
st.set_page_config(
    page_title="勤怠記録自動入力アプリ",
    page_icon="📅",
    layout="wide"
)

# セッション状態の初期化
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = None
if 'validation_result' not in st.session_state:
    st.session_state.validation_result = None
if 'input_results' not in st.session_state:
    st.session_state.input_results = None
if 'debug_info' not in st.session_state:
    st.session_state.debug_info = None


def extract_from_file(file_path: str, file_type: str, poppler_path: str = "", debug: bool = False) -> tuple:
    """
    ファイルから勤怠データを抽出
    
    Returns:
        (records, debug_info): レコードリストとデバッグ情報のタプル
    """
    logger.info(f"extract_from_file開始: file_path={file_path}, file_type={file_type}")
    debug_info = {}
    
    if file_type in ['image', 'pdf']:
        logger.info(f"OCR抽出を使用: use_easyocr=True, poppler_path={poppler_path or 'None'}")
        extractor = OCRExtractor(use_easyocr=True, poppler_path=poppler_path or None)
        if file_type == 'pdf':
            logger.info("PDFファイルの処理を開始")
            debug_info['file_type'] = 'pdf'
            # PDFの場合は各ページのテキストを取得（デバッグモードの場合）
            if debug:
                try:
                    logger.info("PDFデバッグ情報の取得を開始")
                    from pdf2image import convert_from_path
                    kwargs = {}
                    if poppler_path:
                        kwargs["poppler_path"] = poppler_path
                    # デバッグ情報取得のため、PDFを画像に変換（抽出処理の前に実行）
                    images = convert_from_path(file_path, **kwargs)
                    logger.info(f"PDFを画像に変換: {len(images)}ページ")
                    page_texts = []
                    for i, image in enumerate(images):
                        logger.info(f"ページ {i+1}/{len(images)}のテキスト抽出中...")
                        temp_path = f"temp_pdf_page_debug_{i}.png"
                        image.save(temp_path, 'PNG')
                        try:
                            text = extractor.extract_text(temp_path)
                            logger.info(f"ページ {i+1}のテキスト抽出完了: {len(text)}文字")
                            logger.info(f"ページ {i+1}の抽出テキスト内容:\n{text}")
                            page_texts.append({
                                'page': i + 1,
                                'text': text
                            })
                        finally:
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                    debug_info['pdf_pages'] = len(images)
                    debug_info['page_texts'] = page_texts
                    logger.info("PDFデバッグ情報の取得完了")
                except Exception as e:
                    logger.error(f"PDFデバッグ情報の取得エラー: {e}", exc_info=True)
                    debug_info['pdf_debug_error'] = str(e)
            
            # 実際のデータ抽出（デバッグ情報取得後）
            records = extractor.extract_from_pdf(file_path)
            logger.info(f"PDF処理完了: {len(records)}件のレコードを抽出")
        else:
            # 画像の場合、抽出されたテキストを取得
            logger.info("画像ファイルの処理を開始")
            if debug:
                logger.info("OCRテキスト抽出中...")
                text = extractor.extract_text(file_path)
                logger.info(f"OCRテキスト抽出完了: {len(text)}文字")
                logger.info(f"抽出テキスト内容:\n{text}")
                debug_info['extracted_text'] = text
            records = extractor.extract_from_image(file_path)
            logger.info(f"画像処理完了: {len(records)}件のレコードを抽出")
            debug_info['file_type'] = 'image'
    elif file_type == 'excel':
        logger.info("Excelファイルの処理を開始")
        extractor = ExcelExtractor()
        records = extractor.extract_from_excel(file_path)
        logger.info(f"Excel処理完了: {len(records)}件のレコードを抽出")
        debug_info['file_type'] = 'excel'
        if debug:
            # Excelの列情報を取得
            import pandas as pd
            df = pd.read_excel(file_path)
            column_mapping = extractor.detect_columns(df)
            debug_info['excel_columns'] = list(df.columns)
            debug_info['column_mapping'] = column_mapping
            debug_info['excel_preview'] = df.head(10).to_dict('records')
    else:
        logger.error(f"サポートされていないファイル形式: {file_type}")
        raise ValueError(f"サポートされていないファイル形式です: {file_type}")
    
    logger.info(f"extract_from_file完了: {len(records)}件のレコードを返却")
    return records, debug_info


def main():
    """メイン処理"""
    st.title("📅 勤怠記録自動入力アプリ")
    st.markdown("---")

    config = {}
    
    # サイドバー：設定
    with st.sidebar:
        st.header("⚙️ 設定")
        
        config_path = st.text_input("設定ファイルパス", value="config.json")
        
        if os.path.exists(config_path):
            try:
                config = load_config(config_path)
                recoru_config = config.get('recoru', {})
                ocr_config = config.get('ocr', {})
                
                st.success("設定ファイルを読み込みました")
                
                contract_id = st.text_input("契約ID", value=recoru_config.get('contract_id', ''), type='default')
                login_id = st.text_input("ログインID", value=recoru_config.get('login_id', ''), type='default')
                password = st.text_input("パスワード", value=recoru_config.get('password', ''), type='password')
                
                base_url = st.text_input(
                    "Recoru勤怠入力ページURL",
                    value=recoru_config.get('base_url', 'https://app.recoru.in/ap/menuAttendance/?ui=362&pp=1'),
                    help="例: https://app.recoru.in/ap/menuAttendance/?ui=362&pp=1"
                )
                
                profile_path = st.text_input(
                    "Chromeプロファイルパス（オプション）",
                    value=recoru_config.get('profile_path', 'H:\\document\\program\\project\\attendance_record_recoru\\chrome_profile'),
                    help="例: H:\\document\\program\\project\\attendance_record_recoru\\chrome_profile（空欄の場合はデフォルトプロファイル）"
                )

                poppler_path = st.text_input(
                    "Poppler(bin)パス（PDF利用時）",
                    value=str(ocr_config.get('poppler_path', '') or ''),
                    help="例: C:\\poppler\\Library\\bin（PATHに通している場合は空でOK）"
                )
                
                headless_mode = st.checkbox("ヘッドレスモード", value=False)
            except Exception as e:
                st.error(f"設定ファイルの読み込みエラー: {e}")
                contract_id = ""
                login_id = ""
                password = ""
                poppler_path = ""
                headless_mode = False
        else:
            st.warning("設定ファイルが見つかりません")
            contract_id = st.text_input("契約ID", type='default')
            login_id = st.text_input("ログインID", type='default')
            password = st.text_input("パスワード", type='password')
            base_url = st.text_input(
                "Recoru勤怠入力ページURL",
                value="https://app.recoru.in/ap/menuAttendance/?ui=362&pp=1",
                help="例: https://app.recoru.in/ap/menuAttendance/?ui=362&pp=1"
            )
            profile_path = st.text_input(
                "Chromeプロファイルパス（オプション）",
                value="H:\\document\\program\\project\\attendance_record_recoru\\chrome_profile",
                help="例: H:\\document\\program\\project\\attendance_record_recoru\\chrome_profile（空欄の場合はデフォルトプロファイル）"
            )
            poppler_path = st.text_input(
                "Poppler(bin)パス（PDF利用時）",
                value="",
                help="例: C:\\poppler\\Library\\bin（PATHに通している場合は空でOK）"
            )
            headless_mode = st.checkbox("ヘッドレスモード", value=False)
    
    # メインエリア
    tab1, tab2, tab3, tab4 = st.tabs(["📤 ファイル選択", "📊 データ確認", "✅ 検証結果", "🚀 実行"])
    
    # タブ1: ファイル選択
    with tab1:
        st.header("ファイルを選択してください")
        
        file = st.file_uploader(
            "勤怠ファイルをアップロード",
            type=['jpg', 'jpeg', 'png', 'pdf', 'xlsx', 'xls'],
            help="画像ファイル（JPEG、PNG、PDF）またはExcelファイル（.xlsx、.xls）を選択してください"
        )
        
        if file is not None:
            st.success(f"ファイルを読み込みました: {file.name}")
            
            # ファイルを一時保存
            file_ext = Path(file.name).suffix.lower()
            file_type = 'image' if file_ext in ['.jpg', '.jpeg', '.png'] else 'pdf' if file_ext == '.pdf' else 'excel'
            
            temp_path = f"temp_{file.name}"
            with open(temp_path, "wb") as f:
                f.write(file.getbuffer())
            
            if st.button("データを抽出", type="primary"):
                with st.spinner("データを抽出中..."):
                    try:
                        logger.info(f"データ抽出を開始: ファイル={file.name}, タイプ={file_type}")
                        records, debug_info = extract_from_file(temp_path, file_type, poppler_path=poppler_path, debug=True)
                        logger.info(f"抽出完了: {len(records)}件のレコードを抽出")
                        st.session_state.extracted_data = records
                        st.session_state.debug_info = debug_info
                        
                        if len(records) == 0:
                            logger.warning("0件のレコードが抽出されました")
                            st.warning("⚠️ 0件のレコードを抽出しました")
                            st.info("💡 デバッグ情報を確認してください")
                            
                            # デバッグ情報を表示
                            with st.expander("🔍 デバッグ情報", expanded=True):
                                st.write("**ファイルタイプ:**", debug_info.get('file_type', 'unknown'))
                                
                                if debug_info.get('file_type') == 'image' and 'extracted_text' in debug_info:
                                    st.subheader("抽出されたテキスト")
                                    st.text_area("OCRテキスト", debug_info['extracted_text'], height=200, key="debug_text")
                                
                                elif debug_info.get('file_type') == 'pdf':
                                    st.subheader("PDF情報")
                                    if 'pdf_pages' in debug_info:
                                        st.write(f"**総ページ数:** {debug_info['pdf_pages']}")
                                    
                                    if 'page_texts' in debug_info and debug_info['page_texts']:
                                        st.subheader("各ページの抽出テキスト")
                                        for page_info in debug_info['page_texts']:
                                            with st.expander(f"ページ {page_info['page']}", expanded=(page_info['page'] == 1)):
                                                st.text_area(
                                                    f"ページ {page_info['page']}のOCRテキスト",
                                                    page_info['text'],
                                                    height=150,
                                                    key=f"pdf_page_{page_info['page']}"
                                                )
                                    elif 'pdf_debug_error' in debug_info:
                                        st.error(f"PDFデバッグ情報の取得に失敗: {debug_info['pdf_debug_error']}")
                                
                                elif debug_info.get('file_type') == 'excel':
                                    st.subheader("Excel列情報")
                                    st.write("**検出された列:**", debug_info.get('column_mapping', {}))
                                    st.write("**全列名:**", debug_info.get('excel_columns', []))
                                    
                                    if 'excel_preview' in debug_info:
                                        st.subheader("Excelデータプレビュー（最初の10行）")
                                        preview_df = pd.DataFrame(debug_info['excel_preview'])
                                        st.dataframe(preview_df, use_container_width=True)
                        else:
                            st.success(f"✅ {len(records)}件のレコードを抽出しました")
                            logger.info(f"抽出されたレコード: {records}")
                            
                            # 抽出したレコードの詳細を表示
                            with st.expander("📋 抽出したレコードの詳細", expanded=True):
                                st.write(f"**総レコード数:** {len(records)}")
                                for idx, record in enumerate(records, 1):
                                    day_raw = record.get('day')
                                    try:
                                        day_disp = f"{int(day_raw):2d}" if day_raw is not None else "N/A"
                                    except Exception:
                                        day_disp = "N/A"
                                    weekday = record.get('weekday', '')
                                    start = record.get('start_time') or 'なし'
                                    end = record.get('end_time') or 'なし'
                                    status = record.get('status', 'unknown')
                                    st.write(f"{idx}. 日={day_disp}, 曜={weekday or '？'}, 出勤={start:>5s}, 退勤={end:>5s}, 状態={status}")
                        
                        # データ検証
                        logger.info("データ検証を開始")
                        validator = DataValidator()
                        st.session_state.validation_result = validator.validate_records(records)
                        logger.info(f"検証完了: 有効={st.session_state.validation_result['summary']['valid']}, 無効={st.session_state.validation_result['summary']['invalid']}")
                    except Exception as e:
                        st.error(f"データ抽出エラー: {e}")
                        logger.error(f"データ抽出エラー: {e}", exc_info=True)
                    finally:
                        # 一時ファイルを削除
                        if os.path.exists(temp_path):
                            os.remove(temp_path)
    
    # タブ2: データ確認
    with tab2:
        st.header("抽出されたデータ")
        
        if st.session_state.extracted_data:
            # 編集用のデータフレームを作成
            df = pd.DataFrame(st.session_state.extracted_data)
            
            # 勤務時間を計算して追加（表示用）
            if 'start_time' in df.columns and 'end_time' in df.columns:
                df['work_hours'] = df.apply(
                    lambda row: calculate_work_hours(
                        row.get('start_time', ''),
                        row.get('end_time', ''),
                        '00:00'  # break_timeは不要
                    ) if row.get('start_time') and row.get('end_time') else None,
                    axis=1
                )
            
            
            st.subheader("データの編集")
            st.info("💡 以下の表でデータを直接編集できます。編集後は「変更を保存」ボタンをクリックしてください。")
            
            # 編集可能な列を定義（work_hoursは計算列なので編集不可）
            editable_columns = ['day', 'weekday', 'start_time', 'end_time', 'status']
            column_config = {}
            
            # 編集可能な列の設定
            for col in df.columns:
                if col in editable_columns:
                    if col == 'day':
                        column_config[col] = st.column_config.NumberColumn(
                            "日",
                            min_value=1,
                            max_value=31,
                            format="%d"
                        )
                    elif col == 'weekday':
                        column_config[col] = st.column_config.TextColumn(
                            "曜日",
                            help="月、火、水、木、金、土、日のいずれか"
                        )
                    elif col in ['start_time', 'end_time']:
                        column_config[col] = st.column_config.TextColumn(
                            "時刻" if col == 'start_time' else "時刻",
                            help="HH:MM形式（例: 09:30）"
                        )
                    elif col == 'status':
                        column_config[col] = st.column_config.SelectboxColumn(
                            "状態",
                            options=["present", "partial", "off"],
                            help="present: 出退勤あり, partial: 一部のみ, off: 休暇"
                        )
                else:
                    # 編集不可の列
                    column_config[col] = st.column_config.Column(
                        col,
                        disabled=True
                    )
            
            # データエディターで編集
            edited_df = st.data_editor(
                df,
                column_config=column_config,
                use_container_width=True,
                num_rows="fixed",
                key="data_editor"
            )
            
            # 変更を保存するボタン
            col1, col2 = st.columns([1, 4])
            with col1:
                if st.button("変更を保存", type="primary"):
                    # 編集されたデータを元の形式に戻す
                    edited_records = []
                    for idx, row in edited_df.iterrows():
                        # 時刻を正規化
                        start_time = None
                        if pd.notna(row['start_time']) and row['start_time'] != '':
                            start_time = normalize_time(str(row['start_time']))
                            if start_time is None:
                                st.warning(f"行 {idx+1}: 出勤時刻の形式が不正です: {row['start_time']}")
                        
                        end_time = None
                        if pd.notna(row['end_time']) and row['end_time'] != '':
                            end_time = normalize_time(str(row['end_time']))
                            if end_time is None:
                                st.warning(f"行 {idx+1}: 退勤時刻の形式が不正です: {row['end_time']}")
                        
                        record = {
                            'day': int(row['day']) if pd.notna(row['day']) else None,
                            'weekday': row['weekday'] if pd.notna(row['weekday']) else None,
                            'start_time': start_time,
                            'end_time': end_time,
                            'status': row['status'] if pd.notna(row['status']) else 'partial'
                        }
                        edited_records.append(record)
                    
                    # セッション状態を更新
                    st.session_state.extracted_data = edited_records
                    
                    # 検証を再実行
                    logger.info("編集後のデータ検証を開始")
                    validator = DataValidator()
                    st.session_state.validation_result = validator.validate_records(edited_records)
                    logger.info(f"検証完了: 有効={st.session_state.validation_result['summary']['valid']}, 無効={st.session_state.validation_result['summary']['invalid']}")
                    
                    st.success("✅ 変更を保存しました。検証結果タブで確認してください。")
                    st.rerun()
            
            with col2:
                if st.button("元に戻す"):
                    st.rerun()
            
            # データの統計情報を表示
            st.subheader("データ統計")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("総レコード数", len(df))
            with col2:
                present_count = len(df[df['status'] == 'present']) if 'status' in df.columns else 0
                st.metric("出退勤あり", present_count)
            with col3:
                partial_count = len(df[df['status'] == 'partial']) if 'status' in df.columns else 0
                st.metric("一部のみ", partial_count)
            with col4:
                off_count = len(df[df['status'] == 'off']) if 'status' in df.columns else 0
                st.metric("休暇", off_count)
        else:
            st.info("ファイルを選択してデータを抽出してください")
    
    # タブ3: 検証結果
    with tab3:
        st.header("データ検証結果")
        
        # 再検証ボタン（抽出データがある場合のみ表示）
        if st.session_state.extracted_data:
            col_btn1, col_btn2 = st.columns([1, 4])
            with col_btn1:
                if st.button("🔄 再検証", type="primary", help="現在の抽出データで再検証を実行します"):
                    with st.spinner("検証中..."):
                        logger.info("再検証を開始")
                        validator = DataValidator()
                        st.session_state.validation_result = validator.validate_records(st.session_state.extracted_data)
                        logger.info(f"再検証完了: 有効={st.session_state.validation_result['summary']['valid']}, 無効={st.session_state.validation_result['summary']['invalid']}")
                        st.success("再検証が完了しました")
                        st.rerun()
            with col_btn2:
                st.write(f"現在の抽出データ: {len(st.session_state.extracted_data)}件")
        
        if st.session_state.validation_result:
            result = st.session_state.validation_result
            
            col1, col2, col3 = st.columns(3)
            with col1:
                total = result['summary'].get('total', 0) or 0
                st.metric("総レコード数", total)
            with col2:
                valid = result['summary'].get('valid', 0) or 0
                valid_rate = (valid / total * 100) if total else 0.0
                st.metric("有効レコード", valid, delta=f"{valid_rate:.1f}%")
            with col3:
                invalid = result['summary'].get('invalid', 0) or 0
                st.metric("無効レコード", invalid, delta=f"-{invalid}")
            
            if result['invalid_records']:
                st.subheader("⚠️ 無効なレコード")
                for invalid in result['invalid_records']:
                    record = invalid['record']
                    day = record.get('day', 'N/A')
                    weekday = record.get('weekday', '')
                    date_str = build_date_from_components(record) if record.get('day') else 'N/A'
                    with st.expander(f"レコード {invalid['index']} - 日={day}, 曜={weekday or '？'}, 日付={date_str}"):
                        st.json(invalid['record'])
                        st.error("エラー:")
                        for error in invalid['errors']:
                            st.error(f"  - {error}")
            
            if result['valid_records']:
                st.subheader("✅ 有効なレコード")
                valid_df = pd.DataFrame(result['valid_records'])
                st.dataframe(valid_df, width='stretch')
        else:
            if st.session_state.extracted_data:
                st.info("「再検証」ボタンをクリックして検証を実行してください")
            else:
                st.info("データを抽出して検証を実行してください")
    
    # タブ4: 実行
    with tab4:
        st.header("レコルへの自動入力")
        
        if not all([contract_id, login_id, password]):
            st.warning("サイドバーでレコルの認証情報を設定してください")
        elif not st.session_state.validation_result:
            st.warning("先にデータを抽出して検証を実行してください")
        elif not st.session_state.validation_result['valid_records']:
            st.error("有効なレコードがありません")
        else:
            st.info(f"{len(st.session_state.validation_result['valid_records'])}件の有効なレコードを入力します")
            
            if st.button("自動入力を開始", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                log_area = st.empty()
                
                logs = []
                
                def log_callback(message):
                    logs.append(message)
                    log_area.text_area("実行ログ", "\n".join(logs), height=300)
                
                login_success = False
                try:
                    # ログインリトライ設定を取得（configが読み込まれている場合）
                    login_retry_count = 3
                    login_retry_interval = 5
                    if config and 'recoru' in config:
                        login_retry_count = config['recoru'].get('login_retry_count', 3)
                        login_retry_interval = config['recoru'].get('login_retry_interval', 5)
                    
                    client = RecoruClient(
                        contract_id=contract_id,
                        login_id=login_id,
                        password=password,
                        headless=headless_mode,
                        base_url=base_url,
                        profile_path=profile_path if profile_path else None,
                        login_retry_count=login_retry_count,
                        login_retry_interval=login_retry_interval
                    )
                    
                    # ログイン
                    logger.info("レコルへのログインを開始")
                    status_text.text("ログイン中...")
                    if not client.login():
                        logger.error("ログインに失敗しました")
                        st.error("ログインに失敗しました。ブラウザは開いたままです。手動で確認してください。")
                        # ブラウザを閉じずに終了
                        st.session_state['recoru_client'] = client
                        return
                    
                    login_success = True
                    logger.info("ログイン成功")
                    status_text.text("ログイン成功！勤怠データを入力中...")
                    
                    # 勤怠データ入力
                    valid_records = st.session_state.validation_result['valid_records']
                    logger.info(f"勤怠データ入力開始: {len(valid_records)}件のレコード")
                    results = {'success': [], 'failed': []}
                    
                    for i, record in enumerate(valid_records):
                        progress = (i + 1) / len(valid_records)
                        progress_bar.progress(progress)
                        
                        date = record.get('date', 'N/A')
                        status_text.text(f"入力中: {date} ({i+1}/{len(valid_records)})")
                        logger.info(f"レコード {i+1}/{len(valid_records)} を入力中: {date}")
                        
                        if client.input_attendance(record):
                            results['success'].append(date)
                            logger.info(f"✅ {date}: 入力成功")
                            log_callback(f"✅ {date}: 入力成功")
                        else:
                            results['failed'].append(record)
                            logger.warning(f"❌ {date}: 入力失敗")
                            log_callback(f"❌ {date}: 入力失敗")
                    
                    st.session_state.input_results = results
                    
                    # 結果表示
                    logger.info(f"入力処理完了: 成功={len(results['success'])}, 失敗={len(results['failed'])}")
                    st.success("入力処理が完了しました！")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("成功", len(results['success']))
                    with col2:
                        st.metric("失敗", len(results['failed']))
                    
                    if results['failed']:
                        st.warning("失敗したレコード:")
                        for failed in results['failed']:
                            date_str = build_date_from_components(failed) or 'N/A'
                            day = failed.get('day', 'N/A')
                            st.write(f"- 日={day}, 日付={date_str}")
                
                except Exception as e:
                    st.error(f"エラーが発生しました: {e}")
                    logger.error(f"エラー: {e}", exc_info=True)
                    # エラー時もブラウザを閉じない（ログイン成功時のみ閉じる）
                    if 'client' in locals() and hasattr(client, 'driver') and client.driver:
                        st.info("エラーが発生しましたが、ブラウザは開いたままです。手動で確認してください。")
                        st.session_state['recoru_client'] = client
                finally:
                    # ログイン成功して処理が完了した場合のみブラウザを閉じる
                    if login_success and 'client' in locals():
                        client.close()
                    elif 'client' in locals():
                        # ログイン失敗やエラーの場合はブラウザを開いたまま
                        pass
                    progress_bar.progress(1.0)
                    status_text.text("完了")


if __name__ == '__main__':
    main()

