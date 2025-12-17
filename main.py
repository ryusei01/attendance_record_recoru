"""
勤怠記録自動入力アプリ - コマンドライン版
"""
import argparse
import sys
import os
import logging
from pathlib import Path
from typing import Optional
from src.ocr_extractor import OCRExtractor
from src.excel_extractor import ExcelExtractor
from src.data_validator import DataValidator
from src.recoru_client import RecoruClient
from src.utils import load_config

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/app.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def ensure_logs_directory():
    """logsディレクトリが存在しない場合は作成"""
    os.makedirs('logs', exist_ok=True)


def extract_from_file(file_path: str, poppler_path: Optional[str] = None) -> list:
    """
    ファイルから勤怠データを抽出
    
    Args:
        file_path: ファイルパス
    
    Returns:
        勤怠データのリスト
    """
    file_ext = Path(file_path).suffix.lower()
    
    if file_ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp']:
        logger.info(f"画像ファイルからデータを抽出中: {file_path}")
        extractor = OCRExtractor(use_easyocr=True)
        return extractor.extract_from_image(file_path)
    
    elif file_ext == '.pdf':
        logger.info(f"PDFファイルからデータを抽出中: {file_path}")
        extractor = OCRExtractor(use_easyocr=True, poppler_path=poppler_path)
        return extractor.extract_from_pdf(file_path)
    
    elif file_ext in ['.xlsx', '.xls']:
        logger.info(f"Excelファイルからデータを抽出中: {file_path}")
        extractor = ExcelExtractor()
        return extractor.extract_from_excel(file_path)
    
    else:
        raise ValueError(f"サポートされていないファイル形式です: {file_ext}")


def main():
    """メイン処理"""
    parser = argparse.ArgumentParser(description='勤怠記録自動入力アプリ')
    parser.add_argument('--file', '-f', required=True, help='入力ファイル（画像またはExcel）')
    parser.add_argument('--config', '-c', default='config.json', help='設定ファイルのパス（デフォルト: config.json）')
    parser.add_argument('--validate-only', action='store_true', help='検証のみ実行（入力は行わない）')
    parser.add_argument('--headless', action='store_true', help='ヘッドレスモードで実行')
    parser.add_argument('--url', '-u', type=str, help='Recoruの勤怠入力ページURL（例: https://app.recoru.in/ap/menuAttendance/?ui=362&pp=1）')
    parser.add_argument('--profile', '-p', type=str, help='Chromeのプロファイルパス（例: C:\\Users\\username\\AppData\\Local\\Google\\Chrome\\User Data）')
    
    args = parser.parse_args()
    
    ensure_logs_directory()
    
    try:
        # 設定ファイルの読み込み
        logger.info(f"設定ファイルを読み込み中: {args.config}")
        config = load_config(args.config)
        recoru_config = config.get('recoru', {})
        ocr_config = config.get('ocr', {})
        poppler_path = ocr_config.get('poppler_path') or os.environ.get("POPPLER_PATH")
        
        # URLの優先順位: コマンドライン引数 > config.json
        base_url = args.url or recoru_config.get('base_url')
        # プロファイルパスの優先順位: コマンドライン引数 > config.json
        profile_path = args.profile or recoru_config.get('profile_path')
        # ログインリトライ設定
        login_retry_count = recoru_config.get('login_retry_count', 3)
        login_retry_interval = recoru_config.get('login_retry_interval', 5)
        
        if not args.validate_only:
            if not all([recoru_config.get('contract_id'), recoru_config.get('login_id'), recoru_config.get('password')]):
                logger.error("設定ファイルにレコルの認証情報が不足しています")
                sys.exit(1)
        
        # ファイルからデータ抽出
        logger.info("=" * 50)
        logger.info("データ抽出を開始します")
        logger.info("=" * 50)
        
        # PDFの場合はPopplerが必要。PATHに通していない場合は poppler_path を指定する。
        records = extract_from_file(args.file, poppler_path=str(poppler_path) if poppler_path else None)
        logger.info(f"抽出されたレコード数: {len(records)}")
        
        # 抽出したレコードの詳細をログと標準出力に表示
        if records:
            logger.info("=" * 60)
            logger.info("抽出されたレコードの詳細:")
            logger.info("=" * 60)
            for idx, record in enumerate(records, 1):
                day = record.get('day', 'N/A')
                start = record.get('start_time', 'なし')
                end = record.get('end_time', 'なし')
                status = record.get('status', 'unknown')
                logger.info(f"レコード {idx:3d}: 日={day:2d}, 出勤={start:>5s}, 退勤={end:>5s}, 状態={status}")
            logger.info("=" * 60)
        
        if not records:
            logger.warning("抽出されたデータがありません")
            sys.exit(1)
        
        # データ検証
        logger.info("=" * 50)
        logger.info("データ検証を開始します")
        logger.info("=" * 50)
        
        validator = DataValidator()
        validation_result = validator.validate_records(records)
        
        logger.info(f"検証結果:")
        logger.info(f"  総数: {validation_result['summary']['total']}")
        logger.info(f"  有効: {validation_result['summary']['valid']}")
        logger.info(f"  無効: {validation_result['summary']['invalid']}")
        
        if validation_result['invalid_records']:
            logger.warning("無効なレコード:")
            for invalid in validation_result['invalid_records']:
                logger.warning(f"  インデックス {invalid['index']}: {invalid['errors']}")
        
        if not validation_result['valid_records']:
            logger.error("有効なレコードがありません")
            sys.exit(1)
        
        # 検証のみの場合はここで終了
        if args.validate_only:
            logger.info("検証のみ実行しました")
            return
        
        # レコルへの入力
        logger.info("=" * 50)
        logger.info("レコルへの自動入力を開始します")
        logger.info("=" * 50)
        
        client = RecoruClient(
            contract_id=recoru_config['contract_id'],
            login_id=recoru_config['login_id'],
            password=recoru_config['password'],
            headless=args.headless,
            base_url=base_url,
            profile_path=profile_path,
            login_retry_count=login_retry_count,
            login_retry_interval=login_retry_interval
        )
        
        login_success = False
        try:
            # ログイン
            if not client.login():
                logger.error("ログインに失敗しました")
                logger.info("ブラウザは開いたままです。手動で確認してください。")
                return  # ブラウザを閉じずに終了
            
            login_success = True
            
            # 勤怠データ入力
            results = client.input_multiple_attendance(validation_result['valid_records'])
            
            logger.info("=" * 50)
            logger.info("入力結果")
            logger.info("=" * 50)
            logger.info(f"  総数: {results['total']}")
            logger.info(f"  成功: {len(results['success'])}")
            logger.info(f"  失敗: {len(results['failed'])}")
            
            if results['failed']:
                logger.warning("失敗したレコード:")
                for failed in results['failed']:
                    logger.warning(f"  日付: {failed['date']}")
        
        finally:
            # ログイン成功時のみブラウザを閉じる
            if login_success:
                client.close()
            else:
                logger.info("ログイン失敗のため、ブラウザは開いたままです。")
        
        logger.info("処理が完了しました")
    
    except FileNotFoundError as e:
        logger.error(f"ファイルが見つかりません: {e}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"エラーが発生しました: {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()

