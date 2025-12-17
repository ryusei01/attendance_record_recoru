"""
OCRを使用した画像からの勤怠データ抽出
"""
import os
import cv2
import numpy as np
from PIL import Image
import pytesseract
import easyocr
from typing import List, Dict, Optional, Tuple, Union
import re
from datetime import datetime
from .utils import normalize_time


class OCRExtractor:
    """OCRを使用して画像から勤怠データを抽出"""
    
    def __init__(self, use_easyocr: bool = True, poppler_path: Optional[str] = None):
        """
        初期化
        
        Args:
            use_easyocr: EasyOCRを使用するか（True: EasyOCR, False: Tesseract）
            poppler_path: Popplerのbinディレクトリパス（WindowsでPDF処理に必要）
        """
        self.use_easyocr = use_easyocr
        # pdf2imageが参照するPopplerパス（未指定なら環境変数も見る）
        self.poppler_path = poppler_path or os.environ.get("POPPLER_PATH")
        if self.poppler_path:
            self.poppler_path = os.path.expandvars(os.path.expanduser(str(self.poppler_path)))
        if use_easyocr:
            try:
                self.reader = easyocr.Reader(['ja', 'en'], gpu=False)
            except Exception as e:
                print(f"EasyOCRの初期化に失敗しました。Tesseractを使用します: {e}")
                self.use_easyocr = False
                self.reader = None
    
    def detect_skew_angle(self, image: np.ndarray) -> float:
        """
        画像の傾き角度を検出
        
        Args:
            image: グレースケール画像（numpy配列）
        
        Returns:
            傾き角度（度）
        """
        # エッジ検出
        edges = cv2.Canny(image, 50, 150, apertureSize=3)
        
        # Hough変換で直線を検出
        lines = cv2.HoughLines(edges, 1, np.pi / 180, 200)
        
        if lines is None or len(lines) == 0:
            return 0.0
        
        angles = []
        for line in lines:
            rho, theta = line[0]
            # 水平線に近い角度のみを考慮（約0度または180度）
            angle = np.degrees(theta) - 90
            if abs(angle) < 45:  # 45度以内の傾きのみ
                angles.append(angle)
        
        if not angles:
            return 0.0
        
        # 中央値を計算（外れ値の影響を減らす）
        median_angle = np.median(angles)
        return float(median_angle)
    
    def deskew_image(self, image: np.ndarray) -> np.ndarray:
        """
        画像の傾きを補正
        
        Args:
            image: 画像（numpy配列、BGRまたはグレースケール）
        
        Returns:
            補正済み画像
        """
        # グレースケールに変換（まだの場合）
        if len(image.shape) == 3:
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        else:
            gray = image.copy()
        
        # 傾き角度を検出
        angle = self.detect_skew_angle(gray)
        
        # 角度が小さい場合は補正しない（ノイズによる誤検出を避ける）
        if abs(angle) < 0.5:
            return image
        
        # 画像の中心を取得
        h, w = image.shape[:2]
        center = (w // 2, h // 2)
        
        # 回転行列を計算
        rotation_matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
        
        # 回転後の画像サイズを計算
        cos = np.abs(rotation_matrix[0, 0])
        sin = np.abs(rotation_matrix[0, 1])
        new_w = int((h * sin) + (w * cos))
        new_h = int((h * cos) + (w * sin))
        
        # 回転行列を調整（新しいサイズに合わせる）
        rotation_matrix[0, 2] += (new_w / 2) - center[0]
        rotation_matrix[1, 2] += (new_h / 2) - center[1]
        
        # 画像を回転
        if len(image.shape) == 3:
            deskewed = cv2.warpAffine(image, rotation_matrix, (new_w, new_h), 
                                     flags=cv2.INTER_CUBIC, 
                                     borderMode=cv2.BORDER_REPLICATE)
        else:
            deskewed = cv2.warpAffine(image, rotation_matrix, (new_w, new_h), 
                                     flags=cv2.INTER_CUBIC, 
                                     borderMode=cv2.BORDER_REPLICATE)
        
        return deskewed
    
    def preprocess_image(self, image_path: str) -> np.ndarray:
        """
        画像の前処理（ノイズ除去、コントラスト調整、傾き補正など）
        
        Args:
            image_path: 画像ファイルのパス
        
        Returns:
            前処理済み画像（numpy配列）
        """
        # 画像を読み込む
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"画像を読み込めませんでした: {image_path}")
        
        # 傾き補正
        img = self.deskew_image(img)
        
        # グレースケールに変換
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # ノイズ除去
        denoised = cv2.fastNlMeansDenoising(gray, None, 10, 7, 21)
        
        # コントラスト調整
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(denoised)
        
        # 二値化
        _, binary = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        return binary
    
    def extract_text(self, image_path: str) -> str:
        """
        画像からテキストを抽出
        
        Args:
            image_path: 画像ファイルのパス
        
        Returns:
            抽出されたテキスト
        """
        # 画像を読み込んで傾き補正
        img = cv2.imread(image_path)
        if img is None:
            raise ValueError(f"画像を読み込めませんでした: {image_path}")
        
        # 傾き補正
        img = self.deskew_image(img)
        
        # 一時ファイルとして保存（傾き補正済み画像）
        temp_path = f"temp_deskewed_{os.path.basename(image_path)}"
        cv2.imwrite(temp_path, img)
        
        try:
            if self.use_easyocr and self.reader:
                # EasyOCRを使用（傾き補正済み画像）
                results = self.reader.readtext(temp_path)
                text = '\n'.join([result[1] for result in results])
            else:
                # Tesseractを使用
                # 前処理済み画像を使用
                processed_img = self.preprocess_image(temp_path)
                text = pytesseract.image_to_string(processed_img, lang='jpn+eng')
        finally:
            # 一時ファイルを削除
            if os.path.exists(temp_path):
                os.remove(temp_path)
        
        return text
    
    def parse_attendance_data(self, text: str) -> List[Dict[str, Union[int, Optional[str]]]]:
        """
        抽出したテキストから勤怠データをパース
        
        Args:
            text: 抽出されたテキスト
        
        Returns:
            勤怠データのリスト（各要素は日付（日のみ）、出勤時刻、退勤時刻、状態を含む）
        """
        import logging
        logger = logging.getLogger(__name__)
        
        attendance_records = []
        lines = text.split('\n')
        tokens = [ln.strip() for ln in lines if ln.strip()]
        weekdays = {"月", "火", "水", "木", "金", "土", "日"}

        def _weekday(tok: str) -> Optional[str]:
            """
            曜日を検出
            
            Args:
                tok: トークン文字列
            
            Returns:
                曜日文字（月、火、水、木、金、土、日）またはNone
            """
            if not tok:
                return None
            
            # 文字列全体をチェック（最初の文字だけでなく）
            for ch in tok:
                if ch in weekdays:
                    return ch
            
            # 最初の文字もチェック（既存の動作を維持）
            ch = tok[0]
            return ch if ch in weekdays else None

        def _day(tok: str) -> Optional[int]:
            """
            日付を検出（記号やノイズが混在していても抽出）
            
            Args:
                tok: トークン文字列（例: "1 !", "4 :", "27)"）
            
            Returns:
                日付（1-31）またはNone
            """
            if not tok:
                return None
            
            # 時刻パターン（HH.MM や HH:MM）を含む場合は日付として扱わない
            # 例: "9.30", "18.30", "0:30" などは時刻なので日付ではない
            if re.search(r"^\d{1,2}[\.:]\d{2}", tok):
                return None
            
            # 記号や空白を除去してから数字を抽出
            # "1 !" や "4 :" のような場合でも "1" や "4" を抽出
            m = re.search(r"(\d{1,2})", tok)
            if not m:
                return None
            try:
                d = int(m.group(1))
            except ValueError:
                return None
            return d if 1 <= d <= 31 else None

        logger.info("=" * 60)
        logger.info("勤怠データのパース開始")
        logger.info("=" * 60)
        logger.info(f"トークン数: {len(tokens)}")
        logger.info(f"最初の20トークン: {tokens[:20]}")
        
        # 日付を順番に取得する方式に変更
        # 1. まず日付（1-31）を探す
        # 2. その次のトークンが曜日かどうかチェック（なくてもOK）
        # 3. その後のトークンから時刻を探す
        # 4. 日付が順番に存在することを前提に、欠落している日付の空レコードも作成
        
        # 日付マップを作成（日付 -> トークンインデックス）
        day_map: Dict[int, int] = {}  # day -> token_index
        for idx, tok in enumerate(tokens):
            day = _day(tok)
            if day is not None:
                if day not in day_map:  # 最初に見つかった日付のみ記録
                    day_map[day] = idx
                    logger.info(f"日付検出: 日={day}, インデックス={idx}, トークン='{tok}'")
        
        # 日付の範囲を取得（最小日と最大日）
        if not day_map:
            logger.warning("日付が見つかりませんでした")
            return attendance_records
        
        min_day = min(day_map.keys())
        max_day = max(day_map.keys())
        logger.info(f"日付範囲: {min_day}日 ～ {max_day}日")
        
        # 1日から最大日まで順番に処理
        for expected_day in range(1, max_day + 1):
            if expected_day in day_map:
                # 日付が見つかった場合
                i = day_map[expected_day]
                day = expected_day
                
                # 次のトークンが曜日かどうかチェック
                weekday = None
                next_idx = i + 1
                start_scan_idx = next_idx  # 時刻探索の開始位置
                
                if next_idx < len(tokens):
                    weekday = _weekday(tokens[next_idx])
                    if weekday:
                        logger.info(f"日={day}: 曜日検出='{weekday}' (インデックス={next_idx})")
                        start_scan_idx = next_idx + 1  # 曜日の次のトークンから時刻を探す
                    else:
                        # 曜日が見つからない場合、記号や数字の可能性を考慮
                        tok = tokens[next_idx]
                        is_time_pattern = bool(re.search(r"(\d{1,2}[\.:]\d{2})", tok))
                        is_day_number = _day(tok) is not None
                        
                        if is_time_pattern:
                            # 時刻パターンの場合、曜日なしとして扱う
                            logger.warning(f"日={day}: 曜日が見つかりませんでした（次のトークンは時刻パターン: '{tok}'）")
                            start_scan_idx = next_idx  # このトークンから時刻を探す
                        elif is_day_number:
                            # 次の日付の場合、時刻なしとして扱う
                            logger.warning(f"日={day}: 曜日が見つかりませんでした（次のトークンは日付: '{tok}'）")
                            start_scan_idx = len(tokens)  # 時刻なし（探索範囲外）
                        else:
                            # 記号やその他の文字の場合、スキップして次のトークンから探す
                            logger.warning(f"日={day}: 曜日が見つかりませんでした（次のトークンは記号など: '{tok}'）。スキップして時刻を探します。")
                            start_scan_idx = next_idx + 1  # このトークンをスキップ
                
                # 次の日付の位置を事前に取得（探索範囲を限定するため）
                next_day_idx = len(tokens)  # デフォルトは終端
                for check_idx in range(i + 1, len(tokens)):
                    check_day = _day(tokens[check_idx])
                    if check_day is not None and check_day > day:
                        next_day_idx = check_idx
                        logger.info(f"  日={day}: 次の日付（{check_day}）を検出: インデックス={check_idx}")
                        break
                
                # 次の日付が見つかるまで、または終端まで時刻を探す
                found_times: List[str] = []
                off_flag = False
                
                j = start_scan_idx
                while j < next_day_idx:  # 次の日付の位置より前まで
                    tok = tokens[j]
                    
                    # 休暇の検出
                    if "休" in tok or tok in {"欠", "休"} or re.search(r"休(暇|日|業|吸)", tok):
                        off_flag = True
                        logger.info(f"  日={day}: 休暇検出='{tok}' (インデックス={j})")
                    
                    # 時刻の抽出（最初の2つの時刻のみ取得）
                    for raw in re.findall(r"(\d{1,2}[\.:]\d{2})", tok):
                        if len(found_times) >= 2:
                            break  # 2つ取得したら終了
                        nt = normalize_time(raw)
                        if nt:
                            found_times.append(nt)
                            logger.info(f"  日={day}: 時刻検出 '{raw}' -> '{nt}' (トークン='{tok}', インデックス={j})")
                    
                    j += 1
                
                start_time = found_times[0] if len(found_times) >= 1 else None
                end_time = found_times[1] if len(found_times) >= 2 else None
                
                if off_flag or (start_time is None and end_time is None):
                    status = "off"
                elif start_time and end_time:
                    status = "present"
                else:
                    status = "partial"
                
                record = {
                    "day": day,
                    "weekday": weekday,
                    "start_time": start_time,
                    "end_time": end_time,
                    "status": status,
                    "missing_weekday": weekday is None,  # 強調表示用フラグ
                    "missing_day": False,  # 日付は見つかっている
                }
                attendance_records.append(record)
                
                logger.info(
                    f"レコード抽出 [{len(attendance_records)}]: 日={day}, 曜={weekday or '？'}, "
                    f"出勤={start_time or 'なし'}, 退勤={end_time or 'なし'}, 状態={status}"
                )
            else:
                # 日付が見つからなかった場合、空のレコードを作成
                logger.warning(f"日={expected_day}: 日付が見つかりませんでした。空のレコードを作成します。")
                record = {
                    "day": expected_day,
                    "weekday": None,
                    "start_time": None,
                    "end_time": None,
                    "status": "off",
                    "missing_weekday": True,  # 日付も見つかっていないのでTrue
                    "missing_day": True,  # 強調表示用フラグ
                }
                attendance_records.append(record)
        
        logger.info("=" * 60)
        logger.info(f"パース完了: {len(attendance_records)}件のレコードを抽出")
        logger.info("=" * 60)
        
        return attendance_records
    
    def extract_from_image(self, image_path: str) -> List[Dict[str, Optional[str]]]:
        """
        画像ファイルから勤怠データを抽出
        
        Args:
            image_path: 画像ファイルのパス
        
        Returns:
            勤怠データのリスト
        """
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"画像ファイルが見つかりません: {image_path}")
        
        # テキスト抽出
        text = self.extract_text(image_path)
        
        # データパース
        attendance_data = self.parse_attendance_data(text)
        
        return attendance_data
    
    def extract_from_pdf(self, pdf_path: str) -> List[Dict[str, Optional[str]]]:
        """
        PDFファイルから勤怠データを抽出（各ページを画像として処理）
        
        Args:
            pdf_path: PDFファイルのパス
        
        Returns:
            勤怠データのリスト
        """
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDFファイルが見つかりません: {pdf_path}")

        try:
            from pdf2image import convert_from_path
        except ImportError:
            raise ImportError("PDF処理にはpdf2imageが必要です。pip install pdf2imageでインストールしてください。")
        
        # PDFを画像に変換
        kwargs = {}
        if self.poppler_path:
            if not os.path.isdir(self.poppler_path):
                raise FileNotFoundError(
                    "Popplerパスが存在しません。"
                    f" poppler_path={self.poppler_path}"
                    "（例: C:\\poppler\\Library\\bin）"
                )
            kwargs["poppler_path"] = self.poppler_path

        try:
            images = convert_from_path(pdf_path, **kwargs)
        except Exception as e:
            msg = str(e)
            # pdf2imageの典型的なPoppler未設定エラーに対して、具体的な対処を提示
            if ("Unable to get page count" in msg) or ("poppler" in msg.lower()):
                raise RuntimeError(
                    "PDFのページ数取得に失敗しました。Popplerが未インストール、またはPATHが通っていません。\n"
                    "対処:\n"
                    "- Popplerをインストールして `...\\Library\\bin` をPATHに追加する\n"
                    "- もしくは config.json の `ocr.poppler_path` に bin パスを設定する\n"
                    "- もしくは環境変数 `POPPLER_PATH` に bin パスを設定する\n"
                    "（例: C:\\poppler\\Library\\bin）"
                ) from e
            raise
        
        all_records = []
        for i, image in enumerate(images):
            # 一時ファイルとして保存
            temp_path = f"temp_pdf_page_{i}.png"
            image.save(temp_path, 'PNG')
            
            try:
                records = self.extract_from_image(temp_path)
                all_records.extend(records)
            finally:
                # 一時ファイルを削除
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        
        return all_records

