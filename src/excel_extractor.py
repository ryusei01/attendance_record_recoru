"""
Excelファイルからの勤怠データ抽出
"""
import pandas as pd
from typing import List, Dict, Optional
import re
from .utils import normalize_date, normalize_time


class ExcelExtractor:
    """Excelファイルから勤怠データを抽出"""
    
    def __init__(self):
        """初期化"""
        pass
    
    def detect_columns(self, df: pd.DataFrame) -> Dict[str, Optional[int]]:
        """
        データフレームから勤怠関連の列を自動検出
        
        Args:
            df: pandas DataFrame
        
        Returns:
            列名のマッピング（date, start_time, end_time, break_time）
        """
        column_mapping = {
            'date': None,
            'start_time': None,
            'end_time': None,
            'break_time': None
        }
        
        # 列名のパターン
        patterns = {
            'date': [r'日付', r'date', r'年月日', r'日'],
            'start_time': [r'出勤', r'開始', r'start', r'出社', r'始業'],
            'end_time': [r'退勤', r'終了', r'end', r'退社', r'終業'],
            'break_time': [r'休憩', r'break', r'休み']
        }
        
        for col_idx, col_name in enumerate(df.columns):
            col_name_str = str(col_name).lower()
            
            for key, pattern_list in patterns.items():
                if column_mapping[key] is None:
                    for pattern in pattern_list:
                        if re.search(pattern, col_name_str, re.IGNORECASE):
                            column_mapping[key] = col_idx
                            break
                    if column_mapping[key] is not None:
                        break
        
        return column_mapping
    
    def extract_from_excel(self, excel_path: str, sheet_name: Optional[str] = None) -> List[Dict[str, Optional[str]]]:
        """
        Excelファイルから勤怠データを抽出
        
        Args:
            excel_path: Excelファイルのパス
            sheet_name: シート名（Noneの場合は最初のシート）
        
        Returns:
            勤怠データのリスト
        """
        try:
            # Excelファイルを読み込む
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(excel_path)
        except Exception as e:
            raise ValueError(f"Excelファイルの読み込みに失敗しました: {e}")
        
        # 列を自動検出
        column_mapping = self.detect_columns(df)
        
        # 列が見つからない場合のデフォルト（最初の数列を使用）
        if column_mapping['date'] is None:
            column_mapping['date'] = 0
        if column_mapping['start_time'] is None:
            column_mapping['start_time'] = 1 if len(df.columns) > 1 else None
        if column_mapping['end_time'] is None:
            column_mapping['end_time'] = 2 if len(df.columns) > 2 else None
        if column_mapping['break_time'] is None:
            column_mapping['break_time'] = 3 if len(df.columns) > 3 else None
        
        attendance_records = []
        
        for idx, row in df.iterrows():
            record = {}
            
            # 日付の抽出と正規化
            if column_mapping['date'] is not None:
                date_value = row.iloc[column_mapping['date']]
                if pd.notna(date_value):
                    # pandasの日付型の場合
                    if isinstance(date_value, pd.Timestamp):
                        record['date'] = date_value.strftime('%Y-%m-%d')
                    else:
                        record['date'] = normalize_date(str(date_value))
            
            # 出勤時刻の抽出と正規化
            if column_mapping['start_time'] is not None:
                start_value = row.iloc[column_mapping['start_time']]
                if pd.notna(start_value):
                    if isinstance(start_value, pd.Timestamp):
                        record['start_time'] = start_value.strftime('%H:%M')
                    else:
                        record['start_time'] = normalize_time(str(start_value))
            
            # 退勤時刻の抽出と正規化
            if column_mapping['end_time'] is not None:
                end_value = row.iloc[column_mapping['end_time']]
                if pd.notna(end_value):
                    if isinstance(end_value, pd.Timestamp):
                        record['end_time'] = end_value.strftime('%H:%M')
                    else:
                        record['end_time'] = normalize_time(str(end_value))
            
            # 休憩時間の抽出と正規化
            if column_mapping['break_time'] is not None:
                break_value = row.iloc[column_mapping['break_time']]
                if pd.notna(break_value):
                    if isinstance(break_value, pd.Timedelta):
                        hours = int(break_value.total_seconds() // 3600)
                        minutes = int((break_value.total_seconds() % 3600) // 60)
                        record['break_time'] = f"{hours:02d}:{minutes:02d}"
                    else:
                        record['break_time'] = normalize_time(str(break_value))
            
            # 必須項目が揃っている場合のみ追加
            if record.get('date') and record.get('start_time') and record.get('end_time'):
                if not record.get('break_time'):
                    record['break_time'] = "00:00"
                attendance_records.append(record)
        
        return attendance_records

