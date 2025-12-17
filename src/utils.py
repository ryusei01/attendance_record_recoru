"""
ユーティリティ関数
"""
import re
from datetime import datetime
from typing import Optional, Dict, Any, Union
import json
import os


def normalize_date(date_str: str) -> Optional[str]:
    """
    日付文字列をYYYY-MM-DD形式に正規化
    
    Args:
        date_str: 日付文字列（様々な形式に対応）
    
    Returns:
        正規化された日付文字列（YYYY-MM-DD形式）、失敗時はNone
    """
    if not date_str:
        return None
    
    # よくある日付形式のパターン
    patterns = [
        (r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', '%Y-%m-%d'),  # 2024/01/01, 2024-01-01
        (r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', '%m-%d-%Y'),  # 01/01/2024
        (r'(\d{1,2})[/-](\d{1,2})', '%m-%d'),                # 9/30, 01/01 (月/日のみ、年は現在年を使用)
        (r'(\d{4})年(\d{1,2})月(\d{1,2})日', '%Y-%m-%d'),     # 2024年1月1日
    ]
    
    for pattern, date_format in patterns:
        match = re.search(pattern, str(date_str))
        if match:
            try:
                if '年' in pattern:
                    year, month, day = match.groups()
                    return f"{year}-{int(month):02d}-{int(day):02d}"
                else:
                    parts = match.groups()
                    if len(parts) == 2 and '%m-%d' in date_format:
                        # 月/日のみの形式（例: 9/30）
                        month, day = parts
                        # 現在の年を使用（またはテキストから年を取得）
                        current_year = datetime.now().year
                        # テキストから年を探す（例: "2025生" から 2025 を取得）
                        year_match = re.search(r'(\d{4})', str(date_str))
                        if year_match:
                            year = year_match.group(1)
                        else:
                            year = str(current_year)
                        return f"{year}-{int(month):02d}-{int(day):02d}"
                    elif len(parts[0]) == 4:  # YYYY-MM-DD形式
                        year, month, day = parts
                        return f"{year}-{int(month):02d}-{int(day):02d}"
                    else:  # MM-DD-YYYY形式
                        month, day, year = parts
                        return f"{year}-{int(month):02d}-{int(day):02d}"
            except (ValueError, IndexError):
                continue
    
    # パース可能な形式を試す
    for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y']:
        try:
            dt = datetime.strptime(str(date_str).strip(), fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            continue
    
    return None


def normalize_time(time_str: str) -> Optional[str]:
    """
    時刻文字列をHH:MM形式に正規化
    
    Args:
        time_str: 時刻文字列（様々な形式に対応）
    
    Returns:
        正規化された時刻文字列（HH:MM形式）、失敗時はNone
    """
    if not time_str:
        return None
    
    time_str = str(time_str).strip()
    
    # 時刻パターン（ピリオド区切りも追加）
    patterns = [
        r'(\d{1,2})\.(\d{2})',          # HH.MM (ピリオド区切り)
        r'(\d{1,2}):(\d{2})',           # HH:MM
        r'(\d{1,2})時(\d{2})分',         # HH時MM分
        r'(\d{4})',                      # HHMM
    ]
    
    for pattern in patterns:
        match = re.search(pattern, time_str)
        if match:
            try:
                if '時' in pattern:
                    hour, minute = match.groups()
                    hour = int(hour)
                    minute = int(minute)
                elif len(match.group(0)) == 4 and '.' not in pattern and ':' not in pattern:
                    # HHMM形式
                    time_str_clean = match.group(0)
                    hour = int(time_str_clean[:2])
                    minute = int(time_str_clean[2:])
                else:
                    hour, minute = match.groups()
                    hour = int(hour)
                    minute = int(minute)
                
                if 0 <= hour <= 23 and 0 <= minute <= 59:
                    return f"{hour:02d}:{minute:02d}"
            except (ValueError, IndexError):
                continue
    
    return None


def load_config(config_path: str = "config.json") -> Dict[str, Any]:
    """
    設定ファイルを読み込む
    
    Args:
        config_path: 設定ファイルのパス
    
    Returns:
        設定辞書
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"設定ファイルが見つかりません: {config_path}")
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def calculate_work_hours(start_time: str, end_time: str, break_time: str = "00:00") -> Optional[float]:
    """
    勤務時間を計算（時間単位）
    
    Args:
        start_time: 出勤時刻（HH:MM形式）
        end_time: 退勤時刻（HH:MM形式）
        break_time: 休憩時間（HH:MM形式、デフォルトは00:00）
    
    Returns:
        勤務時間（時間単位）、計算失敗時はNone
    """
    try:
        start = datetime.strptime(start_time, '%H:%M')
        end = datetime.strptime(end_time, '%H:%M')
        break_dt = datetime.strptime(break_time, '%H:%M')
        
        # 日をまたぐ場合の処理
        if end < start:
            end = datetime.strptime(end_time, '%H:%M').replace(day=2)
        
        work_minutes = (end - start).total_seconds() / 60
        break_minutes = (break_dt - datetime.strptime("00:00", '%H:%M')).total_seconds() / 60
        
        work_hours = (work_minutes - break_minutes) / 60
        return round(work_hours, 2)
    except (ValueError, AttributeError):
        return None


def build_date_from_components(record: Dict[str, Union[int, Optional[str]]]) -> Optional[str]:
    if record.get('date'):
        return record['date']
    day = record.get('day')
    if day is None:
        return None
    year = record.get('year')
    month = record.get('month')
    now = datetime.now()
    if year is None:
        year = now.year
    if month is None:
        month = now.month
    try:
        return f"{year}-{int(month):02d}-{int(day):02d}"
    except Exception:
        return None

