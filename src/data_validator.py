"""
勤怠データの検証
"""
from typing import List, Dict, Optional, Tuple
from datetime import datetime
from .utils import calculate_work_hours, build_date_from_components


class DataValidator:
    """勤怠データの妥当性を検証"""
    
    def __init__(self):
        """初期化"""
        pass
    
    def validate_record(self, record: Dict[str, Optional[str]]) -> Tuple[bool, List[str]]:
        """
        1件のレコードを検証
        
        Args:
            record: 勤怠レコード（day, start_time, end_time, statusを含む）
        
        Returns:
            (検証結果, エラーメッセージのリスト)
        """
        errors = []
        
        # 必須項目のチェック（dayは必須）
        if record.get('day') is None:
            errors.append("日付（日）が設定されていません")
        else:
            try:
                day = int(record['day'])
                if not (1 <= day <= 31):
                    errors.append(f"日付（日）が範囲外です: {day}")
            except (ValueError, TypeError):
                errors.append(f"日付（日）の形式が不正です: {record.get('day')}")
        
        # 日付文字列を構築（検証用）
        date_str = build_date_from_components(record)
        if date_str:
            try:
                datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                errors.append(f"日付の形式が不正です: {date_str}")
        
        # 出勤時刻のチェック（オプション、status='off'の場合は不要）
        if record.get('start_time'):
            try:
                datetime.strptime(record['start_time'], '%H:%M')
            except ValueError:
                errors.append(f"出勤時刻の形式が不正です: {record['start_time']}")
        elif record.get('status') != 'off':
            # statusが'off'でない場合は出勤時刻が推奨されるが、必須ではない
            pass
        
        # 退勤時刻のチェック（オプション、status='off'の場合は不要）
        if record.get('end_time'):
            try:
                datetime.strptime(record['end_time'], '%H:%M')
            except ValueError:
                errors.append(f"退勤時刻の形式が不正です: {record['end_time']}")
        elif record.get('status') != 'off':
            # statusが'off'でない場合は退勤時刻が推奨されるが、必須ではない
            pass
        
        # 時刻の論理チェック
        if record.get('start_time') and record.get('end_time'):
            start = datetime.strptime(record['start_time'], '%H:%M')
            end = datetime.strptime(record['end_time'], '%H:%M')
            
            # 日をまたぐ場合を考慮
            if end < start:
                # 翌日として扱う
                end = end.replace(day=end.day + 1 if end.day < 28 else 1)
            
            work_minutes = (end - start).total_seconds() / 60
            
            # 勤務時間が負の値になる場合
            if work_minutes < 0:
                errors.append("退勤時刻が出勤時刻より前です")
            
            # 異常に長い勤務時間（24時間超）の警告
            if work_minutes > 24 * 60:
                errors.append(f"勤務時間が24時間を超えています: {work_minutes/60:.1f}時間")
        
        # 勤務時間の計算とチェック（start_timeとend_timeが両方ある場合のみ）
        if record.get('start_time') and record.get('end_time'):
            work_hours = calculate_work_hours(
                record['start_time'],
                record['end_time'],
                '00:00'  # break_timeは不要になったので常に00:00
            )
            
            if work_hours is not None:
                if work_hours < 0:
                    errors.append("退勤時刻が出勤時刻より前です")
                elif work_hours > 16:
                    errors.append(f"勤務時間が異常に長いです: {work_hours:.1f}時間")
        
        return len(errors) == 0, errors
    
    def validate_records(self, records: List[Dict[str, Optional[str]]]) -> Dict[str, any]:
        """
        複数のレコードを検証
        
        Args:
            records: 勤怠レコードのリスト
        
        Returns:
            検証結果の辞書（valid_records, invalid_records, summary）
        """
        valid_records = []
        invalid_records = []
        
        for idx, record in enumerate(records):
            is_valid, errors = self.validate_record(record)
            
            if is_valid:
                valid_records.append(record)
            else:
                invalid_records.append({
                    'index': idx,
                    'record': record,
                    'errors': errors
                })
        
        return {
            'valid_records': valid_records,
            'invalid_records': invalid_records,
            'summary': {
                'total': len(records),
                'valid': len(valid_records),
                'invalid': len(invalid_records)
            }
        }
    
    def check_missing_data(self, records: List[Dict[str, Optional[str]]]) -> List[Dict[str, any]]:
        """
        欠損データを検出
        
        Args:
            records: 勤怠レコードのリスト
        
        Returns:
            欠損データのリスト
        """
        missing_data = []
        
        for idx, record in enumerate(records):
            missing_fields = []
            
            # dayは必須
            if record.get('day') is None:
                missing_fields.append('day')
            
            # statusが'off'でない場合、時刻が推奨される
            if record.get('status') != 'off':
                if not record.get('start_time'):
                    missing_fields.append('start_time')
                if not record.get('end_time'):
                    missing_fields.append('end_time')
            
            if missing_fields:
                missing_data.append({
                    'index': idx,
                    'record': record,
                    'missing_fields': missing_fields
                })
        
        return missing_data

