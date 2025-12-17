"""
レコルへの自動ログインと勤怠データ入力
"""
import time
from typing import List, Dict, Optional
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import logging
from src.utils import build_date_from_components


class RecoruClient:
    """レコルへの自動ログインと勤怠データ入力クライアント"""
    
    def __init__(self, contract_id: str, login_id: str, password: str, headless: bool = False, base_url: Optional[str] = None, profile_path: Optional[str] = None, login_retry_count: int = 3, login_retry_interval: int = 5):
        """
        初期化
        
        Args:
            contract_id: 契約ID
            login_id: ログインID
            password: パスワード
            headless: ヘッドレスモードで実行するか
            base_url: 勤怠入力ページのベースURL（例: https://app.recoru.in/ap/menuAttendance/?ui=362&pp=1）
            profile_path: Chromeのプロファイルパス（例: C:\\Users\\username\\AppData\\Local\\Google\\Chrome\\User Data）
            login_retry_count: ログインリトライ回数（デフォルト: 3）
            login_retry_interval: ログインリトライ間隔（秒、デフォルト: 5）
        """
        self.contract_id = contract_id
        self.login_id = login_id
        self.password = password
        self.headless = headless
        self.base_url = base_url
        self.profile_path = profile_path
        self.login_retry_count = login_retry_count
        self.login_retry_interval = login_retry_interval
        self.driver = None
        self.logger = logging.getLogger(__name__)
    
    def _setup_driver(self):
        """Seleniumドライバーをセットアップ"""
        chrome_options = Options()
        if self.headless:
            chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-blink-features=AutomationControlled')
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        # プロファイルパスが指定されている場合は使用
        if self.profile_path:
            import os
            profile_path = os.path.expandvars(os.path.expanduser(str(self.profile_path)))
            if os.path.exists(profile_path):
                chrome_options.add_argument(f'--user-data-dir={profile_path}')
                # デフォルトプロファイルを使用する場合
                chrome_options.add_argument('--profile-directory=Default')
                self.logger.info(f"Chromeプロファイルを使用: {profile_path}")
            else:
                self.logger.warning(f"指定されたプロファイルパスが存在しません: {profile_path}")
        
        # ユーザーエージェントを設定（ボット検出を回避）
        chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.maximize_window()
    
    def _attempt_login(self) -> bool:
        """
        ログイン試行（1回）
        
        Returns:
            ログイン成功時True、失敗時False
        """
        try:
            if not self.driver:
                self._setup_driver()
            
            # ログインページにアクセス
            self.logger.info("レコルのログインページにアクセス中...")
            self.driver.get("https://app.recoru.in/ap/home/")
            
            # ページが読み込まれるまで待機
            wait = WebDriverWait(self.driver, 30)
            
            # ログインフォームが表示されるまで待機
            try:
                wait.until(EC.presence_of_element_located((By.ID, "loginForm")))
                self.logger.info("ログインフォームを検出しました")
            except TimeoutException:
                self.logger.error("ログインフォームが見つかりませんでした")
                return False
            
            # セッションタイムアウトエラーメッセージを確認
            try:
                error_element = self.driver.find_element(By.ID, "loginForm.errors")
                if error_element and error_element.is_displayed():
                    error_text = error_element.text
                    self.logger.warning(f"エラーメッセージを検出: {error_text}")
                    if "セッションタイムアウト" in error_text:
                        self.logger.info("セッションタイムアウトが検出されました。ログインを再試行します。")
            except NoSuchElementException:
                pass
            
            # 契約IDフィールドを探して入力（実際のフォームでは contractId）
            contract_field = None
            try:
                contract_field = wait.until(EC.presence_of_element_located((By.ID, "contractId")))
            except TimeoutException:
                try:
                    contract_field = self.driver.find_element(By.NAME, "contractId")
                except NoSuchElementException:
                    self.logger.error("契約IDフィールドが見つかりませんでした")
                    return False
            
            contract_field.clear()
            contract_field.send_keys(self.contract_id)
            self.logger.info("契約IDを入力しました")
            time.sleep(0.5)
            
            # ログインIDフィールドを探して入力（実際のフォームでは authId）
            login_field = None
            try:
                login_field = self.driver.find_element(By.ID, "authId")
            except NoSuchElementException:
                try:
                    login_field = self.driver.find_element(By.NAME, "authId")
                except NoSuchElementException:
                    self.logger.error("ログインIDフィールドが見つかりませんでした")
                    return False
            
            login_field.clear()
            login_field.send_keys(self.login_id)
            self.logger.info("ログインIDを入力しました")
            time.sleep(0.5)
            
            # パスワードフィールドを探して入力
            password_field = None
            try:
                password_field = self.driver.find_element(By.ID, "password")
            except NoSuchElementException:
                try:
                    password_field = self.driver.find_element(By.NAME, "password")
                except NoSuchElementException:
                    self.logger.error("パスワードフィールドが見つかりませんでした")
                    return False
            
            password_field.clear()
            password_field.send_keys(self.password)
            self.logger.info("パスワードを入力しました")
            time.sleep(0.5)
            
            # ログインボタンをクリック（実際のフォームでは submit ボタンをクリック）
            submit_button = None
            try:
                # まず、表示されているログインボタンをクリック（これがsubmitボタンをトリガーする）
                login_button = self.driver.find_element(By.XPATH, "//input[@type='button' and @class='common-btn submit' and @value='ログイン']")
                login_button.click()
                self.logger.info("ログインボタンをクリックしました")
            except NoSuchElementException:
                try:
                    # フォールバック: submitボタンを直接クリック
                    submit_button = self.driver.find_element(By.ID, "submit")
                    submit_button.click()
                    self.logger.info("submitボタンをクリックしました")
                except NoSuchElementException:
                    # さらにフォールバック: JavaScriptでクリック
                    self.driver.execute_script("$('#submit').click();")
                    self.logger.info("JavaScriptでログインボタンをクリックしました")
            
            # ログイン後のページ遷移を待機
            time.sleep(3)
            
            # ログイン成功の確認（URLの変更や特定の要素の出現を確認）
            current_url = self.driver.current_url
            self.logger.info(f"現在のURL: {current_url}")
            
            # ログインフォームがまだ表示されているか確認
            try:
                login_form = self.driver.find_element(By.ID, "loginForm")
                if login_form.is_displayed():
                    # エラーメッセージを確認
                    try:
                        error_element = self.driver.find_element(By.ID, "loginForm.errors")
                        if error_element and error_element.is_displayed():
                            error_text = error_element.text
                            self.logger.error(f"ログインエラー: {error_text}")
                    except NoSuchElementException:
                        pass
                    self.logger.error("ログインに失敗しました（ログインフォームがまだ表示されています）")
                    return False
            except NoSuchElementException:
                # ログインフォームが見つからない = ログイン成功の可能性
                pass
            
            # URLが変更されたか、ログインフォームが消えた場合は成功
            if "home" in current_url.lower() or "main" in current_url.lower() or current_url != "https://app.recoru.in/ap/home/":
                self.logger.info("ログインに成功しました")
                return True
            else:
                # ログインフォームが消えているか確認
                try:
                    login_form = self.driver.find_element(By.ID, "loginForm")
                    if not login_form.is_displayed():
                        self.logger.info("ログインに成功しました（ログインフォームが非表示になりました）")
                        return True
                except NoSuchElementException:
                    self.logger.info("ログインに成功しました（ログインフォームが見つかりませんでした）")
                    return True
                
                self.logger.error("ログインに失敗しました")
                return False
        
        except Exception as e:
            self.logger.error(f"ログイン処理中にエラーが発生しました: {e}", exc_info=True)
            return False
    
    def login(self) -> bool:
        """
        レコルにログイン（リトライ機能付き）
        
        Returns:
            ログイン成功時True、失敗時False
        """
        for attempt in range(1, self.login_retry_count + 1):
            self.logger.info(f"ログイン試行 {attempt}/{self.login_retry_count}")
            
            if self._attempt_login():
                self.logger.info("ログインに成功しました")
                return True
            
            # 最後の試行でない場合は待機してリトライ
            if attempt < self.login_retry_count:
                self.logger.warning(f"ログインに失敗しました。{self.login_retry_interval}秒後にリトライします...")
                time.sleep(self.login_retry_interval)
            else:
                self.logger.error(f"ログインに失敗しました（{self.login_retry_count}回試行しました）")
        
        return False
    
    def input_attendance(self, record: Dict[str, Optional[str]], skip_reload: bool = False) -> bool:
        """
        1件の勤怠データを入力
        
        Args:
            record: 勤怠レコード（day, start_time, end_time, statusを含む）
            skip_reload: ページリロードをスキップするか（複数レコード入力時に使用）
        
        Returns:
            入力成功時True、失敗時False
        """
        try:
            if not self.driver:
                raise ValueError("ログインしていません。先にlogin()を呼び出してください。")
            
            # 勤怠入力ページに遷移（skip_reloadがFalseの場合のみ、かつ現在のURLが正しくない場合のみ）
            if not skip_reload:
                target_url = self.base_url or "https://app.recoru.in/ap/menuAttendance/"
                current_url = self.driver.current_url
                
                # 現在のURLが正しいか確認（ベースURLが含まれているか）
                if target_url in current_url or current_url.startswith("https://app.recoru.in/ap/menuAttendance"):
                    self.logger.info(f"既に正しいページにいます: {current_url}")
                    # ページが読み込まれるまで少し待機
                    time.sleep(1)
                else:
                    # 正しいページにいない場合は遷移
                    if self.base_url:
                        self.logger.info(f"勤怠入力ページに遷移: {self.base_url}")
                        self.driver.get(self.base_url)
                    else:
                        # デフォルトの勤怠入力ページURL
                        self.logger.info("デフォルトの勤怠入力ページに遷移")
                        self.driver.get("https://app.recoru.in/ap/menuAttendance/")
                    time.sleep(2)
            
            # 日（day）を取得
            day = record.get('day')
            if day is None:
                self.logger.error("日（day）が取得できませんでした")
                return False
            
            try:
                day_int = int(day)
            except (ValueError, TypeError):
                self.logger.error(f"日（day）の形式が不正です: {day}")
                return False
            
            # 日付をYYYY-MM-DD形式で取得（ログ用）
            date_str = build_date_from_components(record)
            self.logger.info(f"勤怠データを入力中: 日={day_int} (日付={date_str or 'N/A'})")
            
            # 日（day）で行を探す
            # HTMLの例: <label style="color: red;">1/2(金)</label> や <a onclick="...">1/2(金)</a>
            wait = WebDriverWait(self.driver, 10)
            tr_element = None
            
            # 方法1: 日が表示されているラベルやリンクを探す（例: "1/2" や "1"）
            # 月/日の形式（例: "1/2"）を探す
            day_patterns = [
                f"{day_int}/",  # "1/" で始まる
                f"/{day_int}",  # "/2" で終わる
                f"{day_int}(",  # "1(" で始まる（曜日が続く）
            ]
            
            for pattern in day_patterns:
                try:
                    # 日が含まれるリンクやラベルを探す
                    day_element = wait.until(
                        EC.presence_of_element_located((
                            By.XPATH, 
                            f"//a[contains(text(), '{day_int}')] | //label[contains(text(), '{day_int}')]"
                        ))
                    )
                    tr_element = day_element.find_element(By.XPATH, "./ancestor::tr")
                    self.logger.info(f"日 {day_int} の行を発見（パターン: {pattern}）")
                    break
                except (TimeoutException, NoSuchElementException):
                    continue
            
            # 方法2: まだ見つからない場合は、tr要素のIDで探す（YYYYMMDD形式）
            if tr_element is None:
                if date_str:
                    try:
                        from datetime import datetime
                        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                        date_ymd = date_obj.strftime('%Y%m%d')
                        tr_id = f"tr-{date_ymd}-1"
                        tr_element = wait.until(
                            EC.presence_of_element_located((By.ID, tr_id))
                        )
                        self.logger.info(f"行を発見（ID）: {tr_id}")
                    except (ValueError, TimeoutException):
                        pass
            
            # 方法3: 日付リンクで探す
            if tr_element is None and date_str:
                try:
                    from datetime import datetime
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                    date_ymd = date_obj.strftime('%Y%m%d')
                    date_link = self.driver.find_element(
                        By.XPATH, 
                        f"//a[contains(@onclick, '{date_ymd}')]"
                    )
                    tr_element = date_link.find_element(By.XPATH, "./ancestor::tr")
                    self.logger.info(f"日付リンクから行を発見: {date_ymd}")
                except (ValueError, NoSuchElementException):
                    pass
            
            if tr_element is None:
                self.logger.error(f"日 {day_int} の行が見つかりませんでした")
                return False
            
            # 行が見つかったので、その行のIDや日付情報を取得（後で使用するため）
            try:
                tr_id_attr = tr_element.get_attribute('id')
                if tr_id_attr:
                    # tr-idから日付を抽出（例: "tr-20251205-1" -> "20251205"）
                    if tr_id_attr.startswith('tr-'):
                        date_ymd = tr_id_attr.split('-')[1] if len(tr_id_attr.split('-')) > 1 else None
                    else:
                        date_ymd = None
                else:
                    date_ymd = None
            except:
                date_ymd = None
            
            # date_ymdが取得できない場合は、recordから構築
            if not date_ymd and date_str:
                try:
                    from datetime import datetime
                    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
                    date_ymd = date_obj.strftime('%Y%m%d')
                except ValueError:
                    date_ymd = None
            
            # 既に入力があるか確認（出勤時刻または退勤時刻が既に入力されている場合）
            has_existing_input = False
            try:
                # 出勤時刻フィールドを確認
                start_field = None
                selectors = []
                if date_ymd:
                    selectors.extend([
                        f"input.ID-worktimeStart-{date_ymd}-1",
                        f"input[name*='worktimeStart'][name*='{date_ymd}']",
                    ])
                selectors.extend([
                    "input[name*='worktimeStart']",
                    "input.worktimeStart",
                    "input[class*='worktimeStart']"
                ])
                
                for selector in selectors:
                    try:
                        start_field = tr_element.find_element(By.CSS_SELECTOR, selector)
                        break
                    except NoSuchElementException:
                        continue
                
                if start_field:
                    start_value = start_field.get_attribute('value')
                    if start_value and start_value.strip():
                        has_existing_input = True
                        self.logger.info(f"日 {day_int}: 出勤時刻に既に入力があります（{start_value}）。スキップします。")
                
                # 退勤時刻フィールドを確認
                if not has_existing_input:
                    end_field = None
                    selectors = []
                    if date_ymd:
                        selectors.extend([
                            f"input.ID-worktimeEnd-{date_ymd}-1",
                            f"input[name*='worktimeEnd'][name*='{date_ymd}']",
                        ])
                    selectors.extend([
                        "input[name*='worktimeEnd']",
                        "input.worktimeEnd",
                        "input[class*='worktimeEnd']"
                    ])
                    
                    for selector in selectors:
                        try:
                            end_field = tr_element.find_element(By.CSS_SELECTOR, selector)
                            break
                        except NoSuchElementException:
                            continue
                    
                    if end_field:
                        end_value = end_field.get_attribute('value')
                        if end_value and end_value.strip():
                            has_existing_input = True
                            self.logger.info(f"日 {day_int}: 退勤時刻に既に入力があります（{end_value}）。スキップします。")
                
            except Exception as e:
                self.logger.warning(f"既存入力の確認中にエラー: {e}")
            
            # 既に入力がある場合はスキップ
            if has_existing_input:
                self.logger.info(f"日 {day_int}: 既に入力があるため、スキップしました")
                return True
            
            # 出勤区分を選択（常に「1」（出勤）を選択）
            attend_id = '1'  # 出勤
            
            if attend_id:
                try:
                    # 行内のselect要素を探す（複数のパターンを試す）
                    attend_select = None
                    selectors = []
                    if date_ymd:
                        selectors.extend([
                            f"select.ID-attendKbn-{date_ymd}-1",
                            f"select[name*='attendId'][name*='{date_ymd}']",
                        ])
                    selectors.extend([
                        "select[name*='attendId']",
                        "select.ID-attendKbn",
                        "select[class*='attendKbn']"
                    ])
                    
                    for selector in selectors:
                        try:
                            attend_select = tr_element.find_element(By.CSS_SELECTOR, selector)
                            break
                        except NoSuchElementException:
                            continue
                    
                    if attend_select:
                        from selenium.webdriver.support.ui import Select
                        select = Select(attend_select)
                        select.select_by_value(attend_id)
                        self.logger.info(f"出勤区分を選択: {attend_id}")
                        time.sleep(0.5)
                    else:
                        self.logger.warning("出勤区分のselect要素が見つかりませんでした")
                except Exception as e:
                    self.logger.warning(f"出勤区分の選択中にエラー: {e}")
            
            # 出勤時刻入力
            if record.get('start_time'):
                try:
                    # 時刻をHHMM形式に変換（HH:MM -> HHMM）
                    start_time = record['start_time'].replace(':', '')
                    start_field = None
                    selectors = []
                    if date_ymd:
                        selectors.extend([
                            f"input.ID-worktimeStart-{date_ymd}-1",
                            f"input[name*='worktimeStart'][name*='{date_ymd}']",
                        ])
                    selectors.extend([
                        "input[name*='worktimeStart']",
                        "input.worktimeStart",
                        "input[class*='worktimeStart']"
                    ])
                    
                    for selector in selectors:
                        try:
                            start_field = tr_element.find_element(By.CSS_SELECTOR, selector)
                            break
                        except NoSuchElementException:
                            continue
                    
                    if start_field:
                        start_field.clear()
                        start_field.send_keys(start_time)
                        self.logger.info(f"出勤時刻を入力: {start_time}")
                        time.sleep(0.5)
                    else:
                        self.logger.warning("出勤時刻フィールドが見つかりませんでした")
                except Exception as e:
                    self.logger.warning(f"出勤時刻入力中にエラー: {e}")
            
            # 退勤時刻入力
            if record.get('end_time'):
                try:
                    # 時刻をHHMM形式に変換（HH:MM -> HHMM）
                    end_time = record['end_time'].replace(':', '')
                    end_field = None
                    selectors = []
                    if date_ymd:
                        selectors.extend([
                            f"input.ID-worktimeEnd-{date_ymd}-1",
                            f"input[name*='worktimeEnd'][name*='{date_ymd}']",
                        ])
                    selectors.extend([
                        "input[name*='worktimeEnd']",
                        "input.worktimeEnd",
                        "input[class*='worktimeEnd']"
                    ])
                    
                    for selector in selectors:
                        try:
                            end_field = tr_element.find_element(By.CSS_SELECTOR, selector)
                            break
                        except NoSuchElementException:
                            continue
                    
                    if end_field:
                        end_field.clear()
                        end_field.send_keys(end_time)
                        self.logger.info(f"退勤時刻を入力: {end_time}")
                        time.sleep(0.5)
                    else:
                        self.logger.warning("退勤時刻フィールドが見つかりませんでした")
                except Exception as e:
                    self.logger.warning(f"退勤時刻入力中にエラー: {e}")
            
            # メモ入力（オプション）
            if record.get('memo'):
                try:
                    memo_field = None
                    selectors = []
                    if date_ymd:
                        selectors.extend([
                            f"input.ID-worktimeMemo-{date_ymd}-1",
                            f"input[name*='worktimeMemo'][name*='{date_ymd}']",
                        ])
                    selectors.extend([
                        "input[name*='worktimeMemo']",
                        "input.worktimeMemo",
                        "input[class*='worktimeMemo']"
                    ])
                    
                    for selector in selectors:
                        try:
                            memo_field = tr_element.find_element(By.CSS_SELECTOR, selector)
                            break
                        except NoSuchElementException:
                            continue
                    
                    if memo_field:
                        memo_field.clear()
                        memo_field.send_keys(record['memo'])
                        self.logger.info("メモを入力しました")
                        time.sleep(0.5)
                    else:
                        self.logger.warning("メモフィールドが見つかりませんでした")
                except Exception as e:
                    self.logger.warning(f"メモ入力中にエラー: {e}")
            
            # 変更フラグを設定（onblurイベントをトリガー）
            if date_ymd:
                try:
                    # JavaScriptでsetAttendanceChangeFlagを呼び出す
                    self.driver.execute_script(f"setAttendanceChangeFlag('{date_ymd}-1');")
                    self.logger.info("変更フラグを設定しました")
                    time.sleep(0.5)
                except Exception as e:
                    self.logger.warning(f"変更フラグの設定に失敗: {e}")
            
            # 保存処理（通常は自動保存されるが、明示的に保存ボタンを探す）
            # レコルでは各行の変更が自動保存される可能性があるため、保存ボタンは探さない
            
            self.logger.info(f"勤怠データの入力が完了しました: {date_ymd}")
            return True
        
        except Exception as e:
            self.logger.error(f"勤怠データ入力中にエラーが発生しました: {e}", exc_info=True)
            return False
    
    def input_multiple_attendance(self, records: List[Dict[str, Optional[str]]]) -> Dict[str, any]:
        """
        複数の勤怠データを入力
        
        Args:
            records: 勤怠レコードのリスト
        
        Returns:
            入力結果のサマリー
        """
        results = {
            'success': [],
            'failed': [],
            'total': len(records)
        }
        
        if not records:
            return results
        
        # 最初の1回だけページを読み込む
        self.logger.info(f"勤怠入力ページに遷移（{len(records)}件のレコードを入力します）")
        if self.base_url:
            self.driver.get(self.base_url)
        else:
            self.driver.get("https://app.recoru.in/ap/menuAttendance/")
        time.sleep(2)
        
        # 各レコードを入力（ページリロードなし）
        for idx, record in enumerate(records):
            self.logger.info(f"レコード {idx + 1}/{len(records)} を入力中...")
            success = self.input_attendance(record, skip_reload=True)
            if success:
                results['success'].append(build_date_from_components(record))
            else:
                results['failed'].append({
                    'date': record.get('date'),
                    'record': record
                })
            
            # レート制限を避けるために少し待機
            time.sleep(0.5)
        
        return results
    
    def close(self):
        """ブラウザを閉じる"""
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.logger.info("ブラウザを閉じました")
    
    def __enter__(self):
        """コンテキストマネージャーのエントリ"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """コンテキストマネージャーのエグジット"""
        self.close()

