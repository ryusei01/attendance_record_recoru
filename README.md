# 勤怠記録自動入力アプリ

**AI 技術（OCR）**を使用して画像形式の勤怠表や Excel ファイルから勤怠情報を読み取り、レコル（https://app.recoru.in/ap/home/）に自動入力するアプリケーションです。

本アプリケーションは、**無料の AI フレームワーク**（EasyOCR、Tesseract OCR）を活用して、画像からテキストを自動認識・抽出します。

## 機能

- 画像ファイル（JPEG、PNG、PDF）からの勤怠データ抽出（AI OCR 技術使用）
- Excel ファイル（.xlsx、.xls）からの勤怠データ抽出
- レコルへの自動ログインと勤怠データ入力
- データプレビューと編集機能
- 処理ログの表示
- ログイン失敗時の自動リトライ機能
- 既に入力がある場合のスキップ機能
- Chrome プロファイルパスの指定（ログイン状態の保持）
- Recoru の URL 指定機能

## AI 技術について

このアプリケーションは、**無料の AI フレームワーク**を使用して画像からテキストを認識・抽出する機能を実装しています。

### 使用している AI 技術

#### 1. EasyOCR（推奨・デフォルト）

- **技術**: ディープラーニングベースの OCR エンジン
- **特徴**:
  - 80 以上の言語に対応（日本語・英語を含む）
  - 事前学習済みのニューラルネットワークモデルを使用
  - 手書き文字や複雑なレイアウトにも対応
  - GPU/CPU 両方で動作可能（GPU 使用時は高速）
- **利点**:
  - 高い認識精度
  - 日本語の認識に優れている
  - 複雑な表形式の読み取りにも対応
- **実装**: `src/ocr_extractor.py`で使用
- **ライセンス**: Apache License 2.0（無料・商用利用可）

#### 2. Tesseract OCR（フォールバック）

- **技術**: オープンソースの OCR エンジン
- **特徴**:
  - 100 以上の言語に対応
  - 長年の開発実績がある安定した OCR エンジン
  - 画像前処理（ノイズ除去、コントラスト調整）と組み合わせて使用
- **利点**:
  - 軽量で高速
  - システムリソースの使用量が少ない
  - 広く使われている実績のある技術
- **実装**: EasyOCR が使用できない場合のフォールバックとして使用
- **ライセンス**: Apache License 2.0（無料・商用利用可）

### AI 技術の動作フロー

1. **画像前処理**（OpenCV 使用）

   - グレースケール変換
   - ノイズ除去
   - コントラスト調整
   - 二値化処理
   - これにより、OCR の認識精度が向上します

2. **テキスト認識**（AI OCR）

   - EasyOCR または Tesseract OCR が画像からテキストを抽出
   - 日本語と英語の両方を認識
   - 文字の位置情報も取得

3. **データ抽出・正規化**

   - 抽出されたテキストから日付、時刻などの勤怠情報をパース
   - 正規表現を使用してパターンマッチング
   - 日付・時刻形式の正規化（YYYY-MM-DD、HH:MM 形式）

4. **データ検証**
   - 抽出されたデータの妥当性をチェック
   - 異常値の検出
   - エラーレポートの生成

### AI 技術の選択

アプリケーションはデフォルトで**EasyOCR**を使用しますが、以下の場合に Tesseract OCR にフォールバックします：

- EasyOCR の初期化に失敗した場合
- システムリソースが限られている場合

### 認識精度を向上させるための工夫

1. **画像前処理の最適化**

   - OpenCV を使用した画像の品質向上
   - ノイズ除去とコントラスト調整により、AI の認識精度を向上

2. **複数の OCR エンジンのサポート**

   - EasyOCR と Tesseract OCR の両方に対応
   - 用途に応じて最適なエンジンを選択可能

3. **柔軟なデータパース**
   - 様々な日付・時刻形式に対応
   - 正規表現による柔軟なパターンマッチング

### 参考リンク

- **EasyOCR**: https://github.com/JaidedAI/EasyOCR
- **Tesseract OCR**: https://github.com/tesseract-ocr/tesseract
- **OpenCV**: https://opencv.org/

## セットアップ

### 1. 必要なソフトウェアのインストール

#### Tesseract OCR のインストール

- Windows: https://github.com/UB-Mannheim/tesseract/wiki からインストーラーをダウンロード
- macOS: `brew install tesseract`
- Linux: `sudo apt-get install tesseract-ocr`

#### Poppler（PDF 処理用）

- Windows: https://github.com/oschwartz10612/poppler-windows/releases/ からダウンロードし、PATH に追加
  - PATH を触れない場合は、`config.json` の `ocr.poppler_path`（例: `C:\\poppler\\Library\\bin`）または環境変数 `POPPLER_PATH` で bin パスを指定できます
- macOS: `brew install poppler`
- Linux: `sudo apt-get install poppler-utils`

#### Python 3.8 以上

- https://www.python.org/downloads/ からダウンロード

### 2. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 3. 設定ファイルの作成

`config.json.example`を参考に、`config.json`ファイルを作成し、レコルのログイン情報を設定してください。

```bash
cp config.json.example config.json
# その後、config.jsonを編集して認証情報を入力
```

```json
{
  "ocr": {
    "poppler_path": "C:\\path\\to\\poppler\\Library\\bin"
  },
  "recoru": {
    "contract_id": "your_contract_id",
    "login_id": "your_login_id",
    "password": "your_password",
    "base_url": "https://app.recoru.in/ap/menuAttendance/?ui=YOUR_UI&pp=1",
    "profile_path": "C:\\path\\to\\chrome_profile",
    "login_retry_count": 3,
    "login_retry_interval": 5
  }
}
```

**設定項目の説明**:

- `ocr.poppler_path`: Poppler の bin ディレクトリパス（PDF 処理時に使用、PATH に通していない場合のみ必要）
- `recoru.base_url`: Recoru の勤怠入力ページ URL（例: `https://app.recoru.in/ap/menuAttendance/?ui=YOUR_UI&pp=1`）
- `recoru.profile_path`: Chrome のプロファイルパス（ログイン状態を保持する場合に使用、空の場合はデフォルトプロファイル）
- `recoru.login_retry_count`: ログイン失敗時のリトライ回数（デフォルト: 3 回）
- `recoru.login_retry_interval`: ログインリトライ間隔（秒、デフォルト: 5 秒）

**注意**: `config.json`は`.gitignore`に含まれているため、Git にコミットされません。

## 使用方法

### GUI 版（Streamlit）

```bash
streamlit run app.py
```

ブラウザが自動的に開き、アプリケーションが起動します。

### コマンドライン版

```bash
# Excelファイルから抽出
python main.py --file path/to/attendance.xlsx

# 画像ファイルから抽出
python main.py --file path/to/attendance.jpg

# PDFファイルから抽出
python main.py --file path/to/attendance.pdf

# 検証のみ実行（レコルへの入力は行わない）
python main.py --file path/to/attendance.xlsx --validate-only

# ヘッドレスモードで実行
python main.py --file path/to/attendance.xlsx --headless

# RecoruのURLを指定
python main.py --file path/to/attendance.xlsx --url "https://app.recoru.in/ap/menuAttendance/?ui=YOUR_UI&pp=1"

# Chromeプロファイルパスを指定
python main.py --file path/to/attendance.xlsx --profile "C:\\Users\\username\\AppData\\Local\\Google\\Chrome\\User Data"
```

## プロジェクト構造

```
attendance_record_recoru/
├── app.py                 # Streamlit GUIアプリケーション
├── main.py                # コマンドライン版メイン
├── config.json            # 設定ファイル（要作成）
├── requirements.txt       # 依存関係
├── README.md              # このファイル
├── 要件定義.md            # 要件定義書
├── src/
│   ├── __init__.py
│   ├── ocr_extractor.py  # OCR画像抽出モジュール
│   ├── excel_extractor.py # Excel抽出モジュール
│   ├── data_validator.py  # データ検証モジュール
│   ├── recoru_client.py   # レコル自動入力クライアント
│   └── utils.py           # ユーティリティ関数
└── logs/                  # ログファイル保存先
```

## 使用例

### 1. 画像ファイルから勤怠データを抽出して入力

```bash
# GUI版を使用
streamlit run app.py
# ブラウザで画像ファイルをアップロードし、データを抽出・確認後、自動入力ボタンをクリック
```

### 2. Excel ファイルから一括入力

```bash
# コマンドライン版を使用
python main.py --file attendance_2024_01.xlsx
```

## トラブルシューティング

### OCR（AI）の認識精度が低い場合

- 画像の解像度を上げる（300dpi 以上推奨）
- 画像のコントラストを調整する
- 手書きの場合は、印刷されたテキストより認識精度が低くなる可能性があります
- EasyOCR を使用している場合、GPU が利用可能であれば認識精度が向上する可能性があります
- 画像が傾いている場合は、事前に補正すると認識精度が向上します

### レコルへのログインに失敗する場合

- 契約 ID、ログイン ID、パスワードが正しいか確認
- レコルの UI が変更されている可能性があります（セレクターの調整が必要）
- ネットワーク接続を確認
- ログイン失敗時は自動的にリトライされます（デフォルト: 3 回、5 秒間隔）
- ログイン失敗時はブラウザが開いたままになるため、手動で確認できます

### Selenium 関連のエラー

- ChromeDriver が最新版か確認（自動インストールされます）
- Chrome ブラウザがインストールされているか確認

## 注意事項

- レコルの利用規約を遵守してください
- 自動化ツールの使用が許可されているか確認してください
- パスワードなどの機密情報は適切に管理してください
- OCR の認識精度は画像の品質に依存します
- レコルの UI 変更により動作しなくなる可能性があります
- 既に入力がある日付は自動的にスキップされます（上書きしません）
- 正常に処理が完了した場合、ブラウザは開いたままになります（手動で確認可能）

## ライセンス

このプロジェクトは個人利用を目的としています。
