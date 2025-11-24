# Notes DXL Processing & Rendering Pipeline (notesdb-export)

## 概要

このプロジェクトは、HCL Notes/Domino データベース (.nsf) から文書データを DXL (Domino XML) 形式で抽出し、正規化された中間JSONフォーマット (normalized.json) を経由して、HTML形式に変換するためのPython製パイプラインツールです。
Notes特有のリッチテキスト構造や添付ファイルを可能な限り忠実に保持し、再利用可能な形式で出力することを目的としています。

HTMLの他に、DOCX, PDF, Markdown といった複数のドキュメント形式への対応も検討しましたが、現在はHTML形式への出力に注力しています。アーキテクチャとしては他の形式への拡張も可能ですが、公式なサポートはHTMLのみとなります。

## 主な機能

- Notes文書抽出: 指定されたNotes DB（サーバー/ファイルパス）から、ビューまたはUNIDリストに基づいて文書を取得しDXL形式でエクスポート。  
- 高精度DXL解析:  
  - リッチテキストの構造 (runs 配列) を詳細に解析。  
  - 文字スタイル（太字、斜体、下線、色、サイズ、背景色、上付き/下付き等）を保持。  
  - テーブル構造（セル結合、スタイル含む）を抽出。  
  - 文書リンク (doclink)、URLリンクを解析。  
  - 添付ファイル参照 (attachmentref, <picture>) をメタデータとして抽出。  
- 添付ファイル抽出: DXL内の埋め込みファイル (<filedata>, Base64) をデコードし、指定ディレクトリに実ファイルを保存し、JSON内のパス情報を更新。  
- HTML出力: 中間JSONを元に、HTML形式でドキュメントを生成します。
- フォント対応: 設定に基づき日本語フォント（例: Noto Sans JP）をHTML/PDF/DOCX出力に適用。  
- 進捗管理: JSONL形式 (progress.jsonl) で処理状況を記録し、中断・再開に対応。  
- 柔軟な実行: 単一のCLI (main.py) によるサブコマンド実行と、設定プロファイル (profile.json) によるバッチ実行に対応。

## アーキテクチャ概要

本パイプラインは、責務分離の原則に基づき、以下の主要コンポーネントで構成されています。

1. main.py (CLI司令塔):  
   - すべての操作の単一エントリーポイント。  
   - run-manifest, normalize, render などのサブコマンドを解釈し、対応する処理を実行します。  
   - profile.json プロファイルの実行も管理します。  
2. src/pipelines/flows.py (Notes連携フロー):  
   - Notes COM APIを介したDB接続、文書取得、DXLエクスポートを含む統合処理（run-unified, run_from_manifest）をライブラリとして提供します。  
3. src/core/ (コアロジック):  
   - dxl/parser.py: DXLを解析し、*初期* normalized.json オブジェクトを生成します。（パス情報なし）  
   - attachments.py: DXLと初期JSONを入力とし、添付ファイルを抽出し、JSON内のパス情報（content_path, src）を *更新* します。  
   - render/engine.py: 更新済みのJSONを入力とし、各フォーマット (HTML/MD/DOCX/PDF) のレンダリングを行います。  
4. src/utils/: ロギング、ファイル操作、進捗管理などの共通機能を提供します。

データフロー (Decoupledモード):

1. export: Notes DB → DXL  
2. normalize: DXL → 初期JSON (content_path: null)  
3. extract: (DXL + 初期JSON) → 添付ファイル保存 + JSON更新 (content_path: "...")  
4. render: 更新済みJSON → HTML (MD/DOCX/PDFは非対応)

## セットアップ

### 動作環境

- OS: Windows  
- ソフトウェア:  
  - **HCL Notes 12 Client** - COM APIのためにインストールと初期設定が必要です。  
  - **Python 3.13 (32bit版)** - Notes COM API（dllが32bit）を使うため**32bit版**が必要です。  
- その他: 日本語フォント (例: Noto Sans JP) がシステムにインストールされていること。

### インストール手順

1. Python 3.13 (32bit) の準備: [Python公式サイト](https://www.python.org/) などから32bit版インストーラーを入手し、インストールします。
インストール時に `Add Python 3.XX to PATH` 及び、 `py launcher`　にチェックを入れることを推奨します。

2. リポジトリのクローン:  
　　/# "C:\MyPythonProjects"に配置する場合の例（zip展開でもOK）

   ```PowerShell
   cd C:\MyPythonProjects
   ```

   ```PowerShell
   git clone https://github.com/tosuru/notesdb-export
   ```

   ```PowerShell
   cd notedb-export
   ```

3. 仮想環境の作成と有効化 (推奨):  

   ```PowerShell
   py -3.13-32 -m venv .venv  
   ```

   ```PowerShell
   .venv/Scripts/Activate.ps1
   ```

4. 依存ライブラリのインストール:

   ```PowerShell
   pip install --upgrade pip
   ```

   ```PowerShell
   pip install -r requirements.txt
   ```

5. pywin32 の post-install スクリプト実行 (初回のみ):  

  ```PowerShell
   python .venv/Scripts/pywin32_postinstall.py -install
   ```

   *(COMコンポーネントをPythonから利用可能にするために必要です)*

### 設定

1. 環境変数：Notesパスワードは安全のため、環境変数`NOTES_PASSWORD`で設定することを推奨します。`PYTHONPATH`は設定必須です。
  　/# Case1:一時的な環境変数設定（現在のセッション内だけ有効にするコマンド例）

   ```PowerShell
   $env:NOTES_PASSWORD="your_notes_password"
   $env:NOTES_REDIRECT_BASE="http://xxx="
   $env:PYTHONPATH = "$(Get-Location)/src"
   ```

   /# Case2:環境変数の設定ファイルを使用（ルートディレクトリに`.env`ファイルを作成）

   ```.envファイルの場合
   NOTES_PASSWORD="your_notes_password"
   NOTES_REDIRECT_BASE="http://xxx="
   PYTHONPATH="src"
   ```

2. プログラム実行設定：なくても使えますが、作成すると便利です。

   **DBマニュフェスト(`dbs_manifest.json`)**
   - 対象DBを複数登録しておくことができる実施リスト。
   - これを指定すると**複数のDBを連続で出力することが可能**です。
   - `manifest/dbs_manifest_sample.json` を参考に作成し、対象のDBを定義します。

   **実行プロファイル (`profile.json`)**
   - CLIの引数(arg)のセットを複数定義し、名称を付けて管理が可能です。
   - また、マニュフェストも引数の一つとして設定可能です。
   - `profiles/profile_samples.json` を参考に作成し、実行したい処理内容を定義します。

3. オプション設定 (config.yaml): 必要に応じて設定が可能です。プロジェクトルートに config.yaml を作成し、環境に合わせて編集します。
   /# 使用する日本語フォントのパス (Windowsの例)  
   font_path: "C:/Windows/Fonts/NotoSansJP-VF.ttf"  
   font_family: "Noto Sans JP"

   /# Notes接続情報 (オプション - 環境変数やプロファイルで上書き可能)  
   notes_username: ""  # 例: "CN=User Name/O=Organization"  
   notes_password: ""  # 直接書く代わりに環境変数 NOTES_PASSWORD を推奨  
   notes_id_file: ""   /# IDファイルのパス (必要な場合)

## 出力ディレクトリ

```
/<OUT_PATH>/<DBタイトル名>/<フォーム名>/<カテゴリ1>_<カテゴリ2>_<カテゴリ3>/Doc_<作成日>_<タイトル>/
  ├─ Doc_<作成日>_<タイトル>.html                # メタ＋本文＋応答＋添付リンク
  ├─ Doc_<作成日>_<タイトル>.normalized.json     # 中間ファイル
  └─ attachments/                               # 添付ファイル一式（相対リンク先）
```

- `<作成日>` は `YYYYMMDD`
- 文字種は `pathvalidate` でサニタイズ（禁則文字は `_` 置換、末尾ドット除去、長さ上限あり）
- カテゴリの区切りは `"/"`, `">"`, `"\"` を同等扱いし、最大 **3 階層**

**出力先の設定**

- 後述の「使い方」は、"相対パス"を例にを記載していますが、**フルパスでの指定を推奨**します。
- **boxドライブの場合、ローカルキャッシュが最大25GBにデフォルト設定**されている為、ローカルドライブを圧縮するリスクがあります。**一度boxドライブからログアウト**することで**キャッシュファイルの削除が可能**です。
- **最終保管先がboxの場合、添付ファイルのメタ情報（更新日）を維持するため、直接boxドライブの出力先を指定することを推奨**します。

## 使い方

プロジェクトの実行は **すべて `src/app/main.py` から** 行います。

**参考**

以下、CLI(コマンドラインインターフェース)での使用を前提としていますが、
VSCodeの環境があれば、1~3の設定を`.vscode/launch.json`に定義することで、
「実行とデバック」から、選択して実行することが可能です。
サンプル(`.vscode/launch_sample.json`)も用意しています。

**パス記載の注意点**

パスは全て/（スラッシュ）で表記してください。
\（バックスラッシュ）は “エスケープ文字” として使われるため、リテラルのバックスラッシュを表したい場合は \\\ と 2 回書く必要があります。

---

### 1. 統合実行 (Unified Execution)

単一のコマンドで、DXLエクスポートからレンダリングまでの一連の処理を実行します。

#### ● 単一DB実行（マニュフェスト不使用）

単一のNotesデータベースを処理します。

##### ローカルの場合

```powershell
python src/app/main.py run-single-db `
    --db "C:/Path/To/Your/Database.nsf" `
    --out dist/output/single `
    --state dist/state/single `
    --formats html
```

##### サーバー上の場合(以降省略)

- `--server "YourServer/Org"`: **サーバー名の指定が必要**です。

```powershell
python src/app/main.py run-single-db `
    --server "YourServer/Org" `
    --db "To/Your/Database.nsf" `
    --out dist/output/single `
    --state dist/state/single `
    --formats html
```

#### ● マニフェスト実行（複数のDBを処理）

`manifest.json` ファイルに基づき、複数のデータベースを一括で処理します。

```powershell
python src/app/main.py run-manifest `
    --manifest manifest/dbs_manifest.json `
    --out dist/output `
    --state dist/state `
    --formats html `
    --no-keep-dxl
```

**マニフェスト JSON スキーマ (配列)**
各項目:

- title (文字列) — データベース表示タイトル（**保存するフォルダ名**となります）
- server (文字列) — Domino サーバー、またはローカルの場合は ""
- db_file (文字列) — NSF パス。ローカルの場合はフルパスを設定
- view_name (文字列、**オプション**) — プライマリビュー名
- views (配列[文字列]、**オプション**) — 順番にトライするビュー名のリスト
`view_name` と `views` のどちらも指定されていない場合、パイプラインは組み込みのフォールバックを使用します:`["($All)", "AllDocuments", "All Documents", "すべての文書", "全ての文書", "すべてのドキュメント"]`そして最後に NoteCollection / AllDocuments を使用します。

---

### 2. プロファイル実行 (`run-profile`)

`profiles/` 以下に定義された設定で処理を実行します。

#### ● プロファイル名を指定して実行

`profiles\profile.json` 内の `run-manifest` という名前の構成を実行する例です。

```powershell
python src/app/main.py run-profile --profiles profiles\profile.json --name "run-manifest"
```

#### ● 対話的にプロファイルを選択

`--name` を省略すると、利用可能なプロファイルのリストから選択できます。

```powershell
python src/app/main.py run-profile --profiles profiles\profile.json
```

---

### 3. 分離実行 (Decoupled Execution)

処理を各ステップに分割して実行します。

#### ● Step 1: DXL エクスポート (`export`)

NotesデータベースからDXLファイルをエクスポートします。

```powershell
python src/app/main.py export `
    --db "C:/Path/To/Your/Database.nsf" `
    --dxl-out dist/_dxl_export
```

#### ● Step 2: 正規化 (`normalize`)

DXLファイルを中間JSON形式に変換します。

```powershell
python src/app/main.py normalize `
    --dxl-dir dist/_dxl_export `
    --json-dir dist/json `
    --db-title "MyDatabase"
```

#### ● Step 3: 添付ファイル抽出 (`extract`)

DXLとJSONを元に添付ファイルを抽出し、JSON内のパス情報を更新します。

```powershell
python src/app/main.py extract `
    --dxl-dir dist/_dxl_export `
    --json-dir dist/json `
    --attach-dir dist/attachments
```

#### ● Step 4: レンダリング (`render`)

更新されたJSONファイルを指定されたフォーマット（例: HTML）で出力します。

```powershell
python src/app/main.py render `
    --json-dir dist/json `
    --out dist/output `
    --formats html
```

---

## ディレクトリ構造 (主要)

- manifest/: 設定ファイル (dbs_manifest.json など)  
- profiles/: 実行プロファイル (profile.json など)
- src/: ソースコード  
  - app/main.py: 単一の実行エントリーポイント
  - core/: 中核となる処理（DXL解析、レンダリング、添付抽出など）  
  - pipelines/:  
    - flows.py: Notes連携を含む実行フローライブラリ  
  - utils/: 共通ユーティリティ  
- dist/: 生成される成果物の出力先 (デフォルト)  
  - output/: レンダリングされたドキュメント (HTML, MD, DOCX, PDF)  
  - attachments/: (extractコマンド使用時) 抽出された添付ファイル  
  - json/: (normalizeコマンド使用時) 生成された normalized.json ファイル  
  - state/: 進捗管理ファイル (progress.jsonl)  
  - reports/: 処理結果のサマリレポート (JSON, CSV)  
- _dxl/: (exportコマンド使用時) DXLファイルが保存されるディレクトリ

## 既知の課題・制限事項

- 非常に複雑なネスト構造を持つリッチテキストや特殊なOLEオブジェクトの再現には限界がある場合があります。

## ライセンス

MITライセンス
