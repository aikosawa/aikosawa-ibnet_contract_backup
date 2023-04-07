# ドラッグ & ドロップで受け付けたパスをダイアログで表示する
Add-Type -Assembly System.Windows.Forms
[System.Windows.Forms.MessageBox]::Show($args[0], "ファイルのパス")

# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 仮想環境の有効化
venv\Scripts\Activate.ps1


# 環境変数の設定
# $env:CONFIG_FILE_PATH = "設定情報.xlsxが格納されているファイルのフルパスを記述"
# $env:TEMPLATE_FOLDER_PATH = "\\fileserver02\IBNet銀座\非公開用\★IBN 重要★\申込書類正式\文書管理\★システムテンプレート★\templates"
# $env:OUTPUT_FOLDER_PATH = "出力先のフォルダのフルパスを記述"


# ドラッグ & ドロップで受け付けたファイルをmain.pyに渡して起動
py main.py $args[0]
# py main.py 'c:\Users\Gou\Downloads\個人入力データ - 20230130 .xlsx'