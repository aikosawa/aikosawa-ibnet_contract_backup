# すでにvenvがある場合には削除
if (Test-Path .\venv) {
    Remove-Item -Path .\venv -Recurse
}

# グローバルのパッケージを破壊しないために仮想環境を直下のvenvディレクトリに作成
py -m venv venv 

# 仮想環境の有効化
venv\Scripts\Activate.ps1

# 依存ライブラリのインストール
pip install -r requirements.txt