import os
import time
from pywinauto import Application, Desktop

# Wordファイルのパス
word_file_path = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.docx'
pdf_file_path = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.pdf'

# Wordアプリケーションを開く
app = Application(backend="uia").start('winword.exe "{}"'.format(word_file_path))
time.sleep(3)

# Wordウィンドウを取得
word_window = app.window(title_re=".* - Word")

# ファイルメニューを開く
word_window.child_window(title="ファイル", control_type="MenuItem").click_input()
time.sleep(1)

# 印刷を選択
word_window.child_window(title="印刷", control_type="ListItem").click_input()
time.sleep(2)

# Microsoft Print to PDFを選択
print_combobox = word_window.child_window(title="プリンター", control_type="ComboBox")
print_combobox.select("Microsoft Print to PDF")

# 印刷ボタンをクリック
word_window.child_window(title="印刷", control_type="Button").click_input()

# ファイル名と保存先を指定
save_dialog = Desktop(backend="uia").window(title_re=".*名前を付けて保存")
save_dialog.child_window(title="ファイル名(F):", control_type="Edit").set_edit_text(pdf_file_path)
save_dialog.child_window(title="保存", control_type="Button").click_input()

# Wordアプリケーションを閉じる
word_window.close()
