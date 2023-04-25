import os
import time
import win32com.client
from pywinauto import Application, timings

def check_dialog_exists(word_dialog, dialog_title):
    word_dialog.print_control_identifiers()
    return word_dialog.child_window(title=dialog_title).exists()

def check_pdf_dialog_not_exists(pdf_dialog):
    return not pdf_dialog.exists()

def docx_to_pdf(input_file, output_file):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True

    try:
        doc = word.Documents.Open(input_file)

        # 1枚に2ページ印刷するための設定
        word.ActiveDocument.PageSetup.TwoPagesOnOne = True

        # PDF仮想プリンタを指定 (ここでは、"Microsoft Print to PDF"を使用)
        word.ActivePrinter = "Microsoft Print to PDF"

        # 印刷ダイアログを開く
        doc.PrintOut()

        # pywinautoアプリケーションオブジェクトを取得
        # app = Application(backend="uia").connect(process=r"Artificial Intelligence (AI) Host for the Microsoft® Windows® Operating System and Platform x64.")
        app = Application(backend="uia")
        word_dialog = app.window(title=dialog_title)

        word_dialog = app.top_window()

        # "Microsoft Print to PDF"ダイアログのタイトル
        dialog_title = "Microsoft Print to PDF"

        # ダイアログが表示されるまで待機
        timings.wait_until_passes(10, 0.5, lambda: check_dialog_exists(word_dialog, dialog_title))

        # 保存先とファイル名を入力
        pdf_dialog = word_dialog.child_window(title="SaveFileDialog") 
        file_edit = pdf_dialog.child_window(auto_id="1148")  # ファイル名入力フィールドのID
        file_edit.set_text(output_file)

        # 保存ボタンをクリック
        save_button = pdf_dialog.child_window(title="Save")  # 保存ボタンのタイトル
        save_button.click()

        # ダイアログが閉じるまで待機
        timings.wait_until_passes(10, 0.5, lambda: check_pdf_dialog_not_exists(pdf_dialog))

        doc.Close()

    finally:
        word.Quit()

input_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.docx'
output_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.pdf'

docx_to_pdf(input_file, output_file)
