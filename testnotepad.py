import time
from pywinauto import Application, timings


def check_dialog_exists(word_dialog, dialog_title):
    word_dialog.print_control_identifiers()
    return word_dialog.child_window(title=dialog_title).exists()

# メモ帳を開く
app = Application().start("notepad.exe")

# メモ帳のウィンドウを取得
notepad_window = app.UntitledNotepad

# ウィンドウが表示されるまで待機
notepad_window.wait("visible")

# メモ帳にテキストを入力
notepad_window.Edit.type_keys("test")
notepad_window.print_control_identifiers()

# メニューから「名前を付けて保存」を選択
notepad_window.menu_select("ファイル(&F)->名前を付けて保存(&A)")


# "名前をつけて保存"ダイアログのタイトル
dialog_title = "名前をつけて保存"

# ダイアログが表示されるまで待機
# timings.wait_until_passes(20, 0.5, lambda: check_dialog_exists(notepad_window, dialog_title))
time.sleep(5)  # 5秒まつ
notepad_window.print_control_identifiers()

# 保存先とファイル名を入力
save_dialog = notepad_window.child_window(title=dialog_title) 


# # # 保存ダイアログを取得
# save_dialog = app.top_window()

# 保存先のファイル名を入力 (ここでは例として 'test.txt' という名前で保存)
save_dialog.Edit.set_edit_text("test.txt")

# 保存ボタンをクリック
save_dialog.Button.click()

# ウィンドウが閉じるのを待つ
app.UntitledNotepad.wait_not("visible")

# メモ帳を閉じる
notepad_window.close()
