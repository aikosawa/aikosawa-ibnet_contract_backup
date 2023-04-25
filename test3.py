import os
import win32print
import win32api
import win32com.client as win32

def word_to_pdf_2_pages_per_sheet(input_file, output_file):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False

    # デフォルトプリンタ名を取得
    default_printer = win32print.GetDefaultPrinter()

    # PDF出力先を指定
    # output_filename = os.path.join(output_path, output_file)

    # Word文書を開く
    doc = word.Documents.Open(input_file)

    # 印刷設定を変更
    word.ActivePrinter = default_printer
    word.PrintOut(
        OutputFileName=output_file,
        Item=win32.constants.wdPrintDocumentContent,
        Copies=1,
        Pages="1-2",
        Collate=True,
        Background=False,
        PrintToFile=True,
        Range=win32.constants.wdPrintAllDocument,
        ManualDuplexPrint=False,
        PrintZoomColumn=2,  # 1ページに2ページ分を印刷
        PrintZoomRow=1,
        PrintZoomPaperWidth=0,
        PrintZoomPaperHeight=0
    )

    # Word文書を閉じる
    doc.Close(SaveChanges=False)
    word.Quit()

# Wordファイルと出力設定
input_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.docx'
output_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.pdf'
# output_file = r'03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.pdf'

# Wordを2ページ単位でPDFに変換
word_to_pdf_2_pages_per_sheet(input_file, output_file)
