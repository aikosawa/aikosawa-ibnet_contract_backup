import os
import win32print
import win32com.client as win32


def word_to_pdf(input_file, output_file, two_pages_per_sheet=False):
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = False

    # デフォルトプリンタ名を取得
    default_printer = win32print.GetDefaultPrinter()

    # Word文書を開く
    doc = word.Documents.Open(input_file)

    if two_pages_per_sheet:
        # 1枚に2ページ単位で印刷する設定
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
            PrintZoomColumn=2,
            PrintZoomRow=1,
            PrintZoomPaperWidth=0,
            PrintZoomPaperHeight=0
        )
    else:
        # 通常のPDF化を行う
        doc.SaveAs(output_file, FileFormat=win32.constants.wdFormatPDF)

    # Word文書を閉じる
    doc.Close(SaveChanges=False)
    word.Quit()

# Wordファイルと出力設定
input_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.docx'
output_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.pdf'


# # 通常のPDF化
# word_to_pdf(input_file, output_file, two_pages_per_sheet=False)

# 1枚に2ページ単位でPDF化
word_to_pdf(input_file, output_file, two_pages_per_sheet=True)
