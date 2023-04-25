import win32com.client

def save_word_to_pdf_default(word_file, pdf_file):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False

    doc = word.Documents.Open(word_file)

    # word.Options.SaveInterval = True
    # word.ActiveDocument.PageSetup.BookFoldPrinting = True
    # word.ActiveDocument.PageSetup.BookFoldPrintingSheets = 1
    
    # デフォルトの印刷設定を利用してPDFに変換
    doc.ExportAsFixedFormat(pdf_file, ExportFormat=17)  # 17 is the code for PDF format
    
    doc.Close()
    word.Quit()

# Word文書のパス
word_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.docx'
# 変換後のPDFファイルのパス
pdf_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.pdf'


save_word_to_pdf_default(word_file, pdf_file)