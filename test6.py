def has_spread_word(template_name: str) -> bool:
    return "見開き" in template_name

input_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_① Cedar GA_70N.docx'
output_file = r'C:\Users\Gou\Downloads\03_金銭消費貸借契約証書_20211007_渡辺　健太_見開き① Cedar GA_70N.pdf'

# 例: テンプレート名に "見開き" が含まれている場合
template_name1 = "見開きテンプレート.docx"
result1 = has_spread_word(output_file)
print(result1)  # True

# 例: テンプレート名に "見開き" が含まれていない場合
template_name2 = "通常テンプレート.docx"
result2 = has_spread_word(input_file)
print(result2)  # False

