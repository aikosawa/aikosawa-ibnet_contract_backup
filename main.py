from app import jikkin_sheet, config, chohyo_gen
from logging import getLogger, basicConfig
from typing import Any, Mapping
import logging
import openpyxl
import os
import re
import sys
import time

logger = getLogger(__name__)


def logging_keywords(title: str, kv: Mapping[str, Any]):
    logger.debug('')
    logger.debug(f'  +--- {title} ---------------')
    for k, v in kv.items():
        logger.debug(f'  | {k!r} => {type(v)} {v!r}')
    logger.debug('  +' + '-' * 40)
    logger.debug('')


def busy_wait_excel_open_and_close(filename: str, duration: int = 1):
    """
    duration(default: 1)秒間隔でexcelで開いたときのロックファイルの存在を確認し、
    excelで開いて閉じるまでこの関数の呼び出しがスレッドをブロックします。
    """
    lockfile_name = re.sub(r"(?=[^\/]+?\.\w+$)", "~$", filename, count=1)

    opened = False
    while True:
        time.sleep(duration)
        isopen = os.path.exists(lockfile_name)

        if opened and not isopen:
            break
        opened = isopen


def main():

    for i, arg in enumerate(sys.argv):
        logger.debug(f'{i}, {arg}')

    if len(sys.argv) <= 1:
        raise RuntimeError(
            '入力シートが引数に渡されておりません。引数に入力シートのパスを指定する必要があります。詳しくはREADME.mdをご参照ください。')

    input_file_path = os.path.normpath(sys.argv[1])
    config_file_path = os.path.normpath(os.environ.get(
        'CONFIG_FILE_PATH', './workdir/設定情報.xlsx'))
    template_path = os.path.normpath(
        os.environ.get('TEMPLATE_FOLDER_PATH', './templates'))
    output_path = os.path.normpath(
        os.environ.get('OUTPUT_FOLDER_PATH', './workdir/output'))

    cfg = config.load_config_from_workbook(config_file_path)

    chohyo_generator = chohyo_gen.ChohyoGenerator(
        cfg,
        template_path,
        output_path
    )

    logger.debug(f'---- 環境変数 ----')
    logger.debug('')
    logger.debug(f'  設定情報.xlsxの場所 {config_file_path}')
    logger.debug(f'  テンプレートが格納されているパス: {template_path}')
    logger.debug(f'  出力先のパス: {output_path}')
    logger.debug('')

    input_workbook = openpyxl.load_workbook(input_file_path, data_only=True)

    product_inputs = list(jikkin_sheet.get_keywords_for_each_products(
        input_workbook['入力シート']))

    for i, p in enumerate(product_inputs):
        logging_keywords(f'{i}列目の入力値', p.product_kv)

    chohyo_generator.gen_all_doc(product_inputs)

    logger.info('')
    logger.info('すべてのファイルの出力が完了しました。Enterキーを押してプログラムを終了します。')
    logger.info('')

    sys.stdin.flush()
    input()


if __name__ == '__main__':
    level = basicConfig(level=logging._nameToLevel.get(
        os.environ.get('LOG_LEVEL', '').upper(),
        logging.INFO)
    )
    main()
