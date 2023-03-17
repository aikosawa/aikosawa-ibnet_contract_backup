
from datetime import datetime, date
from datetime import timezone, timedelta
from typing import Union
import platform


JST = timezone(timedelta(hours=9), 'JST')


def strftime(dt: Union[datetime, date], format_str: str) -> str:
    """
    OSによってstrftimeで0埋めをしない場合の実装が違うので、その差分を吸収するためのヘルパー関数
    """

    os_type = platform.system()
    if os_type == 'Windows':
        format_str = format_str.replace('%-', '%#')
    else:
        format_str = format_str.replace('%#', '%-')

    return dt.strftime(format_str)
