import math
import os
import re
from datetime import datetime
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright

URL = "https://uel-app-portal.uel.ac.uk/urd/sits.urd/run/siw_ipp_lgn.login?process=siw_ipp_app&code2=0005&code1=CD2299FAD"

email = "songshuhan@rocky2028.ccwu.cc"

password = "12345678Aa!"
# 默认桌面上的学生表；可用环境变量 STUDENT_INFO_XLSX 覆盖绝对路径
_DEFAULT_XLSX = Path.home() / "Desktop" / "学生基础信息_已翻译.xlsx"

_EN_MON = (
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
)


def _dob_display_d_mon_yyyy(dt: datetime) -> str:
    """与 UEL 常见控件一致，例如 14/May/2026（英文月份缩写，日不补零）。"""
    return f"{dt.day}/{_EN_MON[dt.month - 1]}/{dt.year}"


_EXCEL_ID_FLOAT_MSG = (
    "从 Excel 读到的身份证号是「数值/浮点」或科学计数法（如 3.41e+17），"
    "会丢失末尾数字，无法可靠解析。\n"
    "请将 I 列整列设为「文本」格式，重新粘贴完整 18 位号码，保存表格后再运行。"
)


def _id_string_from_excel_cell(raw: object) -> str:
    """把 I 列单元格值规范成身份证字符串；拒绝浮点/科学计数法（Excel 数字列）。"""
    if raw is None:
        raise ValueError("身份证号为空")
    if isinstance(raw, bool):
        raise ValueError(f"身份证号单元格无效: {raw!r}")
    if isinstance(raw, float):
        if math.isnan(raw) or math.isinf(raw):
            raise ValueError("身份证号单元格为非有限数值")
        raise ValueError(_EXCEL_ID_FLOAT_MSG)
    if isinstance(raw, int) and not isinstance(raw, bool):
        s = str(raw)
    else:
        t = str(raw).strip()
        if "e" in t.lower():
            raise ValueError(_EXCEL_ID_FLOAT_MSG)
        s = "".join(t.split()).upper()
    s = re.sub(r"[^\dX]", "", s)
    if len(s) not in (15, 18):
        raise ValueError(
            f"身份证号长度须为 15 或 18 位（当前 {len(s)} 位）。"
            f"若表格里显示正常但仍报错，多半仍是数字格式，请把 I 列改为「文本」后重贴号码。"
        )
    return s


def _dob_yyyymmdd_from_chinese_id(raw: object) -> str:
    """从中国大陆身份证号取出生日 YYYYMMDD（18 位或 15 位）。"""
    s = _id_string_from_excel_cell(raw)
    if len(s) == 18:
        ymd = s[6:14]
        if not ymd.isdigit():
            raise ValueError(f"身份证号生日段非数字: {raw!r}")
        if not (s[17].isdigit() or s[17] == "X"):
            raise ValueError(f"身份证号末位无效: {raw!r}")
    elif len(s) == 15:
        if not s.isdigit():
            raise ValueError(f"15 位身份证号须全为数字: {raw!r}")
        yy, rest = s[6:8], s[8:12]
        y = int(yy)
        # 两位数年份的常见启发式（与旧 15 位证习惯一致）
        year = 2000 + y if y <= 39 else 1900 + y
        ymd = f"{year}{rest}"
    datetime.strptime(ymd, "%Y%m%d")  # 校验合法日期
    return ymd


def surname_forename_dob_from_xlsx(xlsx_path: Path) -> tuple[str, str, str]:
    """从第 1 个工作表同一行读取：G 列 Name（姓/名）、I 列身份证号（生日）。
    Name 按首个空白切成两段：前半为姓、后半为名；无空白则姓为空、整格作为名。
    生日由身份证号解析。默认输出 D/Mon/YYYY（如 14/May/2026）；若设置环境变量
    UEL_DOB_FORMAT（strftime 格式串），则改用该格式（例如 %d/%m/%Y）。"""
    fmt = os.getenv("UEL_DOB_FORMAT", "").strip()
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        for row in ws.iter_rows(min_row=2, min_col=7, max_col=9, values_only=True):
            g_cell, _h, i_cell = row[0], row[1], row[2]
            if g_cell is None:
                continue
            full = str(g_cell).strip()
            if not full:
                continue
            parts = full.split(None, 1)
            surname, forename = (
                (parts[0].strip(), parts[1].strip()) if len(parts) > 1 else ("", full)
            )
            ymd = _dob_yyyymmdd_from_chinese_id(i_cell)
            dt = datetime.strptime(ymd, "%Y%m%d")
            dob_display = dt.strftime(fmt) if fmt else _dob_display_d_mon_yyyy(dt)
            return surname, forename, dob_display
        raise ValueError(f"{xlsx_path} 的 G 列（自第 2 行起）没有可用的姓名")
    finally:
        wb.close()


def main() -> None:
    # 默认可视化打开；如需无头模式：set HEADLESS=1
    headless = os.getenv("HEADLESS", "").strip() in {"1", "true", "True", "yes", "YES"}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()
        page.goto(URL, wait_until="domcontentloaded")
        page.locator("#NEW_USER").click()

        # ID 含点号，CSS 里要转义；第 8 个 option 对应 select_option 的 index=7（从 0 起）
        title = page.locator(r"#TITLE\.IPU\.SRS")
        title.wait_for(state="visible")
        title.click()
        title.select_option(index=7)

        xlsx = Path(os.environ.get("STUDENT_INFO_XLSX", str(_DEFAULT_XLSX))).expanduser()
        surname_text, forename_text, dob_text = surname_forename_dob_from_xlsx(xlsx)

        surname = page.locator(r"#IPU_SURN\.IPU\.SRS")
        surname.wait_for(state="visible")
        surname.fill(surname_text)

        forename = page.locator(r"#IPU_FNM1\.IPU\.SRS")
        forename.wait_for(state="visible")
        forename.fill(forename_text)

        dob = page.locator(r"#IPU_DOB\.IPU\.SRS")
        dob.wait_for(state="visible")
        dob.fill(dob_text)

        usercode = page.locator(r"#USERCODE\.IPU\.SRS")
        usercode.wait_for(state="visible")
        usercode.fill(email)

        haem = page.locator(r"#IPU_HAEM\.IPU\.SRS")
        haem.wait_for(state="visible")
        haem.fill(email)

        pw1 = page.locator(r"#PASSWORD1\.IPU\.SRS")
        pw1.wait_for(state="visible")
        pw1.fill(password)

        pw2 = page.locator(r"#PASSWORD2\.IPU\.SRS")
        pw2.wait_for(state="visible")
        pw2.fill(password)

        proceed = page.locator(r"#PROCEED\.DUM1\.SRS")
        proceed.wait_for(state="visible")
        # proceed.click()

        input(">>> 结束，按 Enter 关闭浏览器...")
        context.close()
        browser.close()


if __name__ == "__main__":
    main()
