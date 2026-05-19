import math
import os

import pandas as pd
from deep_translator import GoogleTranslator

# Excel 第 I 列（第 9 列）：身份证号等长数字须按文本保存，避免科学计数法
ID_COLUMN_INDEX = 8  # pandas / 0-based
ID_COLUMN_OPENPYXL = 9  # openpyxl 列号 1-based，A=1 → I=9

# AF 列（常住址）→ AG 列（邮编），翻译前按浙江省内城市填充
ADDRESS_COLUMN_INDEX = 31  # AF，0-based
POSTCODE_COLUMN_INDEX = 32  # AG，0-based
DEFAULT_ZHEJIANG_POSTCODE = "315000"

# 浙江省地级市 → 主城区邮政编码（长名称优先匹配）
ZHEJIANG_CITY_POSTCODES: tuple[tuple[str, str], ...] = (
    ("绍兴市", "312000"),
    ("杭州市", "310000"),
    ("宁波市", "315000"),
    ("温州市", "325000"),
    ("湖州市", "313000"),
    ("嘉兴市", "314000"),
    ("金华市", "321000"),
    ("衢州市", "324000"),
    ("台州市", "318000"),
    ("丽水市", "323000"),
    ("舟山市", "316000"),
    ("绍兴", "312000"),
    ("杭州", "310000"),
    ("宁波", "315000"),
    ("温州", "325000"),
    ("湖州", "313000"),
    ("嘉兴", "314000"),
    ("金华", "321000"),
    ("衢州", "324000"),
    ("台州", "318000"),
    ("丽水", "323000"),
    ("舟山", "316000"),
)

# --- 1. 路径设置 ---
desktop_path = os.path.join(os.path.expanduser("~"), 'Desktop')
input_file_name = '学生基础信息.xlsx'
output_file_name = '学生基础信息_已翻译.xlsx'

input_path = os.path.join(desktop_path, input_file_name)
output_path = os.path.join(desktop_path, output_file_name)


def _persist_id_string(x) -> str:
    """I 列专用：不翻译，只规整为纯文本（去首尾空白等），便于 Excel 按文本保存。"""
    if isinstance(x, bool):
        return str(int(x))
    if x is None or pd.isna(x):
        return ""
    if isinstance(x, int):
        return str(x)
    if isinstance(x, float):
        if not math.isfinite(x):
            return ""
        s = str(x)
        if "e" in s.lower():
            print(
                "⚠️  I 列某格以科学计数法读入，源表可能已丢精度。"
                "请把「学生基础信息.xlsx」的 I 列设为「文本」后重新粘贴身份证号再翻译。"
            )
        if x.is_integer():
            return str(int(x))
        return s.strip()
    s = str(x).strip()
    return "" if s.lower() == "nan" else s


def _id_read_converter(x):
    """read_excel 第 9 列专用：尽早转成 str，减少 pandas 推断成 float。"""
    return _persist_id_string(x)


def infer_zhejiang_postcode(address) -> str:
    """根据地址文本推断浙江省内邮政编码，无法匹配时返回默认 315000。"""
    if address is None or (isinstance(address, float) and pd.isna(address)):
        return ""
    s = str(address).strip()
    if not s or s.lower() == "nan":
        return ""
    for city, code in ZHEJIANG_CITY_POSTCODES:
        if city in s:
            return code
    return DEFAULT_ZHEJIANG_POSTCODE


def fill_postcode_from_address(df: pd.DataFrame) -> int:
    """根据 AF 列地址填写 AG 列邮编（翻译前执行）。返回本次写入的行数。"""
    if len(df.columns) <= POSTCODE_COLUMN_INDEX:
        return 0
    addr_col = df.columns[ADDRESS_COLUMN_INDEX]
    post_col = df.columns[POSTCODE_COLUMN_INDEX]
    df[post_col] = df[post_col].astype(object)
    filled = 0
    for idx in df.index:
        addr = df.at[idx, addr_col]
        if addr is None or (isinstance(addr, float) and pd.isna(addr)):
            continue
        if str(addr).strip() == "" or str(addr).strip().lower() == "nan":
            continue
        current = df.at[idx, post_col]
        if current is not None and not (isinstance(current, float) and pd.isna(current)):
            cur_s = str(current).strip()
            if cur_s and cur_s.lower() != "nan":
                continue
        df.at[idx, post_col] = infer_zhejiang_postcode(addr)
        filled += 1
    return filled


def _apply_id_column_text_format(worksheet) -> None:
    """将 I 列设为 Excel「文本」格式（@），并保证单元格值为字符串。"""
    for row in range(1, worksheet.max_row + 1):
        cell = worksheet.cell(row=row, column=ID_COLUMN_OPENPYXL)
        v = cell.value
        if v is not None and v != "":
            cell.value = _persist_id_string(v)
        cell.number_format = "@"


# --- 2. 数据读取逻辑 ---
def load_data(path):
    try:
        lower = path.lower()
        if lower.endswith((".xlsx", ".xlsm")):
            xf = pd.ExcelFile(path, engine="openpyxl")
            head = pd.read_excel(xf, sheet_name=0, nrows=0)
            kwargs: dict = {"sheet_name": 0, "engine": "openpyxl"}
            if len(head.columns) > ID_COLUMN_INDEX:
                kwargs["converters"] = {ID_COLUMN_INDEX: _id_read_converter}
            return pd.read_excel(xf, **kwargs)
        if lower.endswith(".xls"):
            head = pd.read_excel(path, nrows=0)
            kwargs = {}
            if len(head.columns) > ID_COLUMN_INDEX:
                kwargs["converters"] = {ID_COLUMN_INDEX: _id_read_converter}
            return pd.read_excel(path, **kwargs) if kwargs else pd.read_excel(path)
        return pd.read_excel(path)
    except Exception:
        encodings = ['utf-8-sig', 'gb18030', 'utf-8', 'gbk']
        for enc in encodings:
            try:
                return pd.read_csv(path, encoding=enc)
            except:
                continue
    raise Exception("无法读取文件，请检查文件路径或格式。")


# --- 3. 核心翻译逻辑 ---
def start_process():
    if not os.path.exists(input_path):
        print(f"❌ 找不到文件: {input_path}")
        return

    try:
        df = load_data(input_path)
        print("✅ 文件加载成功！")
    except Exception as e:
        print(f"❌ {e}")
        return

    if len(df.columns) > POSTCODE_COLUMN_INDEX:
        n = fill_postcode_from_address(df)
        addr_name = df.columns[ADDRESS_COLUMN_INDEX]
        post_name = df.columns[POSTCODE_COLUMN_INDEX]
        print(f"📮 已根据「{addr_name}」(AF) 填写「{post_name}」(AG) 邮编：{n} 行（浙江省内，默认 {DEFAULT_ZHEJIANG_POSTCODE}）")

    translator = GoogleTranslator(source='auto', target='en')

    def safe_translate(text):
        """数字或已是英文等则原样返回；含中文等可译内容则调用 Google 翻译。"""
        val = str(text).strip()
        if not val or val.lower() == 'nan' or val.isdigit() or len(val) < 1:
            return text
        try:
            return translator.translate(val)
        except:
            return text

    print("🚀 开始全自动翻译...")

    # A. 翻译表头
    df.columns = [safe_translate(col) for col in df.columns]

    # B. 普通列：safe_translate。I 列（身份证）不进入翻译，循环结束后只做一次文本规整。
    for i, column in enumerate(df.columns):
        print(f"进度: {i + 1}/{len(df.columns)} - 正在处理: {column}")
        if i == ID_COLUMN_INDEX:
            continue
        df[column] = df[column].apply(safe_translate)

    if len(df.columns) > ID_COLUMN_INDEX:
        ic = df.columns[ID_COLUMN_INDEX]
        print(f"  → I 列「{ic}」不翻译，仅 _persist_id_string 规整为纯文本")
        df[ic] = df[ic].map(_persist_id_string)

    # --- 4. 保存并美化 (自动调整列宽) ---
    print("🎨 正在优化表格样式...")
    try:
        # 使用 ExcelWriter 来控制样式
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Translated')

            # 获取当前工作表对象
            worksheet = writer.sheets['Translated']
            if len(df.columns) > ID_COLUMN_INDEX:
                _apply_id_column_text_format(worksheet)

            # 自动调整列宽逻辑
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter  # 获取列字母

                for cell in col:
                    try:
                        if cell.value:
                            # 计算单元格内容的长度
                            length = len(str(cell.value))
                            if length > max_length:
                                max_length = length
                    except:
                        pass

                # 设置列宽：基础长度 + 2 个字符的缓冲
                # 英文环境 1.2 倍系数比较合适
                adjusted_width = (max_length + 2) * 1.2
                # 限制最大宽度，防止某个单元格内容过多导致列宽无限大
                worksheet.column_dimensions[column].width = min(adjusted_width, 50)

            if len(df.columns) > ID_COLUMN_INDEX:
                letter_i = worksheet.cell(row=1, column=ID_COLUMN_OPENPYXL).column_letter
                cur = worksheet.column_dimensions[letter_i].width or 0
                worksheet.column_dimensions[letter_i].width = max(cur, 22)

        print("\n" + "=" * 40)
        print(f"✨ 任务完成！")
        print(f"📂 翻译并美化后的文件已保存至桌面: {output_file_name}")
        print("=" * 40)

    except Exception as e:
        print(f"❌ 保存失败: {e}")


if __name__ == "__main__":
    start_process()