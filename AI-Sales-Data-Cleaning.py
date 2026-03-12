import streamlit as st
import pandas as pd
import re
import json
from io import BytesIO

st.set_page_config(page_title="Excel 清洗工具", layout="centered")
st.title("📄 Excel 槽位 / 卡片 / Intent 清洗工具")

# ==================================================
# 上传 Excel
# ==================================================
uploaded_file = st.file_uploader(
    "请上传 Excel 文件（.xlsx）",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("⬆️ 请先上传 Excel 文件")
    st.stop()

df = pd.read_excel(uploaded_file, dtype=object)

st.success(f"已加载文件，共 {len(df)} 行，{len(df.columns)} 列")

# ==================================================
# 1. wrapper 剔除
# ==================================================
remove_literals = [
    '[{"messageType":"answer","groupInfo":{"groupId":0,"groupStep":"","groupContentType":""},"content":{"type":"markdown","cardType":"","cardId":"","message":"',
    '"}}]',
    '[{"messageType":"question","groupInfo":{"groupId":0,"groupStep":"","groupContentType":""},"content":{"type":"json","cardType":"question","cardId":"","message":"{\\"message\\":\\"',
    ',\\"imageList\\":[],\\"aiScene\\":\\"aiSeller\\"}',
    ',\\"imageList\\":null,\\"aiScene\\":\\"aiSeller\\"}',
]

# ==================================================
# 2. 基础替换
# ==================================================
replace_map = {
    '[{"messageType":"answer","groupInfo":{"groupId":0,"groupStep":"","groupContentType":""},"content":{"type":"json","cardType":"clue002"':
        "【电话授权卡片】",
    '[{"messageType":"answer","groupInfo":{"groupId":0,"groupStep":"","groupContentType":""},"content":{"type":"json","cardType":"slotQuickQuestion"':
        "【快捷回复卡片】",
    '[{"messageType":"answer","groupInfo":{"groupId":0,"groupStep":"","groupContentType":""},"content":{"type":"json","cardType":"greetQuickQuestion"':
        "【欢迎语卡片】",
    '[{"messageType":"answer","groupInfo":{"groupId":0,"groupStep":"","groupContentType":""},"content":{"type":"json","cardType":"manualService"':
        "【转人工卡片】",
    '{"intention":"意图识别失败，多次尝试也是失败，默认下发回答助手问题意图","isLegal":true,"tool":null}':
        "意图识别失败",
}

# ==================================================
# 3. slotQuickQuestion
# ==================================================
slot_pattern = re.compile(
    r'"cardType"\s*:\s*"slotQuickQuestion".*?"cardId"\s*:\s*"(?P<cid>[^"]+)".*?"message"\s*:\s*"(?P<msg>(?:\\.|[^"\\])*)"',
    flags=re.DOTALL
)

def transform_slot(s):
    def repl(m):
        try:
            # 解析 JSON
            obj = json.loads(m.group('msg').replace('\\"', '"'))
            qs = [x['question'] for x in obj.get('questionList', [])]
            
            # 先拼接 qs
            joined_qs = "\\".join(qs)
            
            # 放入 f-string
            return f'【快捷回复卡片】,"cardId":"{m.group("cid")}","message":"{joined_qs}"'
        except Exception:
            return m.group(0)
    
    return slot_pattern.sub(repl, s)

# ==================================================
# 4. 通用卡片
# ==================================================
def simple_card(s, card_type, label, keep_card_id=True):
    pattern = re.compile(
        rf'"cardType"\s*:\s*"{card_type}".*?"cardId"\s*:\s*"(?P<cid>[^"]+)".*?"message"\s*:\s*"(?P<msg>(?:\\.|[^"\\])*)"',
        flags=re.DOTALL
    )
    def repl(m):
        msg = m.group('msg').replace('\\"', '"')
        if keep_card_id:
            return f'【{label}】"cardId":"{m.group("cid")}","message":"{msg}"'
        else:
            return f'【{label}】"message":"{msg}"'
    return pattern.sub(repl, s)

# ==================================================
# 5. risk
# ==================================================
risk_pattern = re.compile(
    r'\[\{\s*"messageType"\s*:\s*"risk".*?"message"\s*:\s*"(?P<msg>[^"]+)',
    flags=re.DOTALL
)

def transform_risk(s):
    if not isinstance(s, str):
        return s
    return risk_pattern.sub(lambda m: m.group('msg').strip(), s)

# ==================================================
# 6. intent
# ==================================================
intent_pattern = re.compile(
    r'{"intention"\s*:\s*"(?P<intention>.*?)"\s*,\s*"isLegal"\s*:\s*true\s*,\s*"tool"\s*:\s*null\s*}',
    flags=re.DOTALL
)

def transform_intent(s):
    def repl(m):
        raw = m.group('intention')
        cleaned = raw.replace('\\n', '').replace('\\r', '')
        cleaned = cleaned.replace('\\"', '"').replace('\\', '')
        try:
            obj = json.loads(cleaned)
            return ",".join(
                it.get('意图') for it in obj.get('意图列表', []) if it.get('意图')
            )
        except:
            return m.group(0)
    return intent_pattern.sub(repl, s)

# ==================================================
# 7. 兜底清洗
# ==================================================
tail_pattern = re.compile(r'",\s*"imageList"\s*:\s*null.*?$')
slot_head_pattern = re.compile(
    r'\[\{\s*"messageType"\s*:\s*"answer".*?"content"\s*:\s*\{\s*"type"\s*:\s*"json"\s*,',
    flags=re.DOTALL
)

def final_cleanup(s):
    s = slot_head_pattern.sub('', s)
    s = tail_pattern.sub('', s)
    return s

def postprocess(s):
    if not isinstance(s, str):
        return s
    s = s.replace('\\\\\"', '"').replace('\\"', '"').replace('\\\\', '\\')
    return s.strip(' "\'')

# ==================================================
# 8. 主清洗函数
# ==================================================
def clean_cell(val):
    if pd.isna(val):
        return val

    s = str(val)
    s = transform_risk(s)
    s = transform_slot(s)
    s = simple_card(s, "oneImage", "品宣卡片")
    s = simple_card(s, "package", "套餐卡片", False)
    s = transform_intent(s)

    for k, v in replace_map.items():
        s = s.replace(k, v)

    for w in remove_literals:
        s = s.replace(w, "")

    s = final_cleanup(s)
    return postprocess(s)

# ==================================================
# 9. 执行 & 导出
# ==================================================
if st.button("🚀 开始清洗"):
    with st.spinner("正在清洗中..."):
        for col in df.columns:
            df[col] = df[col].apply(clean_cell)

    buffer = BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    st.success("✅ 清洗完成")

    st.download_button(
        label="⬇️ 下载清洗后的 Excel",
        data=buffer,
        file_name="cleaned_final_v4.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
