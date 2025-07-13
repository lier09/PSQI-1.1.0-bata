import pandas as pd
import re
from pypinyin import pinyin, Style
from typing import List, Dict, Optional
import streamlit as st

# --- Helper Parser Functions ---
def parse_frequency(text: str) -> int:
    s = str(text or '').lower().replace('/', '').replace('／', '')
    if '无' in s or 'none' in s or s == '': return 0
    if '<1' in s or '＜1' in s or 'less than once' in s: return 1
    if '1-2' in s or '1–2' in s: return 2
    if '>=3' in s or '>或=3' in s or '≥3' in s or '3 or more' in s: return 3
    return 0

def parse_minutes(text: str) -> int:
    match = re.search(r'(\d+)', str(text or ''))
    return int(match.group(1)) if match else 0

def parse_hours(text: str) -> float:
    match = re.search(r'(\d+(\.\d+)?)', str(text or ''))
    return float(match.group(1)) if match else 0.0

def parse_time(text: str) -> Dict[str, int]:
    cleaned = str(text or '').replace('点', ':').replace('分', '').replace(' ', '')
    parts = cleaned.split(':')
    hour = int(parts[0]) if parts[0].isdigit() else 0
    minute = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
    return {'hour': hour, 'minute': minute}

def map_sum_to_component_score(total_sum: int, thresholds: List[int]) -> int:
    if total_sum == 0: return 0
    if total_sum <= thresholds[0]: return 1
    if total_sum <= thresholds[1]: return 2
    return 3

def calculate_component_1(q6: str) -> int:
    s = str(q6 or '').strip()
    if '很好' in s: return 0
    if '较好' in s: return 1
    if '较差' in s: return 2
    if '很差' in s: return 3
    return 0

def calculate_component_2(q2: str, q5a: str) -> int:
    time_to_sleep = parse_minutes(q2)
    score_q2 = 0
    if time_to_sleep <= 15: score_q2 = 0
    elif time_to_sleep <= 30: score_q2 = 1
    elif time_to_sleep <= 60: score_q2 = 2
    else: score_q2 = 3
    score_q5a = parse_frequency(q5a)
    return map_sum_to_component_score(score_q2 + score_q5a, [2, 4])

def calculate_component_3(q4: str) -> int:
    sleep_hours = parse_hours(q4)
    if sleep_hours > 7: return 0
    if sleep_hours >= 6: return 1
    if sleep_hours >= 5: return 2
    return 3

def calculate_component_4(q1: str, q3: str, q4: str) -> int:
    bed_time = parse_time(q1)
    wake_time = parse_time(q3)
    bed_time_in_minutes = bed_time['hour'] * 60 + bed_time['minute']
    wake_time_in_minutes = wake_time['hour'] * 60 + wake_time['minute']
    if wake_time_in_minutes <= bed_time_in_minutes:
        wake_time_in_minutes += 24 * 60
    time_in_bed_minutes = wake_time_in_minutes - bed_time_in_minutes
    if time_in_bed_minutes <= 0: return 3
    actual_sleep_minutes = parse_hours(q4) * 60
    if actual_sleep_minutes <= 0: return 3
    efficiency = (actual_sleep_minutes / time_in_bed_minutes) * 100
    if efficiency >= 85: return 0
    if efficiency >= 75: return 1
    if efficiency >= 65: return 2
    return 3

def calculate_component_5(row: pd.Series) -> int:
    disturbances = [row.get(f'q5{c}', '') for c in ['b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j']]
    total_sum = sum(parse_frequency(text) for text in disturbances)
    return map_sum_to_component_score(total_sum, [9, 18])

def calculate_component_6(q7: str) -> int:
    return parse_frequency(q7)

def calculate_component_7(q8: str, q9: str) -> int:
    score_q8 = parse_frequency(q8)
    score_q9 = 0
    s = str(q9 or '').strip()
    if '无' in s: score_q9 = 0
    elif '偶尔' in s: score_q9 = 1
    elif '有时' in s: score_q9 = 2
    elif '经常' in s: score_q9 = 3
    return map_sum_to_component_score(score_q8 + score_q9, [2, 4])

def get_pinyin_key(name: str) -> str:
    if not name or not isinstance(name, str): return ''
    return "".join(item[0] for item in pinyin(name.strip(), style=Style.NORMAL))

# --- Streamlit App ---
st.title("PSQI 评分工具（交互式上传版）")

st.markdown("""
✅ **请上传包含问卷条目的 Excel 文件**  
✅ **请注意上传文件的列名需符合以下范式：**  
序号 所用时间 填写日期是： 姓名： 您的年龄 1 2 3 4 5a. 5b. 5c 5d. 5e. 5f. 5g. 5h. 5i. 5j. 6 7 8 9  
示例：`2024-11-29    尹珅    18    23点    10分    6点    6.5小时    无 ...`
""")

uploaded_file = st.file_uploader("上传 Excel 文件", type=["xlsx"])

name_order_text = st.text_area("（可选）输入姓名排序顺序，每行一个姓名，用于自定义结果排序：", height=150)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("文件读取成功，预览数据：")
    st.dataframe(df.head())

    if st.button("开始计算 PSQI 分数"):
        results = []
        for _, row in df.iterrows():
            scores = {
                'C1_SleepQuality': calculate_component_1(row.get('6')),
                'C2_SleepLatency': calculate_component_2(row.get('2'), row.get('5a.')),
                'C3_SleepDuration': calculate_component_3(row.get('4')),
                'C4_SleepEfficiency': calculate_component_4(row.get('1'), row.get('3'), row.get('4')),
                'C5_SleepDisturbances': calculate_component_5(row),
                'C6_MedicationUse': calculate_component_6(row.get('7')),
                'C7_DaytimeDysfunction': calculate_component_7(row.get('8'), row.get('9')),
            }
            total_score = sum(scores.values())
            results.append({
                'Name': row.get('姓名：'),
                'Age': row.get('您的年龄'),
                **scores,
                'TotalScore': total_score
            })

        results_df = pd.DataFrame(results)

        if name_order_text.strip():
            sort_order_names = [name.strip() for name in name_order_text.strip().split('\n') if name.strip()]
            pinyin_order_map = {get_pinyin_key(name): i for i, name in enumerate(sort_order_names)}
            results_df['sort_key'] = results_df['Name'].apply(lambda n: pinyin_order_map.get(get_pinyin_key(n), float('inf')))
            results_df = results_df.sort_values(by='sort_key').drop(columns='sort_key')

        st.success(f"✅ 成功处理 {len(results_df)} 条记录。")
        st.dataframe(results_df)

        @st.cache_data
        def convert_df(df):
            return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')

        csv = convert_df(results_df)
        st.download_button("📥 下载结果 CSV 文件", csv, "psqi_results.csv", "text/csv", key='download-csv')

st.markdown("---")
with open("睡眠质量量表计算方式.pdf", "rb") as f:
    st.download_button("📄 下载 PSQI 计算方法说明（PDF）", f, file_name="睡眠质量量表计算方式.pdf")

st.markdown("*Developed by Dr. Huze Peng, Capital University of Physical Education and Sports.*")
