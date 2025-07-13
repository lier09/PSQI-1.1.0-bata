import streamlit as st
import pandas as pd
import re
from io import BytesIO
from pypinyin import pinyin, Style

# --- Helper Parser Functions ---
def parse_frequency(text: str) -> int:
    s = str(text or '').lower().strip().replace('/', '').replace('／', '')
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

def parse_time(text: str) -> dict:
    cleaned = str(text or '').replace('点', ':').replace('分', '').replace(' ', '').strip()
    parts = cleaned.split(':')
    hour = int(parts[0]) if parts[0].isdigit() else 0
    minute = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
    return {'hour': hour, 'minute': minute}

def map_sum_to_component_score(total_sum: int, thresholds: list) -> int:
    if total_sum == 0: return 0
    if total_sum <= thresholds[0]: return 1
    if total_sum <= thresholds[1]: return 2
    return 3

# --- Component Calculator Functions ---
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
    disturbances = [row.get(f'q5{c}', '') for c in 'bcdefghij']
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

# --- Streamlit App ---
st.title("🛌 PSQI 批量计算工具 (自动匹配列名)")

uploaded_file = st.file_uploader("请上传含问卷条目的 Excel 文件", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("✅ 文件读取成功，预览数据：")
    st.dataframe(df.head())

    # 使用您之前定义的标准列名自动覆盖
    column_names = [
        'id', 'timeTaken', 'date', 'name', 'age', 'q1', 'q2', 'q3', 'q4',
        'q5a', 'q5b', 'q5c', 'q5d', 'q5e', 'q5f', 'q5g', 'q5h', 'q5i', 'q5j',
        'q6', 'q7', 'q8', 'q9'
    ]
    if len(df.columns) >= len(column_names):
        df.columns = column_names[:len(df.columns)]

    if st.button("开始计算 PSQI 分数"):
        results = []
        for _, row in df.iterrows():
            scores = {
                'C1_SleepQuality': calculate_component_1(row.get('q6')),
                'C2_SleepLatency': calculate_component_2(row.get('q2'), row.get('q5a')),
                'C3_SleepDuration': calculate_component_3(row.get('q4')),
                'C4_SleepEfficiency': calculate_component_4(row.get('q1'), row.get('q3'), row.get('q4')),
                'C5_SleepDisturbances': calculate_component_5(row),
                'C6_MedicationUse': calculate_component_6(row.get('q7')),
                'C7_DaytimeDysfunction': calculate_component_7(row.get('q8'), row.get('q9')),
            }
            total_score = sum(scores.values())
            results.append({
                'Name': row.get('name'),
                'Age': row.get('age'),
                **scores,
                'TotalScore': total_score
            })

        results_df = pd.DataFrame(results)
        st.success(f"🎉 成功处理 {len(results_df)} 条记录。")
        st.dataframe(results_df)

        output = BytesIO()
        results_df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="📥 下载结果 Excel 文件",
            data=output,
            file_name="psqi_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
