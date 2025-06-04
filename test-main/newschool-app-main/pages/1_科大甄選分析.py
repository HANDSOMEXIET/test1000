import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import time

st.set_page_config(page_title="科大甄選分析", page_icon="🎓")

st.title("🎓 112-113 各校錄取分數線比較分析")
st.write("本頁比較 112 與 113 學年各校錄取分數線的變化與分布。")

def read_excel_with_retry(file_path, sheet_name, max_retries=3):
    for attempt in range(max_retries):
        try:
            return pd.read_excel(file_path, sheet_name=sheet_name)
        except PermissionError:
            if attempt < max_retries - 1:
                st.warning(f"無法讀取文件 {file_path}，正在重試... (嘗試 {attempt + 1}/{max_retries})")
                time.sleep(2)  # 等待2秒後重試
            else:
                raise
        except Exception as e:
            raise

# 讀取資料
try:
    # 使用相對路徑，回到上一層目錄
    base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # 顯示正在讀取的文件路徑（用於調試）
    st.write("正在讀取文件路徑：", base_path)
    
    # 讀取文件
    df_113 = read_excel_with_retry(os.path.join(base_path, "11309a (1).xlsx"), "Sheet1")
    df_112 = read_excel_with_retry(os.path.join(base_path, "11209.xlsx"), "工作表1")
    df_113g = read_excel_with_retry(os.path.join(base_path, "113科大甄選.xlsx"), "工作表1")
    df_112g = read_excel_with_retry(os.path.join(base_path, "112科大甄選.xlsx"), "工作表1")
    
except Exception as e:
    st.error(f"資料讀取失敗: {e}")
    st.error("請確保：")
    st.error("1. Excel 文件沒有被其他程式打開")
    st.error("2. 您有權限訪問這些文件")
    st.error("3. 文件路徑正確")
    st.stop()

# 只保留學校名稱與錄取分數線
col_school = '學校名稱'
col_score = '錄取總分數'

# 轉型
for df in [df_113, df_112]:
    if col_score in df.columns:
        df[col_score] = pd.to_numeric(df[col_score], errors='coerce')
    else:
        st.error(f"找不到 '{col_score}' 欄位，請檢查資料格式。")
        st.stop()

# 取每校最高分數線（有些學校可能有多科系，這裡以最高分為代表）
school_113 = df_113.groupby(col_school)[col_score].max().reset_index().rename(columns={col_score: '113分數線'})
school_112 = df_112.groupby(col_school)[col_score].max().reset_index().rename(columns={col_score: '112分數線'})

# 合併
compare_df = pd.merge(school_113, school_112, on=col_school, how='outer')
compare_df['分數線變化'] = compare_df['113分數線'] - compare_df['112分數線']

# 排名
compare_df['113排名'] = compare_df['113分數線'].rank(ascending=False, method='min')
compare_df['112排名'] = compare_df['112分數線'].rank(ascending=False, method='min')
compare_df = compare_df.sort_values('113排名')

st.subheader("各校錄取分數線排名與變化")
st.dataframe(compare_df[[col_school, '113分數線', '113排名', '112分數線', '112排名', '分數線變化']].reset_index(drop=True))

# 分數線變化圖
st.subheader("分數線變化圖 (113 - 112)")
fig, ax = plt.subplots(figsize=(10, 5))
ax.bar(compare_df[col_school], compare_df['分數線變化'], color=['#4CAF50' if x >= 0 else '#F44336' for x in compare_df['分數線變化']])
ax.set_ylabel('分數線變化')
ax.set_xlabel('學校名稱')
ax.set_title('各校錄取分數線變化 (113 - 112)')
ax.tick_params(axis='x', labelrotation=90)
st.pyplot(fig)

# 分布圖
st.subheader("錄取分數線分布圖")
fig2, ax2 = plt.subplots(figsize=(8, 4))
ax2.hist(compare_df['112分數線'].dropna(), bins=30, alpha=0.5, label='112', color='#4CAF50')
ax2.hist(compare_df['113分數線'].dropna(), bins=30, alpha=0.5, label='113', color='#2196F3')
ax2.set_xlabel('錄取分數線')
ax2.set_ylabel('學校數')
ax2.set_title('112/113 各校錄取分數線分布')
ax2.legend()
st.pyplot(fig2)

# 顯示113科大甄選資料
st.subheader("113學年度科大甄選資料")

# 計算加權分數
def calculate_weighted_score(row, weights):
    try:
        # 計算加權總分
        weighted_total = (
            float(row['國文分數']) * weights['國文加權'] +
            float(row['英文分數']) * weights[' 英文加權'] +
            float(row['數學B分數']) * weights[' 數學加權'] +
            float(row['專一分數']) * weights[' 專業(一)加權'] +
            float(row['專二分數']) * weights[' 專業(二)加權']
        )
        
        return weighted_total
    except (ValueError, KeyError):
        return None

# 讀取各校系加權資料
school_weights = df_113.groupby(['學校名稱', '系科組學程名稱']).agg({
    '國文加權': 'first',
    ' 英文加權': 'first',
    ' 數學加權': 'first',
    ' 專業(一)加權': 'first',
    ' 專業(二)加權': 'first',
    '錄取總分數': 'first'
}).reset_index()

# 計算每個學生的加權分數並找出可上的最好學校
results = []
# 計算所有學校的平均錄取分數
avg_admission_score = school_weights['錄取總分數'].mean()

for _, student in df_113g.iterrows():
    student_scores = []
    for _, school in school_weights.iterrows():
        weighted_score = calculate_weighted_score(student, school)
        if weighted_score is not None:
            total_weight = (
                school['國文加權'] + school[' 英文加權'] + school[' 數學加權'] +
                school[' 專業(一)加權'] + school[' 專業(二)加權']
            )
            weighted_avg = weighted_score / total_weight if total_weight != 0 else None
            # 取得該校該科系的11309a (1).xlsx的平均
            school_row = df_113[(df_113['學校名稱'] == school['學校名稱']) & (df_113['系科組學程名稱'] == school['系科組學程名稱'])]
            school_mean = school_row['平均'].iloc[0] if not school_row.empty else None
            student_scores.append({
                '學校名稱': school['學校名稱'],
                '系科組學程名稱': school['系科組學程名稱'],
                '加權總分': weighted_score,
                '加權平均': weighted_avg,
                '錄取分數': school['錄取總分數'],
                '平均': school_mean,
                '是否可錄取': weighted_avg is not None and school_mean is not None and weighted_avg >= school_mean
            })
    # 找出加權平均大於等於該校平均的所有學校
    possible_schools = [s for s in student_scores if s['是否可錄取']]
    if possible_schools:
        best_school = max(possible_schools, key=lambda x: x['平均'])
    else:
        continue
    results.append({
        '座號': student['座號'],
        '班級': student['班級'],
        '國文分數': student['國文分數'],
        '英文分數': student['英文分數'],
        '數學B分數': student['數學B分數'],
        '專一分數': student['專一分數'],
        '專二分數': student['專二分數'],
        '原本錄取學校': student.get('錄取學校', '未錄取'),
        '原本錄取校系': student.get('錄取校系', '未錄取'),
        '原本錄取分數': school_weights[
            (school_weights['學校名稱'] == student.get('錄取學校', '')) & 
            (school_weights['系科組學程名稱'] == student.get('錄取校系', ''))
        ]['錄取總分數'].iloc[0] if not school_weights[
            (school_weights['學校名稱'] == student.get('錄取學校', '')) & 
            (school_weights['系科組學程名稱'] == student.get('錄取校系', ''))
        ].empty else '未找到',
        '原本錄取校系平均': (
            df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))]['平均'].iloc[0]
            if not df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))].empty else '未找到'
        ),
        '最佳可錄取學校': best_school['學校名稱'],
        '最佳可錄取科系': best_school['系科組學程名稱'],
        '加權總分': best_school['加權總分'],
        '加權平均': best_school['加權平均'],
        '該校錄取分數': best_school['錄取分數'],
        '最佳校系平均是否較高': (
            True if (isinstance(best_school['平均'], (int, float)) and isinstance(
                df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))]['平均'].iloc[0] if not df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))].empty else None, (int, float))
                and best_school['平均'] > (
                    df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))]['平均'].iloc[0]
                    if not df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))].empty else None
                )
            ) else False if (isinstance(best_school['平均'], (int, float)) and isinstance(
                df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))]['平均'].iloc[0] if not df_113[(df_113['學校名稱'] == student.get('錄取學校', '')) & (df_113['系科組學程名稱'] == student.get('錄取校系', ''))].empty else None, (int, float))) else '未找到'
        )
    })

# 轉換為DataFrame並排序
results_df = pd.DataFrame(results)
if not results_df.empty:
    results_df = results_df.sort_values('加權總分', ascending=False)

# 顯示結果
st.subheader("學生加權分數與最佳可錄取學校")
st.dataframe(results_df[['座號', '班級', '國文分數', '英文分數', '數學B分數', '專一分數', '專二分數', 
                        '原本錄取學校', '原本錄取校系', '原本錄取分數', '原本錄取校系平均', '最佳可錄取學校', '最佳可錄取科系', 
                        '加權總分', '加權平均', '該校錄取分數', '最佳校系平均是否較高']])

# 顯示統計資訊
if not results_df.empty:
    st.subheader("統計資訊")
    st.write(f"總學生數：{len(results_df)}")
    st.write(f"平均加權總分：{results_df['加權總分'].mean():.2f}")
    st.write(f"最高加權總分：{results_df['加權總分'].max():.2f}")
    st.write(f"最低加權總分：{results_df['加權總分'].min():.2f}")
    
    # 顯示各校錄取人數統計
    school_stats = results_df['最佳可錄取學校'].value_counts()
    st.subheader("各校可錄取人數統計")
    st.bar_chart(school_stats)

    # 比較原本錄取學校與最佳可錄取學校（依據最佳校系平均是否較高）
    st.subheader("原本錄取學校與最佳可錄取學校比較（依據最佳校系平均是否較高）")
    better_count = (results_df['最佳校系平均是否較高'] == True).sum()
    same_count = (results_df['最佳校系平均是否較高'] == False).sum()
    unknown_count = (results_df['最佳校系平均是否較高'] == '未找到').sum()
    st.write(f"可以上更好學校的學生數：{better_count}")
    st.write(f"原本就是最佳選擇的學生數：{same_count}")
    st.write(f"無法比較的學生數：{unknown_count}")
    
    # 添加比較結果的圓餅圖
    fig2, ax3 = plt.subplots(figsize=(8, 8))
    comparison_results = [better_count, same_count, unknown_count]
    labels = ['可以上更好學校', '原本就是最佳選擇', '無法比較']
    colors = ['#FF9999', '#66B2FF', '#CCCCCC']
    ax3.pie(comparison_results, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
    ax3.set_title('學生選擇比較結果')
    st.pyplot(fig2)
    
    # 顯示詳細比較表格
    st.subheader("詳細比較")
    st.dataframe(results_df[['座號', '原本錄取學校', '原本錄取校系', '最佳可錄取學校', '最佳可錄取科系', '原本錄取校系平均', '加權平均', '最佳校系平均是否較高']].sort_values('座號'))
    
    # 顯示可以上更好學校的學生名單
    if better_count > 0:
        st.subheader("可以上更好學校的學生名單")
        better_students = results_df[results_df['最佳校系平均是否較高'] == True]
        st.dataframe(better_students[['座號', '原本錄取學校', '原本錄取校系', '最佳可錄取學校', '最佳可錄取科系', '原本錄取校系平均', '加權平均']])

# 112年分析
st.markdown("---")
st.title("112學年度分析")

# 計算112年加權分數
def calculate_weighted_score_112(row, weights):
    try:
        # 計算加權總分
        weighted_total = (
            float(row['國文分數']) * weights['國文加權'] +
            float(row['英文分數']) * weights[' 英文加權'] +
            float(row['數學B分數']) * weights[' 數學加權'] +
            float(row['專一分數']) * weights[' 專業(一)加權'] +
            float(row['專二分數']) * weights[' 專業(二)加權']
        )
        
        return weighted_total
    except (ValueError, KeyError):
        return None

# 讀取112年各校系加權資料
school_weights_112 = df_112.groupby(['學校名稱', '系科組學程名稱']).agg({
    '國文加權': 'first',
    ' 英文加權': 'first',
    ' 數學加權': 'first',
    ' 專業(一)加權': 'first',
    ' 專業(二)加權': 'first',
    '錄取總分數': 'first'
}).reset_index()

# 計算每個學生的加權分數並找出可上的最好學校（以加權平均與校系平均比對）
results_112 = []
for _, student in df_112g.iterrows():
    student_scores = []
    for _, school in school_weights_112.iterrows():
        try:
            weighted_score = calculate_weighted_score_112(student, school)
            total_weight = (
                school['國文加權'] + school[' 英文加權'] + school[' 數學加權'] +
                school[' 專業(一)加權'] + school[' 專業(二)加權']
            )
            weighted_avg = weighted_score / total_weight if total_weight != 0 else None
            school_row = df_112[(df_112['學校名稱'] == school['學校名稱']) & (df_112['系科組學程名稱'] == school['系科組學程名稱'])]
            school_mean = school_row['平均'].iloc[0] if not school_row.empty else None
            student_scores.append({
                '學校名稱': school['學校名稱'],
                '系科組學程名稱': school['系科組學程名稱'],
                '加權總分': weighted_score,
                '加權平均': weighted_avg,
                '平均': school_mean,
                '錄取分數': school['錄取總分數'],
                '是否可錄取': weighted_avg is not None and school_mean is not None and weighted_avg >= school_mean
            })
        except Exception as e:
            continue
    possible_schools = [s for s in student_scores if s['是否可錄取']]
    if possible_schools:
        best_school = max(possible_schools, key=lambda x: x['平均'])
    else:
        continue
    # 原本錄取校系平均
    orig_mean = (
        df_112[(df_112['學校名稱'] == student.get('錄取學校', '')) & (df_112['系科組學程名稱'] == student.get('錄取校系', ''))]['平均'].iloc[0]
        if not df_112[(df_112['學校名稱'] == student.get('錄取學校', '')) & (df_112['系科組學程名稱'] == student.get('錄取校系', ''))].empty else '未找到'
    )
    # 最佳校系平均是否較高
    if isinstance(best_school['平均'], (int, float)) and isinstance(orig_mean, (int, float)):
        is_better = best_school['平均'] > orig_mean
    elif best_school['平均'] == '未找到' or orig_mean == '未找到':
        is_better = '未找到'
    else:
        is_better = False
    results_112.append({
        '座號': student['座號'],
        '班級': student['班級'],
        '國文分數': student['國文分數'],
        '英文分數': student['英文分數'],
        '數學B分數': student['數學B分數'],
        '專一分數': student['專一分數'],
        '專二分數': student['專二分數'],
        '原本錄取學校': student.get('錄取學校', '未錄取'),
        '原本錄取校系': student.get('錄取校系', '未錄取'),
        '原本錄取分數': school_weights_112[
            (school_weights_112['學校名稱'] == student.get('錄取學校', '')) & 
            (school_weights_112['系科組學程名稱'] == student.get('錄取校系', ''))
        ]['錄取總分數'].iloc[0] if not school_weights_112[
            (school_weights_112['學校名稱'] == student.get('錄取學校', '')) & 
            (school_weights_112['系科組學程名稱'] == student.get('錄取校系', ''))
        ].empty else '未找到',
        '原本錄取校系平均': orig_mean,
        '最佳可錄取學校': best_school['學校名稱'],
        '最佳可錄取科系': best_school['系科組學程名稱'],
        '加權總分': best_school['加權總分'],
        '加權平均': best_school['加權平均'],
        '該校錄取分數': best_school['錄取分數'],
        '最佳校系平均': best_school['平均'],
        '最佳校系平均是否較高': is_better
    })
# 轉換為DataFrame並排序
results_df_112 = pd.DataFrame(results_112)
if not results_df_112.empty:
    results_df_112 = results_df_112.sort_values('加權總分', ascending=False)

# 顯示112年結果
st.subheader("112學年度學生加權分數與最佳可錄取學校")
st.dataframe(results_df_112[['座號', '班級', '國文分數', '英文分數', '數學B分數', '專一分數', '專二分數',
    '原本錄取學校', '原本錄取校系', '原本錄取分數', '原本錄取校系平均', '最佳可錄取學校', '最佳可錄取科系',
    '加權總分', '加權平均', '該校錄取分數', '最佳校系平均', '最佳校系平均是否較高']])

# 顯示112年統計資訊
if not results_df_112.empty:
    st.subheader("112學年度統計資訊")
    st.write(f"總學生數：{len(results_df_112)}")
    st.write(f"平均加權總分：{results_df_112['加權總分'].mean():.2f}")
    st.write(f"最高加權總分：{results_df_112['加權總分'].max():.2f}")
    st.write(f"最低加權總分：{results_df_112['加權總分'].min():.2f}")
    
    # 顯示各校錄取人數統計
    school_stats_112 = results_df_112['最佳可錄取學校'].value_counts()
    st.subheader("112學年度各校可錄取人數統計")
    st.bar_chart(school_stats_112)

    # 比較原本錄取學校與最佳可錄取學校（依據最佳校系平均是否較高）
    st.subheader("112學年度原本錄取學校與最佳可錄取學校比較（依據最佳校系平均是否較高）")
    better_count_112 = (results_df_112['最佳校系平均是否較高'] == True).sum()
    same_count_112 = (results_df_112['最佳校系平均是否較高'] == False).sum()
    unknown_count_112 = (results_df_112['最佳校系平均是否較高'] == '未找到').sum()
    st.write(f"可以上更好學校的學生數：{better_count_112}")
    st.write(f"原本就是最佳選擇的學生數：{same_count_112}")
    st.write(f"無法比較的學生數：{unknown_count_112}")
    
    # 添加比較結果的圓餅圖
    fig4, ax6 = plt.subplots(figsize=(8, 8))
    comparison_results_112 = [better_count_112, same_count_112, unknown_count_112]
    labels_112 = ['可以上更好學校', '原本就是最佳選擇', '無法比較']
    colors_112 = ['#FF9999', '#66B2FF', '#CCCCCC']
    ax6.pie(comparison_results_112, labels=labels_112, autopct='%1.1f%%', colors=colors_112, startangle=90)
    ax6.set_title('112學年度學生選擇比較結果')
    st.pyplot(fig4)
    
    # 顯示詳細比較表格
    st.subheader("112學年度詳細比較")
    st.dataframe(results_df_112[['座號', '原本錄取學校', '原本錄取校系', '最佳可錄取學校', '最佳可錄取科系', '原本錄取校系平均', '加權平均', '最佳校系平均', '最佳校系平均是否較高']].sort_values('座號'))
    
    # 顯示可以上更好學校的學生名單
    if better_count_112 > 0:
        st.subheader("112學年度可以上更好學校的學生名單")
        better_students_112 = results_df_112[results_df_112['最佳校系平均是否較高'] == True]
        st.dataframe(better_students_112[['座號', '原本錄取學校', '原本錄取校系', '最佳可錄取學校', '最佳可錄取科系', '原本錄取校系平均', '加權平均']]) 