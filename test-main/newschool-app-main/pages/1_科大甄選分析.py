import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import time

st.set_page_config(page_title="新頁面", page_icon="🆕")

st.title("🆕 112-113 各校錄取分數線比較分析")
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
        total_weight = (
            weights['國文加權'] +
            weights[' 英文加權'] +
            weights[' 數學加權'] +
            weights[' 專業(一)加權'] +
            weights[' 專業(二)加權']
        )
        weighted_sum = (
            float(row['國文分數']) * weights['國文加權'] +
            float(row['英文分數']) * weights[' 英文加權'] +
            float(row['數學B分數']) * weights[' 數學加權'] +
            float(row['專一分數']) * weights[' 專業(一)加權'] +
            float(row['專二分數']) * weights[' 專業(二)加權']
        )
        return weighted_sum / total_weight if total_weight != 0 else None
    except (ValueError, KeyError, ZeroDivisionError):
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

# 讀取 11309a (1).xlsx 的平均欄位
school_avg = df_113[['學校名稱', '系科組學程名稱', '平均']].drop_duplicates()
school_weights = pd.merge(school_weights, school_avg, on=['學校名稱', '系科組學程名稱'], how='left')

# 計算每個學生的加權分數並找出可上的最好學校
results = []
for _, student in df_113g.iterrows():
    student_scores = []
    for _, school in school_weights.iterrows():
        weighted_score = calculate_weighted_score(student, school)
        if weighted_score is not None:
            student_scores.append({
                '學校名稱': school['學校名稱'],
                '系科組學程名稱': school['系科組學程名稱'],
                '加權平均分數': weighted_score,
                '錄取分數': school['錄取總分數'],
                '平均': school['平均'],
                '是否可錄取': weighted_score >= school['錄取總分數']
            })
    # 找出可錄取的平均最高的學校
    possible_schools = [s for s in student_scores if s['是否可錄取']]
    if possible_schools:
        best_school = max(possible_schools, key=lambda x: x['平均'] if pd.notnull(x['平均']) else float('-inf'))
        results.append({
            '座號': student['座號'],
            '班級': student['班級'],
            '國文分數': student['國文分數'],
            '英文分數': student['英文分數'],
            '數學B分數': student['數學B分數'],
            '專一分數': student['專一分數'],
            '專二分數': student['專二分數'],
            '錄取學校': student.get('錄取學校', ''),
            '錄取校系': student.get('錄取校系', ''),
            '最佳可錄取學校': best_school['學校名稱'],
            '最佳可錄取科系': best_school['系科組學程名稱'],
            '加權平均分數': best_school['加權平均分數'],
            '該校錄取分數': best_school['錄取分數'],
            '最佳平均': best_school['平均']
        })

# 轉換為DataFrame並排序
results_df = pd.DataFrame(results)
if not results_df.empty:
    # 新增是否同科系欄位，並排序
    results_df['是否同科系'] = results_df['錄取校系'] == results_df['最佳可錄取科系']
    results_df = results_df.sort_values('是否同科系', ascending=False)

# 顯示結果
st.subheader("學生加權分數與最佳可錄取學校")
if not results_df.empty:
    st.dataframe(results_df[[
        '座號', '班級', '國文分數', '英文分數', '數學B分數', '專一分數', '專二分數',
        '錄取學校', '錄取校系',
        '最佳可錄取學校', '最佳可錄取科系', '加權平均分數', '該校錄取分數', '最佳平均'
    ]])
else:
    st.dataframe(results_df)

# 比較原本錄取校系與最佳可錄取校系的平均
st.subheader("原本錄取學校與最佳可錄取學校比較")
comparison_data = []
for _, student in df_113g.iterrows():
    original_school = student.get('錄取學校', '未錄取')
    original_dept = student.get('錄取校系', '')
    best_row = results_df[results_df['座號'] == student['座號']]
    best_school = best_row['最佳可錄取學校'].iloc[0] if not best_row.empty else '未找到'
    best_dept = best_row['最佳可錄取科系'].iloc[0] if not best_row.empty else ''
    # 取原本錄取校系的平均
    original_avg = None
    best_avg = None
    try:
        original_avg = df_113[(df_113['學校名稱'] == original_school) & (df_113['系科組學程名稱'] == original_dept)]['平均'].iloc[0]
    except:
        pass
    try:
        best_avg = df_113[(df_113['學校名稱'] == best_school) & (df_113['系科組學程名稱'] == best_dept)]['平均'].iloc[0]
    except:
        pass
    is_better = (best_avg is not None and original_avg is not None and best_avg > original_avg)
    comparison_data.append({
        '座號': student['座號'],
        '原本錄取學校': original_school,
        '原本錄取校系': original_dept,
        '最佳可錄取學校': best_school,
        '最佳可錄取校系': best_dept,
        '原本平均': original_avg,
        '最佳平均': best_avg,
        '是否更好': is_better
    })
comparison_df = pd.DataFrame(comparison_data)
st.subheader("詳細比較")
comparison_df = comparison_df.sort_values('座號')
st.dataframe(comparison_df[['座號', '原本錄取學校', '原本錄取校系', '最佳可錄取學校', '最佳可錄取校系', '原本平均', '最佳平均', '是否更好']])

# 顯示統計資訊
if not results_df.empty:
    st.subheader("統計資訊")
    st.write(f"總學生數：{len(results_df)}")
    st.write(f"平均加權平均分數：{results_df['加權平均分數'].mean():.2f}")
    st.write(f"最高加權平均分數：{results_df['加權平均分數'].max():.2f}")
    st.write(f"最低加權平均分數：{results_df['加權平均分數'].min():.2f}")
    
    # 顯示各校錄取人數統計
    school_stats = results_df['最佳可錄取學校'].value_counts()
    st.subheader("各校可錄取人數統計")
    st.bar_chart(school_stats)

    # 比較原本錄取學校與最佳可錄取學校
    st.subheader("原本錄取學校與最佳可錄取學校比較")
    
    # 創建比較數據
    comparison_data = []
    for _, student in df_113g.iterrows():
        original_school = student.get('錄取學校', '未錄取')
        best_school = results_df[results_df['座號'] == student['座號']]['最佳可錄取學校'].iloc[0] if not results_df[results_df['座號'] == student['座號']].empty else '未找到'
        
        comparison_data.append({
            '座號': student['座號'],
            '原本錄取學校': original_school,
            '最佳可錄取學校': best_school,
            '是否更好': best_school != original_school
        })
    
    comparison_df = pd.DataFrame(comparison_data)
    
    # 計算統計數據
    better_count = len(comparison_df[comparison_df['是否更好']])
    same_count = len(comparison_df[~comparison_df['是否更好']])
    
    # 創建圓餅圖
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 7))
    
    # 原本錄取學校分布
    original_schools = comparison_df['原本錄取學校'].value_counts()
    ax1.pie(original_schools, labels=original_schools.index, autopct='%1.1f%%')
    ax1.set_title('原本錄取學校分布')
    
    # 最佳可錄取學校分布
    best_schools = comparison_df['最佳可錄取學校'].value_counts()
    ax2.pie(best_schools, labels=best_schools.index, autopct='%1.1f%%')
    ax2.set_title('最佳可錄取學校分布')
    
    st.pyplot(fig)
    
    # 顯示比較結果
    st.subheader("比較結果")
    st.write(f"總學生數：{len(comparison_df)}")
    st.write(f"可以上更好學校的學生數：{better_count}")
    st.write(f"原本就是最佳選擇的學生數：{same_count}")
    
    # 添加比較結果的圓餅圖
    fig2, ax3 = plt.subplots(figsize=(8, 8))
    comparison_results = [better_count, same_count]
    labels = ['可以上更好學校', '原本就是最佳選擇']
    colors = ['#FF9999', '#66B2FF']
    ax3.pie(comparison_results, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
    ax3.set_title('學生選擇比較結果')
    st.pyplot(fig2)
    
    # 顯示詳細比較表格
    st.subheader("詳細比較")
    comparison_df = comparison_df.sort_values('座號')
    st.dataframe(comparison_df[['座號', '原本錄取學校', '最佳可錄取學校', '是否更好']])
    
    # 顯示可以上更好學校的學生名單
    if better_count > 0:
        st.subheader("可以上更好學校的學生名單")
        better_students = comparison_df[comparison_df['是否更好']]
        st.dataframe(better_students[['座號', '原本錄取學校', '最佳可錄取學校']])

    # 112年分析
    st.markdown("---")
    st.title("112學年度分析")
    
    # 計算112年加權分數
    def calculate_weighted_score_112(row, weights):
        try:
            total_weight = (
                weights['國文加權'] +
                weights[' 英文加權'] +
                weights[' 數學加權'] +
                weights[' 專業(一)加權'] +
                weights[' 專業(二)加權']
            )
            weighted_sum = (
                float(row['國文分數']) * weights['國文加權'] +
                float(row['英文分數']) * weights[' 英文加權'] +
                float(row['數學B分數']) * weights[' 數學加權'] +
                float(row['專一分數']) * weights[' 專業(一)加權'] +
                float(row['專二分數']) * weights[' 專業(二)加權']
            )
            return weighted_sum / total_weight if total_weight != 0 else None
        except (ValueError, KeyError, ZeroDivisionError):
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

    # 計算每個學生的加權分數並找出可上的最好學校
    results_112 = []
    for _, student in df_112g.iterrows():
        student_scores = []
        for _, school in school_weights_112.iterrows():
            weighted_score = calculate_weighted_score_112(student, school)
            if weighted_score is not None:
                student_scores.append({
                    '學校名稱': school['學校名稱'],
                    '系科組學程名稱': school['系科組學程名稱'],
                    '加權總分': weighted_score,
                    '錄取分數': school['錄取總分數'],
                    '是否可錄取': weighted_score >= school['錄取總分數']
                })
        
        # 找出可錄取的最高分學校
        possible_schools = [s for s in student_scores if s['是否可錄取']]
        if possible_schools:
            best_school = max(possible_schools, key=lambda x: x['加權總分'])
            results_112.append({
                '座號': student['座號'],
                '班級': student['班級'],
                '國文分數': student['國文分數'],
                '英文分數': student['英文分數'],
                '數學B分數': student['數學B分數'],
                '專一分數': student['專一分數'],
                '專二分數': student['專二分數'],
                '最佳可錄取學校': best_school['學校名稱'],
                '最佳可錄取科系': best_school['系科組學程名稱'],
                '加權總分': best_school['加權總分'],
                '該校錄取分數': best_school['錄取分數']
            })

    # 轉換為DataFrame並排序
    results_df_112 = pd.DataFrame(results_112)
    if not results_df_112.empty:
        results_df_112 = results_df_112.sort_values('加權總分', ascending=False)

    # 顯示112年結果
    st.subheader("112學年度學生加權分數與最佳可錄取學校")
    st.dataframe(results_df_112)

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

        # 比較原本錄取學校與最佳可錄取學校
        st.subheader("112學年度原本錄取學校與最佳可錄取學校比較")
        
        # 創建比較數據
        comparison_data_112 = []
        for _, student in df_112g.iterrows():
            original_school = student.get('錄取學校', '未錄取')
            best_school = results_df_112[results_df_112['座號'] == student['座號']]['最佳可錄取學校'].iloc[0] if not results_df_112[results_df_112['座號'] == student['座號']].empty else '未找到'
            
            comparison_data_112.append({
                '座號': student['座號'],
                '原本錄取學校': original_school,
                '最佳可錄取學校': best_school,
                '是否更好': best_school != original_school
            })
        
        comparison_df_112 = pd.DataFrame(comparison_data_112)
        
        # 計算統計數據
        better_count_112 = len(comparison_df_112[comparison_df_112['是否更好']])
        same_count_112 = len(comparison_df_112[~comparison_df_112['是否更好']])
        
        # 創建圓餅圖
        fig3, (ax4, ax5) = plt.subplots(1, 2, figsize=(15, 7))
        
        # 原本錄取學校分布
        original_schools_112 = comparison_df_112['原本錄取學校'].value_counts()
        ax4.pie(original_schools_112, labels=original_schools_112.index, autopct='%1.1f%%')
        ax4.set_title('112學年度原本錄取學校分布')
        
        # 最佳可錄取學校分布
        best_schools_112 = comparison_df_112['最佳可錄取學校'].value_counts()
        ax5.pie(best_schools_112, labels=best_schools_112.index, autopct='%1.1f%%')
        ax5.set_title('112學年度最佳可錄取學校分布')
        
        st.pyplot(fig3)
        
        # 顯示比較結果
        st.subheader("112學年度比較結果")
        st.write(f"總學生數：{len(comparison_df_112)}")
        st.write(f"可以上更好學校的學生數：{better_count_112}")
        st.write(f"原本就是最佳選擇的學生數：{same_count_112}")
        
        # 添加比較結果的圓餅圖
        fig4, ax6 = plt.subplots(figsize=(8, 8))
        comparison_results_112 = [better_count_112, same_count_112]
        labels_112 = ['可以上更好學校', '原本就是最佳選擇']
        colors_112 = ['#FF9999', '#66B2FF']
        ax6.pie(comparison_results_112, labels=labels_112, autopct='%1.1f%%', colors=colors_112, startangle=90)
        ax6.set_title('112學年度學生選擇比較結果')
        st.pyplot(fig4)
        
        # 顯示詳細比較表格
        st.subheader("112學年度詳細比較")
        comparison_df_112 = comparison_df_112.sort_values('座號')
        st.dataframe(comparison_df_112[['座號', '原本錄取學校', '最佳可錄取學校', '是否更好']])
        
        # 顯示可以上更好學校的學生名單
        if better_count_112 > 0:
            st.subheader("112學年度可以上更好學校的學生名單")
            better_students_112 = comparison_df_112[comparison_df_112['是否更好']]
            st.dataframe(better_students_112[['座號', '原本錄取學校', '最佳可錄取學校']]) 