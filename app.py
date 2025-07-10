# 健康管理ツール(Python+Streamlit/plotly) 

# 必要なモジュールをインポート
import pandas as pd
import numpy as np
import streamlit as st
import matplotlib.pyplot as plt
from matplotlib import rcParams
import matplotlib.font_manager as fm

#matplotlib.gridspecモジュールからGridSpec関数を直接インポート
from matplotlib.gridspec import GridSpec 

# 表示条件の入力
name = st.sidebar.selectbox('あなたの名前', ("孝則","由香"),
                            index=None, placeholder="名前を選択")
year = st.sidebar.selectbox('測定した年',
    ("2024","2025","2026","2027","2028","2029","2030","2031","2032","2033","2034","2035"),
     index=None, placeholder="年を選択")
#month = st.sidebar.text_input('グラフ表示する月')
month = st.sidebar.selectbox('グラフ表示する月', 
            ("1","2","3","4","5","6","7","8","9","10","11","12"),
            index=None, placeholder="月を選択")

# 実行ボタン
exec_btn =  st.sidebar.button("実行")

# 実行ボタンが押された時の処理
if exec_btn:

    try:
        # Dataフォルダの設定
        data_dir = "Data"

        # 家庭内健康管理データのExcelファイル名を設定
        if name == "孝則":
            name_e = "Taka"
        else:
            name_e = "Yuka"

        excel_file = data_dir + "/" + name_e + "/" + year + ".xlsx"

        # 家庭内健康管理データの読み込み
        df = pd.read_excel(excel_file, sheet_name = month)
        # 開始位置を設定
        start_index = 2
        for row in df.values:
            if row[1] == "最低血圧(起床時)":
                row2 = row[2:]
                for val in row2:
                    if (val is np.nan):
                        start_index = start_index + 1
                    else:
                        break
                break

        # グラフ化する値を設定
        y3_set = False
        y4_set = False
        y5_set = False

        for row in df.values:
            if row[1] == "検査項目":
                x1 = row[start_index:] #日をセット
                continue
            if row[1] == "最低血圧(起床時)":
                y1 = row[start_index:]
                continue
            if row[1] == "最高血圧(起床時)":
                y2 = row[start_index:]
                continue
            if row[1] == "体温（おでこ）" and y3_set == False:
                y3 = row[start_index:]
                y3_set = True
                continue                
            if row[1] == "酸素濃度(%Sp02)" and y4_set == False:
                y4 = row[start_index:]
                y4_set = True
                continue
            if row[1] == "脈拍数(PRbpm)" and y5_set == False:
                y5 = row[start_index:]
                y5_set = True
                continue
            if row[1] == "体重(kg)":
                yt1 = row[start_index:]
                continue
            if row[1] == "体脂肪率(％)":
                yt2 = row[start_index:]
                continue
            if row[1] == "筋肉量(kg)":
                yt3 = row[start_index:]
                continue                
            if row[1] == "推定骨量(kg)":
                yt4 = row[start_index:]
                continue
            if row[1] == "内臓脂肪(レべル)":
                yt5 = row[start_index:]
                continue
            if row[1] == "基礎代謝(kcal)":
                yt6 = row[start_index:]
                continue            
            if row[1] == "歩数（ヘルスケアで計測)":
                yt7 = row[start_index:]

        # キャシュのクリア
        st.cache_data.clear()

        # 日本語フォントを指定 (Windows版専用のため、使用を中止)
        #plt.rcParams['font.family'] = 'MS Gothic'

        # フォントファイルのパス
        FONT_PATH = "fonts/NotoSansJP-Regular.otf"
        #FONT_PATH = "fonts/NotoSansJP-VF.otf"

        # フォントをfontManagerに追加
        fm.fontManager.addfont(FONT_PATH)

        # FontPropertiesオブジェクト生成（名前の取得のため）
        font_prop = fm.FontProperties(fname=FONT_PATH)

        # フォントを設定
        rcParams['font.family'] = font_prop.get_name()

        # グラフの描画
        fig = plt.figure()

        # GridSpecで高さ比率を指定し、縦に3つのグラフを表示
        gs = GridSpec(3, 1, height_ratios=[2, 2, 1], hspace=0.3)  # 高さの比率を設定

        # 各グラフを作成
        axes1 = fig.add_subplot(gs[0])
        axes2 = fig.add_subplot(gs[1])
        axes3 = fig.add_subplot(gs[2])

        # 上のグラフ
        axes1.set_title('家庭内健康管理データの可視化(' + name + ':' +
                            year + '年' + month + '月)')
    
        # 血圧
        axes1.plot(x1, y1, marker='v', linestyle='-', label='最低血圧(起床時) ',
                 color='green', markersize=4)
        axes1.plot(x1, y2, marker='^', linestyle='-', label='最高血圧(起床時) ',
                 color='green', markersize=4)
        axes1.grid()  # グラフにグリッド線（格子線）を表示

        # 血圧（60, 100) の罫線の色と線種を変更
        min_limit = 60
        max_limit = 100
        axes1.axhline(y=min_limit, color='red', linestyle='--',
                       label=f'Y={min_limit}')
        axes1.axhline(y=max_limit, color='red', linestyle='--',
                       label=f'Y={max_limit}')
        
        axes1.set_ylabel('血圧(mmHg)', color='green') # Y軸のラベルと色を設定
        axes1.tick_params(axis='y', colors='green') # Y軸の色を設定
        axes1.minorticks_on() # 主要目盛り（メジャー目盛り）の間に補助的な目盛りを表示

        # 凡例の表示
        axes1.legend(bbox_to_anchor=(1.55, 1.0), loc='upper right', 
                        borderaxespad=0)

        # 酸素濃度
        ax1 = axes1.twinx() # X軸は共有し、Y軸は独立して設定(複数のY軸)
        ax1.plot(x1, y4, marker='s', linestyle="dashed", label='酸素濃度(%Sp02)',
                 color='blue', markersize=3)

        # 脈拍数
        ax1.plot(x1, y5, marker= 'x', linestyle="dashed", label='脈拍数(PRbpm)',
                color='orange', markersize=3)
        ax1.set_ylabel('酸素濃度/脈拍数')
        ax1.set_ylim(60,100)
        ax1.minorticks_on()

         # 凡例の表示
        ax1.legend(bbox_to_anchor=(1.55, 0.34), loc='upper right',
                    borderaxespad=0)

        # 体温用のy軸追加
        ax2 = axes1.twinx() # X軸は共有し、Y軸は独立して設定(複数のY軸)
        ax2.spines['right'].set_position(('outward', 45))  # 右側にオフセット
        ax2.plot(x1, y3, marker='o', linestyle='-', label='体温（おでこ）    ',
                   color='red', markersize=3)
        ax2.set_ylabel('体温 ℃', color='red')
        ax2.set_ylim(35,38)
        ax2.tick_params(axis='y', colors='red')
        ax2.minorticks_on()

        # 凡例の表示
        ax2.legend(bbox_to_anchor=(1.55, 0.58), loc='upper right',
                    borderaxespad=0)
        
        # 2つ目のグラフ
        # 歩数用のy軸と棒グラフを追加
        axes2.bar(x1, yt7, label='歩数', color='lavender')
        axes2.set_ylabel('歩数', color='slateblue')
        axes2.tick_params(axis='y', colors='slateblue')
        #axes2.set_xlim(1,31)
        axes2.minorticks_on()

        # 基礎代謝
        ax3 = axes2.twinx()
        ax3.plot(x1, yt6, marker='o', linestyle='-', label='基礎代謝(kcal)     ',
                   color='red', markersize=3)
        if name == "孝則":
            min_value = 1200
            max_value = 1400
        else:
            min_value = 1000
            max_value = 1200
        ax3.set_ylim(min_value, max_value)
        ax3.set_ylabel('基礎代謝(kcal)', color='red')
        ax3.tick_params(axis='y', colors='red')
        ax3.spines["right"].set_position(("axes", 1.1))
        ax3.minorticks_on()

        # 凡例の表示
        ax3.legend(bbox_to_anchor=(1.55, 1.0), loc='upper right',
                    borderaxespad=0)

        # 体重、体脂肪率、内臓脂肪(レべル)
        ax4 = axes2.twinx()
        ax4.plot(x1, yt1, marker='v', linestyle='-', label='体重(kg)',
                 color='green', markersize=3)
        ax4.plot(x1, yt2, marker='^', linestyle='-', label='体脂肪率(％)',
                 color='green', markersize=3)
        ax4.plot(x1, yt5, marker= 'o', linestyle="dashed", label='内臓脂肪(レべル)',
                color='blue', markersize=3)    
        ax4.grid()
        ax4.set_ylabel('値')
        ax4.set_ylim(0, 80)        
        ax4.minorticks_on()

        # 凡例の表示
        ax4.legend(bbox_to_anchor=(1.55, 0.75), loc='upper right',
                    borderaxespad=0)
        
        # 3つ目のグラフ
        # 推定骨量(kg)
        axes3.plot(x1, yt4, marker= 'o', linestyle="dashed", label='推定骨量(kg)',
                color='orange', markersize=3)

        axes3.grid()
        axes3.set_ylabel('推定骨量', color='orange')
        axes3.set_xlabel('日')
        axes3.tick_params(axis='y', colors='orange')

        if name == "孝則":
            axes3.set_ylim(2.0, 3.0)
        else:
            axes3.set_ylim(1.5, 2.5)            
        axes3.minorticks_on()

        # 凡例の表示
        axes3.legend(bbox_to_anchor=(1.37, 1.0), loc='upper right',
                    borderaxespad=0)
        
        # 筋肉量
        ax5 = axes3.twinx()
        ax5.plot(x1, yt3, marker='s', linestyle='-', label='筋肉量(kg)    ',
                 color='blue', markersize=3)
        ax5.grid()
        ax5.set_ylabel('筋肉量', color='blue')
        ax5.tick_params(axis='y', colors='blue')
        if name == "孝則":
            ax5.set_ylim(20, 50)
        else:
            ax5.set_ylim(10, 40) 
        ax5.minorticks_on()

        # 凡例の表示
        ax5.legend(bbox_to_anchor=(1.37, 0.52), loc='upper right',
                    borderaxespad=0)

        # Streamlitで表示
        st.pyplot(fig)                      

    except Exception as e:
        st.sidebar.error("処理対象とする家庭内健康管理データ(" + 
                        excel_file +")の可視化に失敗しました。(" + e.args[0] + ")")