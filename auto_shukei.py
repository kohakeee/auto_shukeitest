import streamlit as st
import pandas as pd
import openpyxl
import seaborn as sns
import io
import base64
import matplotlib.ticker as ticker
import plotly.graph_objs as go

# Streamlitのページ設定
st.set_page_config(
page_title="売上総利益の差額検証",
layout="wide",
)



def main():
    st.title("売上総利益差額検証アプリ")

    # Excelファイル1をアップロードする
    file1 = st.file_uploader("仕入照会ファイルをアップロードしてください", type=["xlsx"])

    # Excelファイル2をアップロードする
    file2 = st.file_uploader("売上照会ファイルをアップロードしてください", type=["xlsx"])

   # Excelファイル1を読み込み、データフレームに変換する
    if file1 is not None:
        df1 = pd.read_excel(file1)

        # "伝票日付"列を日付形式に変換
        df1["伝票日付"] = pd.to_datetime(df1["伝票日付"], format="%Y%m%d")

        min_date = df1["伝票日付"].min().date()
        max_date = df1["伝票日付"].max().date()

        # min_dateとmax_dateの値を表示
        st.write(f"仕入伝票日付最小: {min_date}, 仕入伝票日付最大: {max_date}")
        

        # 担当を営業担当コードに置換する
        df1 = df1.rename(columns={"担当": "営業担当コード"})
        df1 = df1.rename(columns={"勘定科目": "仕入勘定科目"})

        # 勘定科目ごとに集計する
        grouped1 = df1[df1["仕入勘定科目"] == 6].groupby(["管理部門", "管理部門名", "営業担当コード", "担当名", "仕入勘定科目"]).agg({"仕入本体金額": "sum"})
        grouped1_1 = df1[df1["仕入勘定科目"] == 1].groupby(["管理部門", "管理部門名", "営業担当コード", "担当名", "仕入勘定科目"]).agg({"仕入本体金額": "sum"})

    # 営業担当コード、営業担当名を結合してキーとし、売上本体金額、売上原価を勘定科目ごとに集計する
    if file2 is not None:
        df2 = pd.read_excel(file2)

        # "伝票日付"列を日付形式に変換
        df2["伝票日付"] = pd.to_datetime(df2["伝票日付"], format="%Y%m%d")

        min_date = df2["伝票日付"].min().date()
        max_date = df2["伝票日付"].max().date()

        # min_dateとmax_dateの値を表示
        st.write(f"売上伝票日付最小: {min_date}, 売上伝票日付最大: {max_date}")

        # 仕入本体金額を売上原価に置換する
        df2 = df2.rename(columns={"仕入本体金額": "売上原価"})
        df2 = df2.rename(columns={"勘定科目": "売上勘定科目"})


        # 勘定科目ごとに集計する
        grouped2 = df2[df2["売上勘定科目"] == 46].groupby(["管理部門", "管理部門名", "営業担当コード", "営業担当名", "売上勘定科目"]).agg({"売上本体金額": "sum", "売上原価": "sum"})
        grouped2_1 = df2[df2["売上勘定科目"] == 41].groupby(["管理部門", "管理部門名", "営業担当コード", "営業担当名", "売上勘定科目"],dropna=False).agg({"売上本体金額": "sum", "売上原価": "sum"})

        # 管理部門名の下に営業担当名がツリーとして表示されるようにする
        grouped2 = grouped2.groupby(["管理部門", "管理部門名", "売上勘定科目"]).apply(lambda x: x.groupby(["営業担当コード", "営業担当名"]).agg({"売上本体金額": "sum", "売上原価": "sum"}))
        grouped2 = grouped2.reset_index(level=[0, 1, 2, 3, 4])
        grouped2 = grouped2.reset_index(drop=True)

        # 売上本体金額から売上原価を引いた値を計算して粗利を追加する
        grouped2["粗利"] = grouped2["売上本体金額"] - grouped2["売上原価"]
        grouped2_1["部材粗利"] = grouped2_1["売上本体金額"] - grouped2_1["売上原価"]

        # 営業担当コード、営業担当名、管理部門、管理部門名をキーに売上本体金額、仕入本体金額を同じ表で集計
        merged = pd.merge(grouped1.reset_index(), grouped2.reset_index(), how="outer", on=["管理部門", "管理部門名", "営業担当コード",])
        merged1 = pd.merge(grouped1.reset_index(), grouped2.reset_index(), how="outer", on=["管理部門", "管理部門名", "営業担当コード",])



        # 営業担当別集計結果に管理部門、管理部門名も追加する
        grouped3 = merged.groupby(["営業担当コード", "営業担当名",]).agg({"売上本体金額": "sum", "売上原価": "sum", "仕入本体金額": "sum", "粗利": "sum"})

  


        # 集計結果を表示する
        st.write("集計結果:")

        with st.expander("売上集計結果(製品売上46)"):


            st.write(grouped2)
            grouped2_total = grouped2.agg({"売上本体金額": "sum"})
            st.write(grouped2_total)

            # Plotlyで棒グラフを作成
            fig_grouped2 = go.Figure(go.Bar(x=grouped2.index, y=grouped2['売上本体金額']))
            st.plotly_chart(fig_grouped2)
            
        with st.expander("仕入集計結果(商品仕入6)"):
            st.write(grouped1)
            grouped1_total = grouped1.agg({"仕入本体金額": "sum"})
            st.write(grouped1_total)
            
        with st.expander("売上集計結果(部材売上)"):
            st.write(grouped2_1)
            grouped2_1_total = grouped2_1.agg({"売上本体金額": "sum"})
            st.write(grouped2_1_total)

        with st.expander("仕入集計結果(部材仕入)"):
            st.write(grouped1_1)
            grouped1_1_total = grouped1_1.agg({"仕入本体金額": "sum"})
            st.write(grouped1_1_total)


        # 全体の合計
        total_6 = merged1[merged1["仕入勘定科目"] == 6].agg({"売上本体金額": "sum", "仕入本体金額": "sum", "売上原価": "sum","粗利": "sum"})
        total_7 = merged1[(merged1["管理部門"] >= 195000) & (merged1["管理部門"] <= 195199)].agg({"売上本体金額": "sum", "仕入本体金額": "sum", "売上原価": "sum", "粗利": "sum"})
        total_8 = merged1[(merged1["管理部門"] >= 195200) & (merged1["管理部門"] <= 195299)].agg({"売上本体金額": "sum", "仕入本体金額": "sum", "売上原価": "sum", "粗利": "sum"})
       
        # 3カラム表示
        
        st.write("全社合計:")
        st.write(total_6)

        
        
        st.write("東京合計:")
        st.write(total_7)
    
        st.write("大阪合計:")
        st.write(total_8)        

        # merged1を管理部門、管理部門名、営業担当コード、営業担当名、売上勘定科目、仕入勘定科目でグループ化して集計する
        grouped = merged1.groupby(["管理部門", "管理部門名"]).agg({"売上本体金額": "sum", "売上原価": "sum", "仕入本体金額": "sum", "粗利": "sum"})

        # 集計結果を表示する
        st.write("部課別:")
        st.write(grouped)

        # 集計結果を表示する
        st.write("営業担当別集計:")
        st.dataframe(grouped3, width=800, height=1000)

      
        # 新しいデータフレームを作成
        result = pd.DataFrame(index=["東京支店", "東京営業1課", "東京営業2課", "販売促進課", "海外営業課", "中部営業1課", "中部営業2課"],
                                columns=["管理部門", "売上本体金額", "仕入本体金額"])
        
        result1= pd.DataFrame(index=["大阪支店", "大阪営業1課", "大阪営業2課", "大分出張所"],
                                columns=["管理部門", "売上本体金額", "仕入本体金額"])
        
        
        
        
                    
        # 東京支店、東京営業1課、東京営業2課、販売促進課、海外営業課、中部営業1課、中部営業2課のデータを集計
        tokyo = merged1[(merged1["管理部門"] >= 195000) & (merged1["管理部門"] <= 195199)]
        tokyo1 = merged1[(merged1["管理部門"] == 195010)]
        tokyo2 = merged1[(merged1["管理部門"] == 195020)]
        promotion = merged1[(merged1["管理部門"] == 195030)]
        overseas = merged1[(merged1["管理部門"] == 195040)]
        chu1 = merged1[(merged1["管理部門"] == 195110)]
        chu2 = merged1[(merged1["管理部門"] == 195120)]

        # 大阪支店、大阪営業1課、大阪営業2課、大分出張所のデータを集計
        osaka = merged1[(merged1["管理部門"] >= 195201) & (merged1["管理部門"] <= 195299)]
        osaka1 = merged1[(merged1["管理部門"] == 195210)]
        osaka2 = merged1[(merged1["管理部門"] == 195220)]
        oita = merged1[(merged1["管理部門"] == 195230)]
        

        # 各データの集計結果を算出
        
        tokyo_total = tokyo.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        tokyo1_total = tokyo1.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        tokyo2_total = tokyo2.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        promotion_total = promotion.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        overseas_total = overseas.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        chu1_total = chu1.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        chu2_total = chu2.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        oita_total = oita.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        osaka_total = osaka.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        osaka1_total = osaka1.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        osaka2_total = osaka2.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        oita_total = oita.agg({"売上本体金額": "sum", "仕入本体金額": "sum","売上原価": "sum"})
        

        # 東京データをまとめる
        result = pd.concat([tokyo_total, tokyo1_total, tokyo2_total, promotion_total, overseas_total, chu1_total, chu2_total], axis=1)
        result.columns = ["東京支店", "東京営業1課", "東京営業2課", "販売促進課", "海外営業課", "中部営業1課", "中部営業2課"]

        st.write("東京まとめ:")
        st.write(result)

        

            # 大阪データをまとめる
        result1 = pd.concat([osaka_total, osaka1_total, osaka2_total, oita_total], axis=1)
        result1.columns = ["大阪支店", "大阪営業1課", "大阪営業2課", "大分出張所"]


        st.write("大阪まとめ:")
        st.write(result1)

   


       



if __name__ == "__main__":
    main()
