import streamlit as st
import pandas as pd
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt
import plotly.express as px
import datetime


st.set_page_config(
    page_title="DATA analytics App",
    page_icon="🎀",
    layout="wide"
)
#link อิโมจิ https://emojipedia.org/window

def line_chart(df,col,limit,sort_by):
    fig, ax = plt.subplots(figsize=(17,10))
    plt.rcParams.update({'font.size': 30})

    for key, grp in df.groupby([sort_by]):
        ax = grp.plot(ax=ax, kind='line', x='StopTime', y=col,label=key[0])

    plt.axhline(int(limit), linestyle="--", color="red", label="limit line")

    plt.xlabel("date")
    plt.ylabel("value")
    plt.title(col)

    plt.legend(loc='upper right',fontsize=20)

    st.pyplot(fig)

st.title("DEMO DATA ANALYTIC APP")
today = datetime.date.today()
yesterday = today - datetime.timedelta(days=1)
start_date = st.sidebar.date_input("Start date",yesterday)
end_date = st.sidebar.date_input("End date",today)

sort_by = st.sidebar.radio('Sort By',options=['Mold','Test Time'])

#stop run กด ctrl+C
#Run D:\Python_Project\src>streamlit run demo-streamlit.py
# with st.sidebar.form("my_form"): ใช้กรณีที่ต้องการให่อยู่ใน Sidebar
# st.sidebar.title("DATA ANALYTICS APP")
df = pd.DataFrame(np.random.randn(50, 20), columns=("col %d" % i for i in range(20)))
tab1,tab2 = st.tabs(["REPORT", "UPLOAD FILE"])
with tab1:
    st.header('REPORT')
    st.write("REPORT")
    st.area_chart(df)
    df = pd.read_excel("./db/file.xlsx")
    df['stopDate'] = pd.to_datetime(df['StopTime'])
    # print(df['stopDate'])
    mask = (df['stopDate'] >= str(start_date)) & (df['stopDate'] <= str(end_date))
    df_process = df.loc[mask]
    st.dataframe(df_process,width=1500)
    st.dataframe(df)

    df_plot = df_process[["stopDate","Band V bright",sort_by]]
    # fig, ax = plt.subplots()
    # ax.plot(df_plot["stopDate"], df_plot["Band V bright"])
    # st.pyplot(fig)

    line_chart(df_process,"Band V bright",7,sort_by)


    with st.expander("See data explanation"):
        st.dataframe(df_plot,width=700)
        st.download_button(
        "EXPORT CSV",
        df_plot.to_csv(index=False).encode('utf-8'),
        "file.csv",
        "text/csv")
with tab2:
    st.header('UPLOAD FILE')
    with st.form("upload_file"):
                        uploaded_file = st.file_uploader("Choose a file")
                        if uploaded_file is not None:
                            df_upload_file = pd.read_excel(uploaded_file)
                            df_upload_file.to_excel("./db/file.xlsx")

                        submitted = st.form_submit_button("Submit")
                        if submitted:
                            if uploaded_file is not None:
                                st.dataframe(df_upload_file,width=1000)