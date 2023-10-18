import streamlit as st
import pandas as pd
import numpy as np
import xlwings as xw
import matplotlib.pyplot as plt
import plotly.express as px


st.set_page_config(
    page_title="DATA analytics App",
    page_icon="💕",
    layout="wide"
)
#link อิโมจิ https://emojipedia.org/window
st.title("TITLE")
st.header("header")
st.subheader("subheader")
st.write("Hello World")
st.caption("caption")

df = pd.DataFrame(np.random.randn(50, 20), columns=("col %d" % i for i in range(20)))
st.dataframe(df,width=1500)

st.area_chart(df)
st.line_chart(df)
st.scatter_chart(df)

data = pd.DataFrame(data=np.random.rand(4,4)*100000,
index=["Q1","Q2","Q3","Q4"],
columns=["East","West","North","South"])
data.index.name = "Quaters"
data.columns.name = "Region"
fig = data.plot.bar().get_figure()
st.pyplot(fig)
#stop run กด ctrl+C
#Run D:\Python_Project\src>streamlit run demo-streamlit.py


df = px.data.gapminder()

fig = px.scatter(
    df.query("year==2007"),
    x="gdpPercap",
    y="lifeExp",
    size="pop",
    color="continent",
    hover_name="country",
    log_x=True,
    size_max=60,
)
st.plotly_chart(fig, theme="streamlit", use_container_width=True)

if st.button("Submit"): 
    st.write("Click submit")

def convert_df(df):
    return df.to_csv().encode('utf-8')

csv = convert_df(df)

st.download_button(
    label="Download data CSV",
    data=csv,
    file_name='df.csv',
    mime='text/csv',
)
df = pd.DataFrame(np.random.randn(50, 20), columns=("col %d" % i for i in range(20)))
col1,col2,col3,col4 = st.columns(4)

with col1:
    st.header('col1')
    st.area_chart(df)

with col2:
    st.header('col2')
    st.area_chart(df)

with col3:
    st.header('col3')
    st.area_chart(df)

with col4:
    st.header('col4')
    st.area_chart(df)

col1,col2,col3 = st.columns(3)

with col1:
    st.header('col1')
    st.area_chart(df)

with col2:
    st.header('col2')
    st.area_chart(df)

with col3:
    st.header('col3')
    st.area_chart(df)

col1,col2 = st.columns(2)

with col1:
    st.header('col1')
    st.area_chart(df)

with col2:
    st.header('col2')
    st.area_chart(df)

contact_lists = ["Email","Home Phone","Mobile phone"]

selectbox_value = st.selectbox("select Contacted", contact_lists, index=None, placeholder="Select Contact Method")
st.write(selectbox_value)

# with st.sidebar.form("my_form"): ใช้กรณีที่ต้องการให่อยู่ใน Sidebar
# st.sidebar.title("DATA ANALYTICS APP")
with st.form("my_form"):
    multi_value = st.multiselect("select Contacted", contact_lists, placeholder="Select Contact Method")

    submit = st.form_submit_button("SUBMIT")
    if submit:
        st.write(multi_value)

tab1,tab2 = st.tabs(["tab1", "tab2"])
with tab1:
    st.header('tab1')

with tab2:
    st.header('tab2')
    with st.form("upload_file"):
                        uploaded_file = st.file_uploader("Choose a file")
                        if uploaded_file is not None:
                            df_upload_file = pd.read_excel(uploaded_file)


                        submitted = st.form_submit_button("Submit")
                        if submitted:
                            if uploaded_file is not None:
                                st.dataframe(df_upload_file,width=1000)