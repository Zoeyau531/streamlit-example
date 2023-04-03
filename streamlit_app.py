from collections import namedtuple
import altair as alt
import math
import pandas as pd
import streamlit as st

csvfile = st.file_uploader(AnnualLeaveRecord.csv)


st.write(csvfile)


wb.save(r"C:\Users\Zoe\Desktop\Python Learn\Annual Leave\AnnualLeaveRecord.xlsx")

wb.close()


