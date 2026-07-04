import streamlit as st
import pandas as pd

st.title("My Streamlit App")

df = pd.DataFrame({
    "name": ["A", "B", "C"],
    "value": [10, 20, 30]
})

st.bar_chart(df.set_index("name"))