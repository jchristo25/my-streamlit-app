import streamlit as st


st.write('TARIKA ANALYSIS')

pages = {
    "Visualizations": [
        st.Page("TARIKA_STREAMLIT.py", title="TARIKA"),
    ],
}

pg = st.navigation(pages)
pg.run()