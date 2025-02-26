import openpyxl
import pandas as pd
import streamlit as st
from graphs import bar_graph
from streamlit_option_menu import option_menu


def get_data(csv_file):
    return pd.read_csv(csv_file)



if __name__ == '__main__':

    with st.sidebar:
       selected = option_menu(
           menu_title='Graph Type',
           menu_icon='map',
           options=['Bar', 'Pie', 'Line'],
           icons=['bar-chart-line', 'pie-chart', 'graph-up']
       ) 

    st.file_uploader('Upload File', type='csv', key='csvfile')

    if st.session_state['csvfile'] not in [None, '']:
        df = get_data(st.session_state['csvfile'])

        with st.container(border=True):
            st.dataframe(df, use_container_width=True)

            if selected == 'Bar':
                bar_graph(df)

            
            if selected == 'Pie':
                st.warning('Under Development')

            if selected == 'Line':
                st.warning('Under Development')
