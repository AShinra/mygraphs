import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.plotarea import DataTable
from io import BytesIO
import tempfile
from PIL import Image


def bar_graph(df):

    # Create a new Workbook
    wb = openpyxl.Workbook()
    wb.create_sheet('Bar')
    ws = wb['Bar']
    wb.save('Bar.xlsx')

    btn_create_graph = st.button('Create Graph')

    
    if btn_create_graph:
        
        l, w = df.shape

        # append data to excel sheet
        writer = pd.ExcelWriter('Bar.xlsx', engine='openpyxl')
        df.to_excel(writer, sheet_name='Bar', startrow = 0,index = False)
        writer.close()

        wb = openpyxl.load_workbook('Bar.xlsx')
        ws = wb['Bar']

        # Create a BarChart object
        _chart = BarChart()
        _chart.type = 'col'
        _chart.style = 2
        _chart.grouping = 'stacked'
        _chart.overlap = 100
        _chart.gapWidth = 10
        # _chart.x_axis.delete = False
        # _chart.y_axis.delete = False
        _chart.legend.position = 'b'

        # Define the data (sales values) for the chart
        data = Reference(ws, min_col=2, min_row=1, max_row=l+1, max_col=w)

        # Define the categories (years)
        categories = Reference(ws, min_col=1, min_row=2, max_row=l+1, max_col=1)

        # Add data and categories to the chart
        _chart.add_data(data, titles_from_data=True)
        _chart.set_categories(categories)
        _chart.shape = 4
        _chart.dataLabels = DataLabelList()
        _chart.dataLabels.showVal = True
        _chart.DataTable = DataTable()
        _chart.height = 15
        _chart.width = 30

        _chart.title = "Custom Bar Chart"
        _chart.y_axis.title = "Values"
        _chart.x_axis.title = "Categories"
        _chart.y_axis.majorGridlines = None
        _chart.x_axis.majorGridlines = None
        _chart.y_axis.tickLblPos = "low"
        _chart.x_axis.tickLblPos = "low"


        # Add the chart to the worksheet
        ws.add_chart(_chart, "J1")

        # Save the Workbook to a file
        wb.save("Bar.xlsx")        

    

    return