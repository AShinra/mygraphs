import streamlit as st
import pandas as pd
import altair as alt
import io
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import vl_convert as vlc

# Sample data
def get_data():
    return pd.DataFrame({
        'Category': ['A', 'B', 'C', 'D', 'E'],
        'Value': [10, 25, 15, 30, 20]
    })

# Save Altair chart as PNG in-memory
def get_altair_chart_as_image(chart):
    png_bytes = vlc.vega_to_png(chart.to_dict())
    img_bytes = io.BytesIO(png_bytes)
    img_bytes.seek(0)
    return img_bytes

# Save data and Altair chart to Excel
def save_to_excel(df, chart, filename='bar_chart_data.xlsx'):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    
    # Write data headers
    ws.append(list(df.columns))
    
    # Write data rows
    for row in df.itertuples(index=False):
        ws.append(row)
    
    # Get Altair chart as in-memory image
    img_bytes = get_altair_chart_as_image(chart)
    img = Image(img_bytes)
    ws.add_image(img, "E2")
    
    # Save the workbook
    wb.save(filename)
    return filename

# Streamlit App
st.title("Bar Graph with Vega-Altair and Excel Export")

# Load data
data = get_data()

# Allow users to modify values
for i in range(len(data)):
    data.at[i, 'Value'] = st.number_input(f"Value for {data.at[i, 'Category']}", value=data.at[i, 'Value'])

# Create Altair bar chart
chart = alt.Chart(data).mark_bar().encode(
    x='Category',
    y='Value',
    color='Category'
).properties(width=500, height=300).interactive()

# Display chart in Streamlit
st.altair_chart(chart, use_container_width=True)

# Export data to Excel if button is clicked
if st.button("Download Excel File"):
    filename = save_to_excel(data, chart)
    with open(filename, "rb") as f:
        st.download_button("Download File", f, file_name=filename, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
