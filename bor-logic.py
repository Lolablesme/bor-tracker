from pyscript import display
from bs4 import BeautifulSoup
import os
import pandas as pd
import plotly.express as px
import requests
import js
from requests.packages import target

url = f"https://www.moh.gov.sg/others/resources-and-statistics/healthcare-institution-statistics-beds-occupancy-rate-(bor)/"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36"
}

# =====================
filename_xlsx = "Bed_Occupancy_Rate.xlsx"
filename_csv = "Bed_Occupancy_Rate.csv"


# =====================
def delete_file(filename):
    if os.path.exists(filename):
        os.remove(filename)


# ======================
hospitals = ['AH', 'CGH', 'KTPH', 'NTFGH', 'NUH(A)', 'SGH', 'SKH', 'TTSH', 'WH']
hospital_colours = {
    "AH": "#1f77b4",  # Deep Blue
    "CGH": "#ff7f0e",  # Safety Orange
    "KTPH": "#2ca02c",  # Emerald Green
    "NTFGH": "#d62728",  # Crimson Red
    "NUH(A)": "#9467bd",  # Royal Purple
    "SGH": "#8c564b",  # Cedar Brown
    "SKH": "#e377c2",  # Pink Magenta
    "TTSH": "#17becf",  # Bright Teal
    "WH": "#bcbd22",  # Olive Gold
    "Average": "#595959"  # Dark Gray
}
# ======================

page = requests.get(url, headers=headers).text
doc = BeautifulSoup(page, "html.parser")

links = doc.find_all("a")

xlsx_url = ""
for link in links:
    if link['href'].endswith(".xlsx"):
        xlsx_url = link['href']
        break
    else:
        continue

if xlsx_url == "":
    exit("No xlsx sheet found")

response = requests.get(xlsx_url, headers=headers)
if response.status_code == 200:
    print(f"Downloaded from {xlsx_url}\n")
    with open(filename_xlsx, "wb") as f:
        f.write(response.content)
else:
    exit(f"Failed to download. Response code: {response.status_code}")

data_xlsx_sheet1 = pd.read_excel(filename_xlsx, skiprows=2, sheet_name=0)
data_xlsx_sheet1 = data_xlsx_sheet1.drop(data_xlsx_sheet1.columns[0], axis='columns')

data_xlsx_sheet2 = pd.read_excel(filename_xlsx, skiprows=2, sheet_name=1)
data_xlsx_sheet2 = data_xlsx_sheet2.iloc[:-2]  # Drop rows with unneeded text

combined_xlsx = pd.concat([data_xlsx_sheet1, data_xlsx_sheet2], ignore_index=True)
combined_xlsx.to_csv(filename_csv, index=False)

data_csv = pd.read_csv(filename_csv)

# Delete created files if they exist
delete_file(filename_xlsx)
delete_file(filename_csv)

# Plot the graphs

for col in hospitals:
    if data_csv[col].dtype == "float64":
        data_csv[col] = data_csv[col] * 100

data_csv['Average'] = data_csv[hospitals].mean(axis='columns', skipna=True)
data_csv['Date'] = pd.to_datetime(data_csv['Date'], errors='coerce')

print(data_csv.head())
print(data_csv.tail())

plots = []

# All hospitals
plt_main = px.line(data_csv, x="Date", y=hospitals + ['Average'],
                   title="Hospital Bed Occupancy Trends (All Hospitals)",
                   labels={"value": "Occupancy Rate(%)", "variable": "Hospital"},
                   color_discrete_map=hospital_colours,
                   render_mode='svg'
                   )
plt_main.update_layout(yaxis=dict(range=[0, 100], ticksuffix="%"),
                       hovermode="x unified")
plt_main.update_traces(hovertemplate="<b>%{fullData.name}</b><br>" +
                                     "Date: %{x|%Y-%m-%d}<br>" +
                                     "Occupancy: %{y:.1f}%<extra></extra>", opacity=0.5)
plots.append(plt_main)
# plt_main.show()

for hospital in hospitals:
    plt_hospital = px.line(data_csv, x="Date", y=[hospital] + ['Average'],
                           title=f"Hospital Bed Occupancy Trends ({hospital})",
                           labels={"value": "Occupancy Rate(%)", "variable": "Hospital"},
                           color_discrete_map=hospital_colours,
                           render_mode='svg'
                           )
    plt_hospital.update_layout(yaxis=dict(range=[0, 100], ticksuffix="%"),
                               hovermode="x unified")
    plt_hospital.update_traces(hovertemplate="<b>%{fullData.name}</b><br>" +
                                             "Date: %{x|%Y-%m-%d}<br>" +
                                             "Occupancy: %{y:.1f}%<extra></extra>", opacity=0.5)
    plots.append(plt_hospital)
    # plt_hospital.show()

try:
    data_csv['Date'] = pd.to_datetime(data_csv['Date'])
    display(f"Last update on: {data_csv['Date'].iloc[-1].strftime('%Y-%m-%d')}", target="script-area")

    for plot in plots:
        display(plot, target="script-area")

finally:
    loading_dialog = js.document.getElementById('loading')
    if loading_dialog:
        loading_dialog.close()
