import streamlit as st
import pandas as pd
import warnings
import plotly.express as px  # Added Plotly for interactive charts
import plotly.graph_objects as go

import openpyxl
from io import BytesIO

import requests

from datetime import date, timedelta

import datetime


warnings.filterwarnings("ignore")
st.set_page_config(page_title="KPI Dashboard", page_icon="📊", layout="wide")



# Simulated user authentication system with username, division & email
user_roles = {
    "APCMAY": {"division": "Osvita", "email": "venkateshbabu.pr@abbott.com"},
    "APCMAY1": {"division": "Endura", "email": "arijit.gupta@abbott.com"},
    "APCMAY2": {"division": "General Medicine", "email": "basheer.ahmed@abbott.com"},
    "APCMAY3": {"division": "Multi Therapy", "email": "nayan.borthakhur@abbott.com"},
    "APCMAY4": {"division": "NovaNXT", "email": "kailash.parihar@abbott.com"}
}

# Create session state for authentication
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.user_division = None

# Login section
if not st.session_state.authenticated:
    st.title("🔐 Login Page")
    username = st.text_input("Enter your username:")
    email = st.text_input("Enter your email:")

    if username in user_roles and user_roles[username]["email"] == email:
        st.session_state.authenticated = True
        st.session_state.username = username
        st.session_state.user_division = user_roles[username]["division"]
        st.success(f"Welcome {username}! You are authenticated under {st.session_state.user_division} division.")
    else:
        st.error("Access denied. Contact admin for access.")
        st.stop()  # Prevents further execution

# Load data
file_path = r"C:\Users\NADARJX\OneDrive - Abbott\Documents\APC KPI\KPI new- Jun 2025.xlsx"
file_path1 = r"C:\Users\NADARJX\OneDrive - Abbott\Documents\APC KPI\Chronic Missing Report APC - Mar to May.xlsx"
file_path2= r"C:\Users\NADARJX\OneDrive - Abbott\Documents\APC KPI\Comex_Apc.xlsx"
df = pd.read_excel(file_path)
df1 = pd.read_excel(file_path1)
df2 = pd.read_excel(file_path2)
@st.cache_data
def load_data():
    df= pd.read_excel("file_path2.xlsx")
    return df
def load_data1():
    df1=pd.read_excel("file_path1")
    return df1
def load_data2():
    df2=pd.read_excel("file_path2")
    return df2

###url = "https://github.com/NADARJX/KPI/blob/main/KPI%20new-%20May%202025.xlsx"
###df = pd.read_excel(url)

###df = pd.read_excel(url, engine='openpyxl')

st.sidebar.title("📁 Dashboard Menu")
page = st.sidebar.radio(
    "Go to",
    ["📊 KPI Dashboard", "📉 Missed Doctors Report"],
    key="sidebar_nav"
)




#################

# Convert Last Submitted DCR Date to datetime format
df["Last Submitted DCR Date"] = pd.to_datetime(df["Last Submitted DCR Date"], errors='coerce', dayfirst=True)
df = df.dropna(subset=["Last Submitted DCR Date"])

# Apply RLS - Filter data based on authenticated user's division
df_filtered = df[df["Division Name"] == st.session_state.user_division]

# Sidebar Filters - Consolidated selection options



st.markdown("""<style>[data-testid="stSidebar"] {background-color: #ADD8E6;  /* Light blue */}</style>""",unsafe_allow_html=True
)
# Sidebar Filters - Consolidated selection options

st.sidebar.header("Choose your filters")
# Abbott Designation filter
selected_designation = st.sidebar.selectbox("Select Abbott Designation", ["All"] + list(df_filtered["Abbott Designation"].dropna().unique()))
if selected_designation != "All":
    df_filtered = df_filtered[df_filtered["Abbott Designation"] == selected_designation]

# ABM filter (Always Available)
abm_options = df_filtered["ABM"].dropna().unique()
selected_abm = st.sidebar.multiselect("Select ABM", abm_options)
if selected_abm:
    df_filtered = df_filtered[df_filtered["ABM"].isin(selected_abm)]

# Zone filter
zone_options = df_filtered["Zone"].dropna().unique()
selected_zone = st.sidebar.multiselect("Select Zone", zone_options)
if selected_zone:
    df_filtered = df_filtered[df_filtered["Zone"].isin(selected_zone)]

# NSM filter
nsm_options = df_filtered["NSM"].dropna().unique()
selected_nsm = st.sidebar.multiselect("Select NSM", nsm_options)
if selected_nsm:
    df_filtered = df_filtered[df_filtered["NSM"].isin(selected_nsm)]

# ZBM filter
zbm_options = df_filtered["ZBM"].dropna().unique()
selected_zbm = st.sidebar.multiselect("Select ZBM", zbm_options)
if selected_zbm:
    df_filtered = df_filtered[df_filtered["ZBM"].isin(selected_zbm)]

# Territory filter (TBM)
territory_options = df_filtered["Territory"].dropna().unique()
selected_territory = st.sidebar.multiselect("Select Territory", territory_options)
if selected_territory:
    df_filtered = df_filtered[df_filtered["Territory"].isin(selected_territory)]

# Find latest submission date based on Territory selection
latest_date = "NA"
if selected_territory and not df_filtered.empty:
    latest_date = df_filtered["Last Submitted DCR Date"].max().strftime("%d-%b-%Y")

# Display **only one** dynamic heading with latest DCR submission date
st.markdown(
    f"<h1 style='text-align: left; font-size: 38px;'>📊 KPI Dashboard - {st.session_state.user_division} (Last Updated: {latest_date})</h1>",
    unsafe_allow_html=True
)

# Ensure the date column is in datetime format
df["Last Submitted DCR Date"] = pd.to_datetime(df["Last Submitted DCR Date"])

# Date input from user
selected_date = st.date_input("Select a date", datetime.date.today())

########

# Get selected Division Name from user input with a "None" option
division_options = ["None"] + list(df2["DIV_NAME"].dropna().unique())  # Add None as first option
selected_division = st.sidebar.selectbox("Select Division Name", division_options, key="division_filter")

# Apply filtering only if a valid selection is made
if selected_division != "None":
    filtered_df2 = df2[df2["DIV_NAME"] == selected_division]
else:
    filtered_df2 = df2  # Keep full dataset if None is selected

# Compute total EHIER_CD count based on filtered data
total_ehier_cd = filtered_df2["EHIER_CD"].count()
 

# Compute number of unique Territories who submitted DCR for selected date
if selected_date:
    num_dcr_users = df_filtered["Territory"].nunique()
else:
    num_dcr_users = 0  # Default to 0 if no date is selected

# Create two columns for side-by-side KPI cards
col1, col2 = st.columns(2)

# KPI Card: Total EHIER_CD based on selected Division Name
with col1:
    st.markdown(
        f"""
        <div style='border: 4px solid #003366; padding: 5px; width: 200px; margin: auto; text-align: center; background-color: #FF5733; border-radius: 6px;'>
            <p style='font-size: 30px; font-weight: bold; color: white;'>{total_ehier_cd}</p>
            <h2 style='color: white; font-weight: bold; font-size: 14px;'>Total Users for {selected_division}</h2>
        </div>
        """,
        unsafe_allow_html=True
    )

# KPI Card: Number of DCR Updated Users (Filtered by Date)
with col2:
    st.markdown(
        f"""
        <div style='border: 4px solid #003366; padding: 5px; width: 200px; margin: auto; text-align: center; background-color: #007BFF; border-radius: 6px;'>
            <p style='font-size: 30px; font-weight: bold; color: white;'>{num_dcr_users}</p>
            <h2 style='color: white; font-weight: bold; font-size: 14px;'>DCR Updated Users </h2>
        </div>
        """,
        unsafe_allow_html=True
    )







###############################################
# Aggregated data for charts
category_df = df_filtered.groupby("Division Name", as_index=False)["Call Days"].sum().round(0)
doctor_avg_df = df_filtered.groupby("Division Name", as_index=False)["Doctor Call Avg"].mean().round(0)

plan_actual_df = df_filtered.groupby("Division Name", as_index=False)[["Plan DR Calls", "Actual DR Calls"]].sum().round(0)

pc_freq_df = df_filtered.groupby("Division Name", as_index=False)["2PC Freq Cov %"].mean().round(0)
total_dr_cov_df = df_filtered.groupby("Division Name", as_index=False)["Total DR Cov %"].mean().round(0)

leaves_df = df_filtered.groupby("Division Name", as_index=False)["Leaves"].sum().round(0)



###########

# Calculate the average Call Days
category_avg_df = df_filtered.groupby("Division Name", as_index=False)["Call Days"].mean().round(0)
category_avg_df["Call Days"] = category_avg_df["Call Days"].round(0)  # Round for better readability

# Merge total and average data
category_combined_df = category_df.merge(category_avg_df, on="Division Name", suffixes=("_Total", "_Avg"))

# Ensure Call Days are integers without formatting
category_combined_df["Call Days_Total"] = category_combined_df["Call Days_Total"].astype(int)
category_combined_df["Call Days_Avg"] = category_combined_df["Call Days_Avg"].round(0)  

# Create a Streamlit column layout
col1, col2 = st.columns(2)

# Create the grouped bar chart
fig1 = px.bar(category_combined_df, x="Division Name", y=["Call Days_Avg"], 
             title="**Average Call Days by Division**",
             barmode="group", color_discrete_map={"Call Days_Avg": "orange"},
             text_auto=True, height=500, width=800)  # Adjusted width for better visibility

# Update traces for better visibility
fig1.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"), width=0.3  # Adjust bar width
)
# Update layout to improve legend placement
fig1.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))

# Display the chart in col1




# **Bar Chart for Doctor Call Avg**

fig2 = px.bar(doctor_avg_df, x="Division Name", y="Doctor Call Avg", title="Doctor Call Avg by Division",
             color_discrete_sequence=["green"], text=doctor_avg_df["Doctor Call Avg"].map("{:.2f}".format),height= 470,width=100)
fig2.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.3)  # Bigger font for readability

fig2.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)

     
# Create a Streamlit column layout
col3, col4 = st.columns(2)     



# Convert values to integers (round off)
plan_actual_df["Plan DR Calls"] = plan_actual_df["Plan DR Calls"].astype(int)
plan_actual_df["Actual DR Calls"] = plan_actual_df["Actual DR Calls"].astype(int)

fig3 = px.bar(plan_actual_df, x="Division Name", y=["Plan DR Calls", "Actual DR Calls"], 
             title="Plan vs Actual DR Calls", barmode="stack",
             color_discrete_map={"Plan DR Calls": "orange", "Actual DR Calls": "green"}, height=550,
             width= 20,text_auto=True)  # Automatically adds data labels

# Update traces to position text inside the bars
fig3.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.3  # Bigger font for readability
)
fig3.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)




# Fill NaN values with 0 and convert to integers
df_filtered["2PC DR Total"] = df_filtered["2PC DR Total"].fillna(0).astype(int)
df_filtered["2PC Freq Met"] = df_filtered["2PC Freq Met"].fillna(0).astype(int)



# **2PC Freq Cov %**

fig4 = px.bar(pc_freq_df, x="Division Name", y="2PC Freq Cov %", title="2PC Freq Cov % by Division",
             color="2PC Freq Cov %", text="2PC Freq Cov %", height= 520 )
fig4.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.5  # Bigger font for readability
)
fig4.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)


     
col5, col6 = st.columns(2)    

# **Total DR Coverage %**

fig5 = px.bar(total_dr_cov_df, x="Division Name", y="Total DR Cov %",
             title="Total DR Coverage % by Division", color="Total DR Cov %", text="Total DR Cov %", height=500)
fig5.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.3  # Bigger font for readability
)


fig5.update_layout(legend=dict(orientation="v",yanchor="top",y=0.2,xanchor="center",x=0.2)
)

fig5.update_layout(margin=dict(l=20, r=20, t=40, b=40))




# Fill NaN values with 0 and round percentages
df_filtered["Call Days"] = df_filtered["Call Days"].fillna(0).astype(int)
df_filtered["Plan DR Calls"] = df_filtered["Plan DR Calls"].fillna(0).astype(int)
df_filtered["Actual DR Calls"] = df_filtered["Actual DR Calls"].fillna(0).astype(int)
df_filtered["Doctor Call Avg"] = df_filtered["Doctor Call Avg"].fillna(0).round(0)
df_filtered["2PC Freq Cov %"] = df_filtered["2PC Freq Cov %"].fillna(0).round(0)
df_filtered["Total DR Cov %"] = df_filtered["Total DR Cov %"].fillna(0).round(0)
##df_filtered["Non Field Work"] = df_filtered["Non Field Work"].fillna(0).round(0)

# Adding Total Days, Field Work, and Leaves columns
df_filtered["Total Days"] = df_filtered["Total Days"].fillna(0).astype(int)
df_filtered["Field Work"] = df_filtered["Field Work"].fillna(0).astype(int)
df_filtered["Leaves"] = df_filtered["Leaves"].fillna(0).astype(int)


# Group data by Division Name with correct aggregation
metrics_df = df_filtered.groupby("Division Name", as_index=False).agg({
    "Call Days": "sum",
    "Plan DR Calls": "sum",
    "Actual DR Calls": "sum",
    "Leaves": "sum",
    "Field Work": "sum",
    "Total Days": "sum",
    "Doctor Call Avg": "mean",
    "2PC Freq Cov %": "mean",
    "Total DR Cov %": "mean"
})

# Round mean values for better readability
metrics_df["Doctor Call Avg"] = metrics_df["Doctor Call Avg"].round(0)
metrics_df["2PC Freq Cov %"] = metrics_df["2PC Freq Cov %"].round(0)
metrics_df["Total DR Cov %"] = metrics_df["Total DR Cov %"].round(0)
##metrics_df["Non Field Work"] = metrics_df["Non Field Work"].round(0)


# Create a bar chart for all KPI



# Create a grouped bar chart with all KPIs including Non Field Work
fig6 = px.bar(metrics_df, x="Division Name", 
             y=["Leaves", "Field Work","Total Days", 
                ],
             title="**Comparison of Working Days Divisions**",
             barmode="group", height= 550)

# Show data labels without commas
fig6.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=18, color="black",weight="bold"),width= 0.2  # Bigger font for readability
)
fig6.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)
# Display chart




# Fill NaN values with 0 and round percentages
df_filtered["Call Days"] = df_filtered["Call Days"].fillna(0).astype(int)
df_filtered["Plan DR Calls"] = df_filtered["Plan DR Calls"].fillna(0).astype(int)
df_filtered["Actual DR Calls"] = df_filtered["Actual DR Calls"].fillna(0).astype(int)
df_filtered["Doctor Call Avg"] = df_filtered["Doctor Call Avg"].fillna(0).round(0)
df_filtered["2PC Freq Cov %"] = df_filtered["2PC Freq Cov %"].fillna(0).round(0)
df_filtered["Total DR Cov %"] = df_filtered["Total DR Cov %"].fillna(0).round(0)
df_filtered["Total Days"] = df_filtered["Total Days"].fillna(0).round(0)

# Group data by Zone
summary_table = df_filtered.groupby("Zone", as_index=False).agg({
    "Call Days": "sum",
    "Plan DR Calls": "sum",
    "Actual DR Calls": "sum",
    "Doctor Call Avg": "mean",
    "2PC Freq Cov %": "mean",
    "Total DR Cov %": "mean"
})

# Create a line chart
fig7 = px.line(summary_table, x="Zone", y=["Plan DR Calls", "Actual DR Calls"],
              title="**Call and Doctor Visit Trends by Zone**",
              markers=True)

# Update layout for better visibility
fig7.update_layout(
    title=dict(
        text="Calls Trends by Zone",
        font=dict(size=18, color="black", family="Arial", weight="bold")  # Bold title
    ),
    xaxis_title="Zone",
    yaxis_title="Count",
    legend_title="Metrics"
)

# Display in Streamlit

col8, col9 = st.columns(2)

st.markdown(
    """
    <style>
    th {
        font-weight: bold !important;
        color: black !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# Ensure numeric conversion
df_filtered["1PC Freq Cov %"] = pd.to_numeric(df_filtered["1PC Freq Cov %"], errors="coerce")
df_filtered["2PC Freq Cov %"] = pd.to_numeric(df_filtered["2PC Freq Cov %"], errors="coerce")

# Group by Division to calculate averages
division_data = df_filtered.groupby("Division Name", as_index=False)[["1PC Freq Cov %", "2PC Freq Cov %"]].mean().round(2)



for index, row in division_data.iterrows():
    division_name = row["Division Name"]
    avg_1pc = row["1PC Freq Cov %"]
    avg_2pc = row["2PC Freq Cov %"]

    # Gauge Chart for 1PC Coverage %
    fig8= go.Figure(go.Indicator(
        mode="gauge+number",
        value=avg_1pc,
        title={"text": f"{division_name} - 1PC Coverage %"},
        gauge={"axis": {"range": [0, 100]}, "bar": {"color": "blue"}}
    ))

    # Gauge Chart for 2PC Coverage %
    fig9 = go.Figure(go.Indicator(
        mode="gauge+number",
        value=avg_2pc,
        title={"text": f"{division_name} - 2PC Coverage %"},
        gauge={"axis": {"range": [0, 100]}, "bar": {"color": "green"}}
    ))

  


############### 
col10, col11 =st.columns(2)
# Get unique division names
division_names = df_filtered["Division Name"].unique()

# Loop through each division and create a stacked bar chart
for division in division_names:
    # Filter data for the current division
    df_division = df_filtered[df_filtered["Division Name"] == division]

    # Group and melt the data
    df_grouped = df_division.groupby("Division Name")[["Total DR Total", "Total DR Visited", "Total DR MIssed"]].sum().reset_index()
    df_melted = df_grouped.melt(id_vars="Division Name", var_name="Category", value_name="Value")

    # Calculate percentage
    df_melted["Percentage"] = df_melted["Value"] / df_melted["Value"].sum() * 100

# Create bar chart without percentage
fig10 = px.bar(df_melted, x="Division Name", y="Value", color="Category",title=f"Doctor Visit Distribution for {division}",text="Value")

fig10.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=12, color="black",weight="bold"), width=0.3  # Adjust bar width
)
# Update layout to improve legend placement
fig10.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))

####### 
# Get unique division names
division_names = df_filtered["Division Name"].unique()

# Loop through each division and create a stacked bar chart
for division in division_names:
    # Filter data for the current division
    df_division = df_filtered[df_filtered["Division Name"] == division]

# Group by Division and Designation and calculate average coverage
    df_sum = df_division.groupby(["Abbott Designation"], as_index=False)["Call Days"].sum()

    
# Create bar chart of average doctor coverage by designation within each division
fig11 = px.bar(df_sum,
             x="Abbott Designation",
             y="Call Days",
             color="Abbott Designation",
          
             title="Total Call Days by Designation within Division",
             labels={"Call Days": "Total Call Days"},
             template="plotly_white",
             text=df_sum["Call Days"].round(0)
)


fig11.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=12, color="black",weight="bold"), width=0.3  # Adjust bar width
)
# Update layout to improve legend placement
fig11.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))

#######    
# Get unique division names
division_names = df_filtered["Division Name"].unique()

# Loop through each division and create a stacked bar chart
for division in division_names:
    # Filter data for the current division
    df_division = df_filtered[df_filtered["Division Name"] == division]

    # Create the stacked bar chart
    fig12 = px.bar(
        df_division,
        x="Full Name",
        y="Total DR Cov %",
        color="Zone",  # Use Abbott Designation column if applicable
        title=f"Total Doctor Coverage % for {division}",
        labels={"Total DR Cov %": "Doctor Coverage %", "Full Name": "Employee Name"},
        template="plotly_white",
        barmode="stack"  # Enables stacked bar mode
    )

    # Add data labels
    fig12.update_traces(texttemplate='%{y:.2f}%', textposition='outside')

 


###PAGEs###########




# Common division filtering
### division = st.session_state.user_division


# Load Excel file
    
df1 = pd.read_excel(file_path1, sheet_name="Base Data", engine="openpyxl")

    # Apply RLS - Filter data based on authenticated user's division
df_filtered1 = df1[df1["Divison Name"] == st.session_state.user_division]

    # Sidebar Filters
 
st.sidebar.header("Choose your filters (Missed Doctors)")

    # Division Name filter
division_options = df_filtered1["Divison Name"].dropna().unique()
selected_division = st.sidebar.multiselect("Select Division", division_options)
if selected_division:
        df_filtered1= df_filtered1[df_filtered1["Divison Name"].isin(selected_division)]

    # TBM Name filter
tbm_options = df_filtered1["TBM Name"].dropna().unique()
selected_tbm = st.sidebar.multiselect("Select TBM Name", tbm_options)
if selected_tbm:
        df_filtered1 = df_filtered1[df_filtered1["TBM Name"].isin(selected_tbm)]

    # ABM Name filter
abm_options = df_filtered1["ABM Name"].dropna().unique()
selected_abm = st.sidebar.multiselect("Select ABM", abm_options)
if selected_abm:
        df_filtered1 = df_filtered1[df_filtered1["ABM Name"].isin(selected_abm)]

    # ZBM Name filter
zbm_options = df_filtered1["ZBM Name"].dropna().unique()
selected_zbm = st.sidebar.multiselect("Select ZBM", zbm_options)
if selected_zbm:
        df_filtered1 = df_filtered1[df_filtered1["ZBM Name"].isin(selected_zbm)]

    # Month filter
month_options = df_filtered1["Month"].dropna().unique()
selected_month = st.sidebar.multiselect("Select Month", month_options)
if selected_month:
        df_filtered1 = df_filtered1[df_filtered1["Month"].isin(selected_month)]

    # Frequency filter
freq_options = df_filtered1["To be Met"].dropna().unique()
selected_freq = st.sidebar.multiselect("Select Frequency", freq_options)
if selected_freq:
        df_filtered1 = df_filtered1[df_filtered1["To be Met"].isin(selected_freq)]

    # --- Chart 1: Unique Customer Count by Specialty ---

##st.markdown("## **Unique Doctors Missed in Division (Last 3 Months)**")
specialty_counts = df_filtered1.groupby('Specialty By Practice')['Customer Code'].nunique().reset_index()
specialty_counts = specialty_counts.sort_values(by='Customer Code', ascending=False)


    
fig13 = px.bar(
        specialty_counts,
        x='Specialty By Practice',
        y='Customer Code',
        text='Customer Code',
        labels={'Customer Code': 'Unique Customer Count'},
        color_discrete_sequence=["#E6ADDE"]
    )
fig13.update_traces(
        texttemplate="<b>%{y:.0f}</b>",
        textfont=dict(size=16, color="black"),
        width=0.8
    )
fig13.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))


    # --- Chart 2: Total Frequency of Doctors ---
###st.markdown("### **Total Frequency of Doctors 1, 2, 3**")
division_names = df_filtered1['Divison Name'].unique()
selected_division_chart = st.selectbox("Select Division Name", division_names)

filtered_data = df_filtered1[df_filtered1['Divison Name'] == selected_division_chart]
frequency_data = (
        filtered_data.groupby('Specialty By Practice')['To be Met']
        .sum()
        .reset_index()
        .sort_values(by='To be Met', ascending=False)
    )

fig14 = px.bar(
        frequency_data,
        x='Specialty By Practice',
        y='To be Met',
        text='To be Met',
        labels={'To be Met': 'Frequency', 'Specialty By Practice': 'Specialty'},
        color_discrete_sequence=["#008004"])
fig14.update_traces(
        texttemplate="<b>%{y:.0f}</b>",
        textfont=dict(size=16, color="black"),
        width=0.8
    )
fig14.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))


    # --- Summary Table ---
##st.subheader("Missing HCP Details Summary")
summary_table = df1[['Customer Code', 'HCP Name', 'Specialty By Practice', 'To be Met']]

if page == "📊 KPI Dashboard":
    st.subheader("********")
   
    with col1:
        st.plotly_chart(fig1, use_container_width=True)
 
    with col2:
        st.plotly_chart(fig2, use_container_width=True)
   
    with col3:
        st.plotly_chart(fig3, use_container_width=True)   

    with col4:
        st.plotly_chart(fig4, use_container_width=True)   
  
    with col5:
        st.plotly_chart(fig5, use_container_width=True)
 
    with col6:
        st.plotly_chart(fig6, use_container_width=True)
   
    st.plotly_chart(fig7, use_container_width=True)
   
    with col8:
        st.plotly_chart(fig8, use_container_width=True)

    with col9:
        st.plotly_chart(fig9, use_container_width=True)
  
    with col10:
        st.plotly_chart(fig10, use_container_width=True) 
    
    with col11:
        st.plotly_chart(fig11, use_container_width=True)                       
         
# --- Page 2: Missed Doctors Report ---
elif page == "📉 Missed Doctors Report":
    st.title("📉 Missed Doctors - Last 3 Months")
    st.plotly_chart(fig13, use_container_width=True)
    st.plotly_chart(fig14, use_container_width=True)
    st.dataframe(summary_table)
    
       
        
# **Download Option**
st.subheader("Download Processed Data")
csv = df_filtered.to_csv(index=False).encode('utf-8')
st.download_button(label="📂 Download CSV", data=csv, file_name="processed_data.csv", mime="text/csv")
    
