import streamlit as st
import pandas as pd
import warnings
import plotly.express as px

warnings.filterwarnings("ignore")

st.set_page_config(page_title="KPI Dashboard", page_icon="📊", layout="wide")

# Adjust the title size and align it to the left
st.markdown("<h1 style='text-align: left; font-size: 36px;'>📊 KPI Dashboard</h1>", unsafe_allow_html=True)

# File uploader
fl = st.file_uploader("📂 Upload a file", type=["csv", "txt", "xlsx", "xls"])

if fl is not None:
    df = pd.read_csv(fl)
else:
    # Use the RAW GitHub URL instead
    file_path = "https://github.com/NADARJX/KPI/blob/main/KPI%20new-%20May%202025.xlsx"

    try:
        df = pd.read_csv(file_path, encoding="utf-8")
    except pd.errors.ParserError:
        st.error("❌ CSV parsing error! Please check file format or encoding.")
    except Exception as e:
        st.error(f"Unexpected error: {e}")

# Convert column to datetime format
df["Last Submitted DCR Date"] = pd.to_datetime(df["Last Submitted DCR Date"], errors='coerce', dayfirst=True)
df = df.dropna(subset=["Last Submitted DCR Date"])

# Sidebar filters dynamically
st.sidebar.header("Choose your filters")
df_filtered = df.copy()

filters = ["Territory Headquarter", "Territory", "Zone", "ABM", "ZBM", "NSM"]
selected_filters = {}

# **User selects a division for data filtering**
selected_division = st.sidebar.selectbox("Select Division", df["Division Name"].unique())

# **Filter data based on the selected division**
df_filtered = df[df["Division Name"] == selected_division]

for filter_col in filters:
    unique_values = df_filtered[filter_col].dropna().unique()
    selected_values = st.sidebar.multiselect(f"Pick your {filter_col}", unique_values)
    if selected_values:
        df_filtered = df_filtered[df_filtered[filter_col].isin(selected_values)]
        selected_filters[filter_col] = selected_values
        
########

# Ensure numeric columns are properly converted, replacing errors with NaN
numeric_columns = [
    "Call Days", "Doctor Call Avg", "Plan DR Calls", "Actual DR Calls",
    "2PC Freq Cov %", "Total DR Cov %", "Leaves"
]
df_filtered[numeric_columns] = df_filtered[numeric_columns].apply(pd.to_numeric, errors="coerce")

# Replace NaN values with 0 to ensure correct summation
df_filtered.fillna(0, inplace=True)

# Convert integer-based values to correct data types (keeping floats where necessary)
df_filtered = df_filtered.astype({
    "Call Days": "float64",
    "Doctor Call Avg": "float64",
    "Plan DR Calls": "int64",
    "Actual DR Calls": "int64",
    "2PC Freq Cov %": "float64",
    "Total DR Cov %": "float64",
    "Leaves": "int64"
})

# Aggregated data for charts
category_df = df_filtered.groupby("Division Name", as_index=False)["Call Days"].sum().round(2)
doctor_avg_df = df_filtered.groupby("Division Name", as_index=False)["Doctor Call Avg"].mean().round(2)

plan_actual_df = df_filtered.groupby("Division Name", as_index=False)[["Plan DR Calls", "Actual DR Calls"]].sum().round(2)

pc_freq_df = df_filtered.groupby("Division Name", as_index=False)["2PC Freq Cov %"].mean().round(2)
total_dr_cov_df = df_filtered.groupby("Division Name", as_index=False)["Total DR Cov %"].mean().round(2)

leaves_df = df_filtered.groupby("Division Name", as_index=False)["Leaves"].sum()




# **Bar Chart for Call Days with Average**
st.subheader("Division-wise Call Days/Avg and Avg Call")

# Calculate the average Call Days
category_avg_df = df_filtered.groupby("Division Name", as_index=False)["Call Days"].mean()
category_avg_df["Call Days"] = category_avg_df["Call Days"].round(2)  # Round for better readability

# Merge total and average data
category_combined_df = category_df.merge(category_avg_df, on="Division Name", suffixes=("_Total", "_Avg"))

# Ensure Call Days are integers without formatting
category_combined_df["Call Days_Total"] = category_combined_df["Call Days_Total"].astype(int)
category_combined_df["Call Days_Avg"] = category_combined_df["Call Days_Avg"].round(2)  

# Create a Streamlit column layout
col1, col2 = st.columns(2)

# Create the grouped bar chart
fig = px.bar(category_combined_df, x="Division Name", y=["Call Days_Total", "Call Days_Avg"], 
             title="**Total vs Average Call Days by Division**",
             barmode="group", color_discrete_map={"Call Days_Total": "violet", "Call Days_Avg": "orange"},
             text_auto=True, height=500, width=800)  # Adjusted width for better visibility

# Update traces for better visibility
fig.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"), width=0.3  # Adjust bar width
)
# Update layout to improve legend placement
fig.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5))

# Display the chart in col1
with col1:
    st.plotly_chart(fig, use_container_width=True)



# **Bar Chart for Doctor Call Avg**

fig = px.bar(doctor_avg_df, x="Division Name", y="Doctor Call Avg", title="Doctor Call Avg by Division",
             color_discrete_sequence=["green"], text=doctor_avg_df["Doctor Call Avg"].map("{:.2f}".format),height= 470,width=100)
fig.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.3)  # Bigger font for readability

fig.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)
with col2:
     st.plotly_chart(fig, use_container_width=True)
     
# Create a Streamlit column layout
col3, col4 = st.columns(2)     



# Convert values to integers (round off)
plan_actual_df["Plan DR Calls"] = plan_actual_df["Plan DR Calls"].astype(int)
plan_actual_df["Actual DR Calls"] = plan_actual_df["Actual DR Calls"].astype(int)

fig = px.bar(plan_actual_df, x="Division Name", y=["Plan DR Calls", "Actual DR Calls"], 
             title="Plan vs Actual DR Calls", barmode="stack",
             color_discrete_map={"Plan DR Calls": "orange", "Actual DR Calls": "green"}, height=550,
             width= 20,text_auto=True)  # Automatically adds data labels

# Update traces to position text inside the bars
fig.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.3  # Bigger font for readability
)
fig.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)

with col3:
   st.plotly_chart(fig, use_container_width=True)


# Fill NaN values with 0 and convert to integers
df_filtered["2PC DR Total"] = df_filtered["2PC DR Total"].fillna(0).astype(int)
df_filtered["2PC Freq Met"] = df_filtered["2PC Freq Met"].fillna(0).astype(int)



# **2PC Freq Cov %**

fig = px.bar(pc_freq_df, x="Division Name", y="2PC Freq Cov %", title="2PC Freq Cov % by Division",
             color="2PC Freq Cov %", text="2PC Freq Cov %", height= 520 )
fig.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.5  # Bigger font for readability
)
fig.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)

with col4:
     st.plotly_chart(fig, use_container_width=True)
     
col5, col6 = st.columns(2)    

# **Total DR Coverage %**

fig = px.bar(total_dr_cov_df, x="Division Name", y="Total DR Cov %",
             title="Total DR Coverage % by Division", color="Total DR Cov %", text="Total DR Cov %", height=500)
fig.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=16, color="black",weight="bold"),width= 0.3  # Bigger font for readability
)


fig.update_layout(legend=dict(orientation="v",yanchor="top",y=0.2,xanchor="center",x=0.2)
)

fig.update_layout(margin=dict(l=20, r=20, t=40, b=40))

with col5:
    st.plotly_chart(fig, use_container_width=True)


# Fill NaN values with 0 and round percentages
df_filtered["Call Days"] = df_filtered["Call Days"].fillna(0).astype(int)
df_filtered["Plan DR Calls"] = df_filtered["Plan DR Calls"].fillna(0).astype(int)
df_filtered["Actual DR Calls"] = df_filtered["Actual DR Calls"].fillna(0).astype(int)
df_filtered["Doctor Call Avg"] = df_filtered["Doctor Call Avg"].fillna(0).round(2)
df_filtered["2PC Freq Cov %"] = df_filtered["2PC Freq Cov %"].fillna(0).round(2)
df_filtered["Total DR Cov %"] = df_filtered["Total DR Cov %"].fillna(0).round(2)
df_filtered["Non Field Work"] = df_filtered["Non Field Work"].fillna(0).round(2)

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
    "Total DR Cov %": "mean",
    "Non Field Work": "sum"
})

# Round mean values for better readability
metrics_df["Doctor Call Avg"] = metrics_df["Doctor Call Avg"].round(2)
metrics_df["2PC Freq Cov %"] = metrics_df["2PC Freq Cov %"].round(2)
metrics_df["Total DR Cov %"] = metrics_df["Total DR Cov %"].round(2)
metrics_df["Non Field Work"] = metrics_df["Non Field Work"].round(2)


# Create a bar chart for all KPI



# Create a grouped bar chart with all KPIs including Non Field Work
fig = px.bar(metrics_df, x="Division Name", 
             y=["Leaves", "Field Work", "Non Field Work","Total Days", 
                ],
             title="**Comparison of Working Days Divisions**",
             barmode="group", height= 550)

# Show data labels without commas
fig.update_traces(
    texttemplate="<b>%{y:.0f}</b>",  # Bold labels without commas
    textfont=dict(size=18, color="black",weight="bold"),width= 0.2  # Bigger font for readability
)
fig.update_layout(legend=dict(orientation="h",yanchor="top",y=-0.2,xanchor="center",x=0.5)
)
# Display chart
with col6:
    st.plotly_chart(fig, use_container_width=True)



# Fill NaN values with 0 and round percentages
df_filtered["Call Days"] = df_filtered["Call Days"].fillna(0).astype(int)
df_filtered["Plan DR Calls"] = df_filtered["Plan DR Calls"].fillna(0).astype(int)
df_filtered["Actual DR Calls"] = df_filtered["Actual DR Calls"].fillna(0).astype(int)
df_filtered["Doctor Call Avg"] = df_filtered["Doctor Call Avg"].fillna(0).round(2)
df_filtered["2PC Freq Cov %"] = df_filtered["2PC Freq Cov %"].fillna(0).round(2)
df_filtered["Total DR Cov %"] = df_filtered["Total DR Cov %"].fillna(0).round(2)
df_filtered["Total Days"] = df_filtered["Total Days"].fillna(0).round(2)

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
fig = px.line(summary_table, x="Zone", y=["Plan DR Calls", "Actual DR Calls"],
              title="Call and Doctor Visit Trends by Zone",
              markers=True)

# Update layout for better visibility
fig.update_layout(
    title=dict(
        text="Calls Trends by Zone",
        font=dict(size=18, color="black", family="Arial", weight="bold")  # Bold title
    ),
    xaxis_title="Zone",
    yaxis_title="Count",
    legend_title="Metrics"
)

# Display in Streamlit
st.plotly_chart(fig, use_container_width=True)


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
# Prepare data for visualization
df_pie = df_filtered[["Division Name", "Total DR Total", "Total DR Visited", "Total DR MIssed"]].melt(id_vars=["Division Name"], 
                                var_name="Category", value_name="Value")

# Create a donut pie chart
fig = px.pie(df_pie, names="Category", values="Value", 
             title="Doctor Visit Distribution",
             hole=0.4, color="Category",
             color_discrete_map={"Total DR Total": "blue", "Total DR Visited": "green", "Total DR MIssed": "red"})

fig.update_traces(textinfo="label+value+percent", textfont=dict(size=14, color="yellow"))

# Update layout for better readability
fig.update_layout(
    height= 600,  # Reduce height
    width= 600,   # Reduce width
    title=dict(
        text="Doctor Visit Distribution",
        font=dict(size=18, color="black", family="Arial", weight="bold")  # Slightly smaller bold title
    ),
    legend_title="Visit Status"
)

# Display in Streamlit
st.plotly_chart(fig, use_container_width=True)

# **Download Option**
st.subheader("Download Processed Data")
csv = df_filtered.to_csv(index=False).encode('utf-8')
st.download_button(label="📂 Download CSV", data=csv, file_name="processed_data.csv", mime="text/csv")
