import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
import io
import base64

# Page configuration with professional blue and beige theme
st.set_page_config(
    page_title="Crane Expenses Dashboard",
    page_icon="🏗️",
    layout="wide"
)

# Custom CSS for enhanced styling
st.markdown("""
<style>
    .main {
        background-color: #f8f5f0;
    }
    .stApp {
        background: linear-gradient(135deg, #e6f0ff 0%, #f8f5f0 100%);
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #e6f0ff;
        border-radius: 8px 8px 0px 0px;
        padding: 10px 20px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0077b6;
        color: white;
        box-shadow: 0 4px 8px rgba(0, 119, 182, 0.2);
    }
    .title-header {
        color: #0077b6;
        text-align: center;
        padding: 20px;
        font-weight: 700;
        text-shadow: 1px 1px 3px rgba(0,0,0,0.1);
    }
    .info-box {
        background-color: #e6f0ff;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid #0077b6;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .metric-card {
        background: linear-gradient(135deg, #0077b6 0%, #03045e 100%);
        color: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    .progress-bar {
        border-radius: 10px;
        overflow: hidden;
        margin-bottom: 10px;
    }
    .alert-box {
        background-color: #fff3cd;
        border: 1px solid #ffeaa7;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Header with improved styling
st.markdown('<h1 class="title-header">🏗️ Crane Expenses Dashboard</h1>', unsafe_allow_html=True)
st.markdown("_Professional Analysis Dashboard v1.0_")

@st.cache_data
def load_data(file):
    data = pd.read_excel(file)
    return data

# Budget data (dynamic based on years in data)
@st.cache_data
def get_budget_data(years):
    """Generate budget data for all years present in the dataset"""
    budget = {}
    base_budget = 150000
    for year in years:
        budget[str(year)] = {
            'JANVIER': base_budget, 'FÉVRIER': base_budget * 0.97, 
            'MARS': base_budget * 1.03, 'AVRIL': base_budget * 1.07,
            'MAI': base_budget * 1.10, 'JUIN': base_budget * 1.13,
            'JUILLET': base_budget * 1.17, 'AOÛT': base_budget * 1.20,
            'SEPTEMBRE': base_budget * 1.23, 'OCTOBRE': base_budget * 1.27,
            'NOVEMBRE': base_budget * 1.30, 'DÉCEMBRE': base_budget * 1.33
        }
        base_budget *= 1.1  # Increase base budget by 10% each year
    return budget

def create_download_links(df, filtered_df):
    """Create download buttons for Excel reports"""
    
    # Excel download
    towrite = io.BytesIO()
    df.to_excel(towrite, index=False, engine='openpyxl')
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="crane_expenses_full.xlsx" style="display: inline-block; padding: 10px 20px; background: #0077b6; color: white; border-radius: 5px; text-decoration: none; margin: 10px 0;">📊 Download Full Excel Report</a>'
    
    # Filtered Excel download
    towrite_filtered = io.BytesIO()
    filtered_df.to_excel(towrite_filtered, index=False, engine='openpyxl')
    towrite_filtered.seek(0)
    b64_filtered = base64.b64encode(towrite_filtered.read()).decode()
    href_filtered = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_filtered}" download="crane_expenses_filtered.xlsx" style="display: inline-block; padding: 10px 20px; background: #00b4d8; color: white; border-radius: 5px; text-decoration: none; margin: 10px 0;">📊 Download Filtered Excel Report</a>'
    
    st.markdown(href, unsafe_allow_html=True)
    st.markdown(href_filtered, unsafe_allow_html=True)

def check_budget_alerts(data, budget_data):
    """Check for budget alerts and unusual spending patterns"""
    alerts = []
    
    # Monthly budget alerts
    monthly_spending = data.groupby(['Annee', 'Mois'])['Deponse'].sum().reset_index()
    for _, row in monthly_spending.iterrows():
        year = str(row['Annee'])
        month = row['Mois']
        spending = row['Deponse']
        
        if year in budget_data and month in budget_data[year]:
            budget = budget_data[year][month]
            if spending > budget * 1.15:  # 15% over budget
                alerts.append(f"🚨 {month} {year}: Spending {spending:,.0f} DH exceeds budget {budget:,.0f} DH by {(spending/budget-1)*100:.1f}%")
            elif spending > budget:
                alerts.append(f"⚠️ {month} {year}: Spending {spending:,.0f} DH is over budget {budget:,.0f} DH by {(spending/budget-1)*100:.1f}%")
    
    # Provider spending alerts
    provider_avg = data.groupby('Prestataire/ Frs')['Deponse'].mean()
    overall_avg = data['Deponse'].mean()
    for provider, avg_spend in provider_avg.items():
        if avg_spend > overall_avg * 2:  # 2x higher than average
            alerts.append(f"🔍 {provider}: Average spending {avg_spend:,.0f} DH is significantly higher than overall average")
    
    return alerts

upload_file = st.file_uploader("📤 Upload GRUE PAR PRESTATAIRE Excel file", type=["xlsx"])
if upload_file is None:
    st.info("Please upload an Excel file to begin analysis.", icon="ℹ️")
    st.stop()

df = load_data(upload_file)

# Get unique years and sort them
all_years = sorted(df['Annee'].unique())
budget_data = get_budget_data(all_years)

# Update title with dynamic years
year_range = f"{min(all_years)}-{max(all_years)}"
st.markdown(f'<h1 class="title-header">🏗️ Crane Expenses Dashboard {year_range}</h1>', unsafe_allow_html=True)

# Advanced Filters Section
st.sidebar.markdown("### 🔧 Advanced Filters")

# Year filter (sorted)
selected_years = st.sidebar.multiselect(
    "Select Year(s)", 
    options=all_years, 
    default=all_years
)

# Month filter
months = ['JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 
          'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE']
selected_months = st.sidebar.multiselect(
    "Select Month(s)",
    options=months,
    default=months
)

# Provider filter
providers = sorted(df['Prestataire/ Frs'].unique())
selected_providers = st.sidebar.multiselect(
    "Select Service Provider(s)",
    options=providers,
    default=providers
)

# Crane filter
cranes = sorted(df['Grue'].unique())
selected_cranes = st.sidebar.multiselect(
    "Select Crane(s)",
    options=cranes,
    default=cranes
)

# Spending threshold filter
min_spend, max_spend = st.sidebar.slider(
    "Expense Range (DH)",
    min_value=int(df['Deponse'].min()),
    max_value=int(df['Deponse'].max()),
    value=(int(df['Deponse'].min()), int(df['Deponse'].max()))
)

# Apply filters
filtered_df = df[
    (df['Annee'].isin(selected_years)) &
    (df['Mois'].isin(selected_months)) &
    (df['Prestataire/ Frs'].isin(selected_providers)) &
    (df['Grue'].isin(selected_cranes)) &
    (df['Deponse'] >= min_spend) &
    (df['Deponse'] <= max_spend)
]

with st.expander("📊 Data Preview", expanded=False):
    st.dataframe(filtered_df, use_container_width=True)

# Download Reports Section
st.sidebar.markdown("---")
st.sidebar.markdown("### 📥 Download Reports")
create_download_links(df, filtered_df)

# Alert System
st.sidebar.markdown("---")
st.sidebar.markdown("### 🚨 Budget Alerts")
alerts = check_budget_alerts(filtered_df, budget_data)
if alerts:
    for alert in alerts[:3]:  # Show first 3 alerts
        st.sidebar.markdown(f'<div class="alert-box">{alert}</div>', unsafe_allow_html=True)
    if len(alerts) > 3:
        st.sidebar.info(f"+ {len(alerts) - 3} more alerts...")
else:
    st.sidebar.success("No budget alerts - spending within normal ranges")

def create_3d_surface_plot(data):
    """Create a 3D surface plot of expenses with proper date formatting"""
    month_mapping = {
        'JANVIER': 1, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'JUIN': 6,
        'JUILLET': 7, 'AOÛT': 8, 'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DÉCEMBRE': 12
    }
    
    plot_data = data.copy()
    plot_data['Month_Num'] = plot_data['Mois'].map(month_mapping)
    
    # Sort data by year and month for correct ordering
    plot_data = plot_data.sort_values(['Annee', 'Month_Num'])
    
    # Create date labels for the y-axis
    plot_data['Date_Label'] = plot_data.apply(lambda row: f"{row['Month_Num']}/{row['Annee']}", axis=1)
    
    # Pivot data for surface plot
    pivot_data = plot_data.pivot_table(
        values='Deponse',
        index=['Annee', 'Month_Num', 'Date_Label'],
        columns='Prestataire/ Frs',
        aggfunc='sum',
        fill_value=0
    ).reset_index().sort_values(['Annee', 'Month_Num'])
    
    providers = pivot_data.columns[3:]
    
    # Create mesh grid
    X, Y = np.meshgrid(range(len(providers)), range(len(pivot_data)))
    Z = np.zeros((len(pivot_data), len(providers)))
    
    for i, provider in enumerate(providers):
        Z[:, i] = pivot_data[provider].values
    
    # Create custom y-axis labels
    y_labels = pivot_data['Date_Label'].tolist()
    
    fig = go.Figure(data=[go.Surface(z=Z, x=X, y=Y, colorscale='Blues', opacity=0.9)])
    
    fig.update_layout(
        title='3D Surface: Expenses Analysis by Provider and Time',
        scene=dict(
            xaxis=dict(
                title='Service Providers', 
                tickvals=list(range(len(providers))), 
                ticktext=[p[:12] + '...' if len(p) > 12 else p for p in providers]
            ),
            yaxis=dict(
                title='Time Period (Month/Year)', 
                tickvals=list(range(len(pivot_data))),
                ticktext=y_labels
            ),
            zaxis=dict(title='Expenses (DH)'),
            camera=dict(eye=dict(x=1.8, y=1.8, z=1.2))
        ),
        width=1000,
        height=600,
        margin=dict(l=65, r=50, b=65, t=90)
    )
    
    return fig

def create_3d_scatter_plot(data):
    """Create a 3D scatter plot with proper date formatting for any years"""
    month_mapping = {
        'JANVIER': 1, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'JUIN': 6,
        'JUILLET': 7, 'AOÛT': 8, 'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DÉCEMBRE': 12
    }
    
    plot_data = data.copy()
    plot_data['Month_Num'] = plot_data['Mois'].map(month_mapping)
    
    # Sort data by year and month for correct ordering
    plot_data = plot_data.sort_values(['Annee', 'Month_Num'])
    
    # Create date labels for the y-axis
    plot_data['Date_Label'] = plot_data.apply(lambda row: f"{row['Month_Num']}/{row['Annee']}", axis=1)
    
    # UNIVERSAL FIX: Calculate time index based on minimum year
    min_year = data['Annee'].min()
    plot_data['Time_Index'] = (plot_data['Annee'] - min_year) * 12 + plot_data['Month_Num']
    
    fig = px.scatter_3d(
        plot_data,
        x='Prestataire/ Frs',
        y='Time_Index',
        z='Deponse',
        color='Grue',
        size='Deponse',
        hover_data=['Mois', 'Annee', 'Date_Label'],
        title='3D Scatter: Expenses Distribution by Provider and Time',
        labels={
            'Prestataire/ Frs': 'Service Provider',
            'Time_Index': 'Time Period',
            'Deponse': 'Expenses (DH)',
            'Grue': 'Crane ID'
        }
    )
    
    # Create custom y-axis labels (UNIVERSAL - works for any years)
    unique_times = sorted(plot_data['Time_Index'].unique())
    y_labels = []
    for time_idx in unique_times:
        year = min_year + (time_idx - 1) // 12
        month = (time_idx - 1) % 12 + 1
        y_labels.append(f"{month}/{year}")
    
    fig.update_layout(
        scene=dict(
            xaxis=dict(title='Service Providers', tickangle=45),
            yaxis=dict(
                title='Time Period (Month/Year)', 
                tickvals=unique_times,
                ticktext=y_labels
            ),
            zaxis=dict(title='Expenses (DH)'),
            camera=dict(eye=dict(x=0, y=-2.5, z=1.5))
        ),
        width=1000,
        height=600,
        margin=dict(l=65, r=50, b=65, t=90)
    )
    
    return fig

def create_advanced_bar_chart(data):
    """Create an advanced 2D bar chart that looks professional"""
    monthly_data = data.groupby(['Mois', 'Prestataire/ Frs', 'Annee'])['Deponse'].sum().reset_index()
    
    # Order months correctly
    month_order = ['JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 
                  'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE']
    monthly_data['Mois'] = pd.Categorical(monthly_data['Mois'], categories=month_order, ordered=True)
    
    # Sort by year and month for correct ordering
    monthly_data = monthly_data.sort_values(['Annee', 'Mois'])
    
    # Create a combined label for x-axis
    monthly_data['Month_Year'] = monthly_data.apply(
        lambda row: f"{month_order.index(row['Mois']) + 1}/{row['Annee']}", axis=1
    )
    
    fig = px.bar(
        monthly_data,
        x='Month_Year',
        y='Deponse',
        color='Prestataire/ Frs',
        barmode='group',
        title='Monthly Expenses by Service Provider',
        labels={'Deponse': 'Expenses (DH)', 'Month_Year': 'Month/Year'},
        hover_data=['Mois', 'Annee']
    )
    
    fig.update_layout(
        xaxis_tickangle=-45,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        width=1000,
        height=500,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )
    
    return fig

def create_budget_vs_actual_chart(data, budget_data):
    """Create a budget vs actual comparison chart"""
    # Calculate actual spending by month
    monthly_actual = data.groupby(['Annee', 'Mois'])['Deponse'].sum().reset_index()
    
    # Add budget data
    monthly_actual['Budget'] = monthly_actual.apply(
        lambda row: budget_data.get(str(row['Annee']), {}).get(row['Mois'], 0), 
        axis=1
    )
    
    # Sort by year and month for correct ordering
    month_order = ['JANVIER', 'FÉVRIER', 'MARS', 'AVRIL', 'MAI', 'JUIN', 
                  'JUILLET', 'AOÛT', 'SEPTEMBRE', 'OCTOBRE', 'NOVEMBRE', 'DÉCEMBRE']
    monthly_actual['Mois'] = pd.Categorical(monthly_actual['Mois'], categories=month_order, ordered=True)
    monthly_actual = monthly_actual.sort_values(['Annee', 'Mois'])
    
    # Create month-year label
    monthly_actual['Month_Year'] = monthly_actual.apply(
        lambda row: f"{month_order.index(row['Mois']) + 1}/{row['Annee']}", axis=1
    )
    
    fig = go.Figure()
    
    # Add actual spending
    fig.add_trace(go.Bar(
        x=monthly_actual['Month_Year'],
        y=monthly_actual['Deponse'],
        name='Actual Spending',
        marker_color='#0077b6'
    ))
    
    # Add budget line
    fig.add_trace(go.Scatter(
        x=monthly_actual['Month_Year'],
        y=monthly_actual['Budget'],
        name='Budget',
        mode='lines+markers',
        line=dict(color='red', width=3, dash='dash'),
        marker=dict(size=8)
    ))
    
    fig.update_layout(
        title='Budget vs Actual Spending',
        xaxis_title='Month/Year',
        yaxis_title='Amount (DH)',
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=500
    )
    
    return fig

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### 📈 Visual Analytics")
    
    # Tabs for different visualizations
    tab1, tab2, tab3, tab4 = st.tabs(["3D Surface", "3D Scatter", "Bar Chart", "Budget vs Actual"])
    
    with tab1:
        fig_surface = create_3d_surface_plot(filtered_df)
        st.plotly_chart(fig_surface, use_container_width=True)
        
    with tab2:
        fig_scatter = create_3d_scatter_plot(filtered_df)
        st.plotly_chart(fig_scatter, use_container_width=True)
        
    with tab3:
        fig_bar = create_advanced_bar_chart(filtered_df)
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with tab4:
        fig_budget = create_budget_vs_actual_chart(filtered_df, budget_data)
        st.plotly_chart(fig_budget, use_container_width=True)

with col2:
    st.markdown("### 🔍 Quick Insights")
    
    # Summary statistics
    total_expenses = filtered_df['Deponse'].sum()
    avg_monthly = filtered_df.groupby(['Annee', 'Mois'])['Deponse'].sum().mean()
    providers_count = filtered_df['Prestataire/ Frs'].nunique()
    cranes_count = filtered_df['Grue'].nunique()
    
    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
    st.metric("💰 Total Expenses", f"{total_expenses:,.0f} DH")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
    st.metric("📅 Avg Monthly", f"{avg_monthly:,.0f} DH")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
    st.metric("👥 Service Providers", providers_count)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
    st.metric("🏗️ Cranes", cranes_count)
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Top providers
    st.markdown("#### 🏆 Top Service Providers")
    top_providers = filtered_df.groupby('Prestataire/ Frs')['Deponse'].sum().nlargest(5)
    for provider, amount in top_providers.items():
        st.markdown(f"**{provider}**: {amount:,.0f} DH")
        st.markdown('<div class="progress-bar">', unsafe_allow_html=True)
        st.progress(float(amount/top_providers.max()), text=f"{float(amount/top_providers.max())*100:.1f}%")
        st.markdown('</div>', unsafe_allow_html=True)

# Additional analysis section
st.markdown("---")
st.markdown("### 📊 Detailed Analysis")

if selected_years:
    # Provider analysis
    st.markdown("#### 📋 Provider Performance")
    provider_summary = filtered_df.groupby('Prestataire/ Frs').agg({
        'Deponse': ['sum', 'mean', 'count'],
        'Grue': 'nunique'
    }).round(0)
    provider_summary.columns = ['Total Expenses', 'Avg Expense', 'Transaction Count', 'Unique Cranes']
    provider_summary = provider_summary.sort_values('Total Expenses', ascending=False)
    
    # Format the numbers with commas
    for col in ['Total Expenses', 'Avg Expense', 'Transaction Count']:
        provider_summary[col] = provider_summary[col].apply(lambda x: f"{x:,.0f}")
    
    st.dataframe(provider_summary, use_container_width=True)
    
    # Crane analysis
    st.markdown("#### 🏗️ Crane Utilization")
    crane_summary = filtered_df.groupby('Grue').agg({
        'Deponse': ['sum', 'mean'],
        'Prestataire/ Frs': 'nunique'
    }).round(0)
    crane_summary.columns = ['Total Expenses', 'Avg Expense', 'Service Providers']
    crane_summary = crane_summary.sort_values('Total Expenses', ascending=False)
    
    # Format the numbers with commas
    for col in ['Total Expenses', 'Avg Expense']:
        crane_summary[col] = crane_summary[col].apply(lambda x: f"{x:,.0f}")
    
    st.dataframe(crane_summary, use_container_width=True)

else:
    st.info("Select years to view detailed analysis")
