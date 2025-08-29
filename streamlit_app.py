import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

st.markdown(
    """
    <style>
        header[data-testid="stHeader"] {
            background-color: #e6f2ff; /* même bleu clair que la page */
        }
    </style>
    """,
    unsafe_allow_html=True
)


# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="Tableau de Bord - Dépenses des Grues",
    page_icon="📊",
    layout="wide"
)

# --- STYLES CSS PERSONNALISÉS ---
st.markdown("""
<style>
    /* Fond général */
    .stApp {
        background: linear-gradient(135deg, #eef2ff 0%, #f9fafb 100%);
        font-family: 'Segoe UI', sans-serif;
    }

    /* Titre principal */
    .title-header {
        color: #1e3a8a;
        text-align: center;
        padding: 20px;
        font-size: 2.2em;
        font-weight: 700;
    }

    /* Styles des boîtes de métriques */
    .stMetric {
        background: white;
        border-radius: 12px;
        padding: 18px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.05);
        text-align: center;
    }

    /* Expander */
    .streamlit-expanderHeader {
        background: #e0e7ff;
        border-radius: 8px;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        background: #e0e7ff;
        border-radius: 10px 10px 0 0;
        padding: 12px 24px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: #2563eb;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# --- HEADER ---
st.markdown('<h1 class="title-header">Tableau de Bord des Dépenses des Grues (2024-2025)</h1>', unsafe_allow_html=True)
st.markdown("_Prototype v1.0 - Analyse Professionnelle_")

# --- CHARGEMENT DES DONNÉES ---
@st.cache_data
def load_data(file):
    data = pd.read_excel(file)
    return data

upload_file = st.file_uploader("Importer le fichier Excel : GRUE PAR PRESTATAIRE", type=["xlsx"])
if upload_file is None:
    st.info("Veuillez importer un fichier Excel pour commencer l'analyse.")
    st.stop()

df = load_data(upload_file)

with st.expander("Aperçu des Données", expanded=False):
    st.dataframe(df, use_container_width=True)

# --- FONCTIONS DE VISUALISATION ---
def create_3d_surface_plot(data):
    """Créer un graphique en surface 3D des dépenses"""
    month_mapping = {
        'JANVIER': 1, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'JUIN': 6,
        'JUILLET': 7, 'AOÛT': 8, 'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DÉCEMBRE': 12
    }
    
    plot_data = data.copy()
    plot_data['Month_Num'] = plot_data['Mois'].map(month_mapping)
    
    # Create period labels with month and year
    plot_data['Period'] = plot_data['Mois'] + ' ' + plot_data['Annee'].astype(str)
    
    # Create a sorted list of all periods
    all_periods = []
    for year in sorted(plot_data['Annee'].unique()):
        for month_num in range(1, 13):
            month_name = list(month_mapping.keys())[list(month_mapping.values()).index(month_num)]
            period = f"{month_name} {year}"
            all_periods.append(period)
    
    # Create a mapping from period to index
    period_to_idx = {period: idx for idx, period in enumerate(all_periods)}
    
    # Add period index to data
    plot_data['Period_Idx'] = plot_data['Period'].map(period_to_idx)
    
    pivot_data = plot_data.pivot_table(
        values='Deponse',
        index=['Period_Idx', 'Period'],
        columns='Prestataire/ Frs',
        aggfunc='sum',
        fill_value=0
    ).reset_index().sort_values('Period_Idx')
    
    providers = pivot_data.columns[2:]
    
    X, Y = np.meshgrid(range(len(providers)), range(len(pivot_data)))
    Z = np.zeros((len(pivot_data), len(providers)))
    
    for i, provider in enumerate(providers):
        Z[:, i] = pivot_data[provider].values
    
    fig = go.Figure(data=[go.Surface(z=Z, x=X, y=Y, colorscale='Blues', opacity=0.9)])
    
    fig.update_layout(
        title='Surface 3D : Analyse des Dépenses',
        scene=dict(
            xaxis=dict(title='Prestataires', tickvals=list(range(len(providers))), 
                      ticktext=[p[:12] + '...' if len(p) > 12 else p for p in providers]),
            yaxis=dict(title='Période', tickvals=list(range(len(pivot_data))),
                      ticktext=[p.split(' ')[0][:3] + ' ' + p.split(' ')[1] for p in pivot_data['Period']]),
            zaxis=dict(title='Dépenses (DH)'),
            camera=dict(eye=dict(x=1.8, y=1.8, z=1.2))
        ),
        width=1200,
        height=700
    )
    return fig

def create_3d_scatter_plot(data):
    """Créer un nuage de points 3D"""
    month_mapping = {
        'JANVIER': 1, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'JUIN': 6,
        'JUILLET': 7, 'AOÛT': 8, 'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DÉCEMBRE': 12
    }
    
    plot_data = data.copy()
    plot_data['Month_Num'] = plot_data['Mois'].map(month_mapping)
    
    # Create period labels with month and year for hover
    plot_data['Period'] = plot_data['Mois'] + ' ' + plot_data['Annee'].astype(str)
    
    fig = px.scatter_3d(
        plot_data,
        x='Prestataire/ Frs',
        y='Month_Num',
        z='Deponse',
        color='Grue',
        size='Deponse',
        hover_data=['Period'],
        title='Nuage de Points 3D : Répartition des Dépenses',
        labels={
            'Prestataire/ Frs': 'Prestataire',
            'Month_Num': 'Mois',
            'Deponse': 'Dépenses (DH)',
            'Grue': 'Identifiant Grue'
        }
    )
    
    # Update y-axis to show month names instead of numbers
    month_names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    fig.update_layout(
        scene=dict(
            xaxis=dict(title='Prestataires', tickangle=45),
            yaxis=dict(title='Mois', tickvals=list(range(1,13)), ticktext=month_names),
            zaxis=dict(title='Dépenses (DH)'),
            camera=dict(eye=dict(x=1.8, y=1.8, z=1.2))
        ),
        width=1200,
        height=700
    )
    return fig

def create_advanced_bar_chart(data):
    """Créer un graphique en barres professionnel"""
    # Create period labels with month and year
    data['Period'] = data['Mois'] + ' ' + data['Annee'].astype(str)
    
    monthly_data = data.groupby(['Period', 'Prestataire/ Frs', 'Annee', 'Mois'])['Deponse'].sum().reset_index()
    
    # Create proper ordering for months
    month_mapping = {
        'JANVIER': 1, 'FÉVRIER': 2, 'MARS': 3, 'AVRIL': 4, 'MAI': 5, 'JUIN': 6,
        'JUILLET': 7, 'AOÛT': 8, 'SEPTEMBRE': 9, 'OCTOBRE': 10, 'NOVEMBRE': 11, 'DÉCEMBRE': 12
    }
    
    monthly_data['Month_Num'] = monthly_data['Mois'].map(month_mapping)
    monthly_data = monthly_data.sort_values(['Annee', 'Month_Num'])
    
    fig = px.bar(
        monthly_data,
        x='Period',
        y='Deponse',
        color='Prestataire/ Frs',
        barmode='group',
        title='Dépenses Mensuelles par Prestataire',
        labels={'Deponse': 'Dépenses (DH)', 'Period': 'Mois et Année'},
        hover_data=['Annee']
    )
    
    fig.update_layout(
        xaxis_tickangle=-45,
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        width=1200,
        height=600
    )
    return fig

# --- CONTENU PRINCIPAL ---
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### Analyse Visuelle")
    tab1, tab2, tab3 = st.tabs(["Surface 3D", "Nuage de Points 3D", "Graphique en Barres"])
    
    with tab1:
        st.plotly_chart(create_3d_surface_plot(df), use_container_width=True)
    with tab2:
        st.plotly_chart(create_3d_scatter_plot(df), use_container_width=True)
    with tab3:
        st.plotly_chart(create_advanced_bar_chart(df), use_container_width=True)

with col2:
    st.markdown("### Aperçu Rapide")
    
    total_expenses = df['Deponse'].sum()
    avg_monthly = df.groupby(['Annee', 'Mois'])['Deponse'].sum().mean()
    providers_count = df['Prestataire/ Frs'].nunique()
    cranes_count = df['Grue'].nunique()
    
    st.metric("Dépenses Totales", f"{total_expenses:,.0f} DH")
    st.metric("Moyenne Mensuelle", f"{avg_monthly:,.0f} DH")
    st.metric("Prestataires", providers_count)
    st.metric("Grues", cranes_count)
    
    st.markdown("#### Meilleurs Prestataires")
    top_providers = df.groupby('Prestataire/ Frs')['Deponse'].sum().nlargest(5)
    for provider, amount in top_providers.items():
        st.progress(amount/top_providers.max(), text=f"{provider}: {amount:,.0f} DH")

# --- ANALYSE DÉTAILLÉE ---
st.markdown("---")
st.markdown("### Analyse Détaillée")

years = sorted(df['Annee'].unique())
selected_years = st.multiselect("Sélectionner l'année(s) à analyser", options=years, default=years)

if selected_years:
    filtered_df = df[df['Annee'].isin(selected_years)]
    
    st.markdown("#### Performance des Prestataires")
    provider_summary = filtered_df.groupby('Prestataire/ Frs').agg({
        'Deponse': ['sum', 'mean', 'count'],
        'Grue': 'nunique'
    }).round(0)
    provider_summary.columns = ['Dépenses Totales', 'Dépense Moyenne', 'Nombre Transactions', 'Nombre de Grues']
    st.dataframe(provider_summary.sort_values('Dépenses Totales', ascending=False))
    
    st.markdown("#### Utilisation des Grues")
    crane_summary = filtered_df.groupby('Grue').agg({
        'Deponse': ['sum', 'mean'],
        'Prestataire/ Frs': 'nunique'
    }).round(0)
    crane_summary.columns = ['Dépenses Totales', 'Dépense Moyenne', 'Prestataires Différents']
    st.dataframe(crane_summary.sort_values('Dépenses Totales', ascending=False))
else:
    st.info("Sélectionnez une ou plusieurs années pour voir l'analyse détaillée.")