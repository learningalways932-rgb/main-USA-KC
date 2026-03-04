import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import warnings
warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(page_title="USA Sales Report", layout="wide", page_icon="🇺🇸")

# Custom CSS
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Serif+Display&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Sans', sans-serif;
    }
    .main { padding: 0rem 1rem; }

    .metric-card {
        background: linear-gradient(135deg, #0f1b2d 0%, #1a2e4a 100%);
        border: 1px solid rgba(62, 175, 189, 0.25);
        border-radius: 16px;
        padding: 22px 18px;
        text-align: center;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3), inset 0 1px 0 rgba(255,255,255,0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 40px rgba(62,175,189,0.2);
    }
    .metric-label {
        font-size: 11px;
        font-weight: 600;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        color: #6ec6d0;
        margin-bottom: 8px;
    }
    .metric-value {
        font-size: 28px;
        font-weight: 700;
        color: #ffffff;
        font-family: 'DM Serif Display', serif;
        letter-spacing: -0.5px;
    }
    .metric-icon { font-size: 22px; margin-bottom: 6px; }

    .table-container {
        background: linear-gradient(160deg, #0f1b2d 0%, #152236 100%);
        border: 1px solid rgba(62, 175, 189, 0.15);
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 4px 24px rgba(0,0,0,0.25);
        margin-bottom: 24px;
    }
    .table-title {
        font-size: 13px;
        font-weight: 600;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        color: #3eafbd;
        margin-bottom: 14px;
        padding-bottom: 10px;
        border-bottom: 1px solid rgba(62,175,189,0.2);
    }

    .chart-container {
        background: linear-gradient(160deg, #0f1b2d 0%, #152236 100%);
        border: 1px solid rgba(62, 175, 189, 0.15);
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 24px rgba(0,0,0,0.25);
        margin-bottom: 24px;
    }
    .chart-title {
        font-size: 13px;
        font-weight: 600;
        letter-spacing: 1.2px;
        text-transform: uppercase;
        color: #3eafbd;
        margin-bottom: 4px;
    }

    div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }
    div[data-testid="stDataFrame"] th {
        background-color: #1a3050 !important;
        color: #6ec6d0 !important;
        font-weight: 600 !important;
        font-size: 11px !important;
        letter-spacing: 0.8px !important;
        text-transform: uppercase !important;
        padding: 12px 8px !important;
        border-right: 1px solid rgba(62,175,189,0.15) !important;
    }
    div[data-testid="stDataFrame"] td {
        color: #d0e8ef !important;
        font-size: 12px !important;
        padding: 9px 8px !important;
    }

    .sidebar-header {
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 2px;
        text-transform: uppercase;
        color: #3eafbd;
        margin-bottom: 1rem;
    }

    .stAlert { display: none; }
    </style>
    """, unsafe_allow_html=True)

# ── Palette & chart helpers ────────────────────────────────────────────────────
TEAL_PALETTE = [
    "#3EAFBD", "#2E86AB", "#52C4C0", "#74D4C4", "#1A6B8A",
    "#E9C46A", "#F4A261", "#E76F51", "#264653", "#2A9D8F",
    "#95D5B2", "#B7E4C7", "#457B9D", "#0D3B66", "#6D3A7C"
]

PLOT_BG    = "#0f1b2d"
PAPER_BG   = "#0f1b2d"
GRID_COLOR = "rgba(62,175,189,0.08)"
AXIS_COLOR = "rgba(62,175,189,0.3)"
TEXT_COLOR = "#d0e8ef"
TITLE_COLOR= "#3eafbd"

AXIS_FONT  = dict(size=12, color=TEXT_COLOR,  family="DM Sans, sans-serif")
TICK_FONT  = dict(size=11, color=TEXT_COLOR,  family="DM Sans, sans-serif")
LABEL_FONT = dict(size=10, color="#ffffff",   family="DM Sans, sans-serif")
TITLE_FONT = dict(size=15, color=TITLE_COLOR, family="DM Serif Display, serif")

def _dark_layout(fig, xaxis_title, yaxis_title, extra_xaxis=None, height=500):
    xax = dict(
        title=dict(text=xaxis_title, font=AXIS_FONT, standoff=12),
        tickfont=TICK_FONT,
        tickangle=-90,
        linecolor=AXIS_COLOR,
        linewidth=1,
        showgrid=False,
        ticks="outside",
        ticklen=4,
        tickcolor=AXIS_COLOR,
        automargin=True,
    )
    if extra_xaxis:
        xax.update(extra_xaxis)

    fig.update_layout(
        height=height,
        font=dict(family="DM Sans, sans-serif", size=11, color=TEXT_COLOR),
        title_font=TITLE_FONT,
        title_x=0.5,
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=PLOT_BG,
        margin=dict(t=70, b=140, l=70, r=30),
        xaxis=xax,
        yaxis=dict(
            title=dict(text=yaxis_title, font=AXIS_FONT, standoff=10),
            tickfont=TICK_FONT,
            linecolor=AXIS_COLOR,
            linewidth=1,
            gridcolor=GRID_COLOR,
            gridwidth=1,
            zeroline=False,
        ),
        coloraxis_showscale=False,
        showlegend=False,
    )
    fig.update_traces(
        textfont=dict(size=10, color="#ffffff", family="DM Sans, sans-serif"),
        textangle=0,
        textposition="outside",
        cliponaxis=False,
    )
    return fig


# ── Header ─────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='text-align:center; padding: 28px 0 10px 0;'>
  <div style='font-family:"DM Serif Display",serif; font-size:42px; color:#3eafbd; letter-spacing:-1px;'>🇺🇸 USA Sales Report</div>
  <div style='font-size:13px; color:#6ec6d0; letter-spacing:3px; text-transform:uppercase; margin-top:6px; font-weight:500;'>Comprehensive Sales Analytics Dashboard</div>
</div>
<hr style='border:none; border-top:1px solid rgba(62,175,189,0.2); margin:18px 0;'>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel File with Sheets 'A' and 'B'", type=['xlsx', 'xls'])


@st.cache_data(ttl=3600)
def load_and_process_data(uploaded_file):
    try:
        sheet_a = pd.read_excel(uploaded_file, sheet_name='A')
        sheet_b = pd.read_excel(uploaded_file, sheet_name='B')

        sheet_a.columns = sheet_a.columns.astype(str).str.strip()
        sheet_b.columns = sheet_b.columns.astype(str).str.strip()

        def find_column(df, possible_names):
            df_cols_upper = {col.upper().strip(): col for col in df.columns}
            for name in possible_names:
                if name.upper().strip() in df_cols_upper:
                    return df_cols_upper[name.upper().strip()]
            return None

        # Sheet A columns
        season_col      = find_column(sheet_a, ['SEASON', 'Season'])
        brand_col       = find_column(sheet_a, ['BRAND', 'Brand'])
        category_col    = find_column(sheet_a, ['CATEGORY', 'Category'])
        subcategory_col = find_column(sheet_a, ['Subcategory', 'SUBCATEGORY', 'Sub Category'])
        style_names_col = find_column(sheet_a, ['STYLE NAMES', 'Style Names', 'STYLE_NAMES'])
        style_no_col    = find_column(sheet_a, ['STYLE NO.', 'STYLE NO', 'Style No', 'STYLE_NUMBER'])
        color_col       = find_column(sheet_a, ['COLOR', 'Color'])
        colab_col_a     = find_column(sheet_a, ['COLAB', 'Colab'])
        initial_qty_col = find_column(sheet_a, ['INITIAL QTY', 'Initial Qty', 'INITIAL_QTY'])
        total_qty_col   = find_column(sheet_a, ['Total Qty', 'TOTAL QTY', 'Total_Qty'])
        balance_col     = find_column(sheet_a, ['Balance', 'BALANCE'])

        required_cols_a = {
            'BRAND': brand_col, 'SEASON': season_col, 'CATEGORY': category_col,
            'Subcategory': subcategory_col, 'COLOR': color_col, 'COLAB': colab_col_a,
            'INITIAL QTY': initial_qty_col, 'Total Qty': total_qty_col, 'Balance': balance_col
        }
        missing_a = [k for k, v in required_cols_a.items() if v is None]
        if missing_a:
            st.error(f"❌ Missing required columns in Sheet A: {', '.join(missing_a)}")
            st.info("Available columns in Sheet A: " + ", ".join(sheet_a.columns))
            st.stop()

        sheet_a_clean = pd.DataFrame({
            'SEASON':      sheet_a[season_col].astype(str).str.strip().str.upper(),
            'BRAND':       sheet_a[brand_col].astype(str).str.strip().str.upper(),
            'CATEGORY':    sheet_a[category_col].astype(str).str.strip().str.upper() if category_col else 'N/A',
            'SUBCATEGORY': sheet_a[subcategory_col].astype(str).str.strip().str.upper() if subcategory_col else 'N/A',
            'STYLE_NAMES': sheet_a[style_names_col].astype(str).str.strip().str.upper() if style_names_col else 'N/A',
            'STYLE_NO':    sheet_a[style_no_col].astype(str).str.strip().str.upper() if style_no_col else 'N/A',
            'COLOR':       sheet_a[color_col].astype(str).str.strip().str.upper(),
            'COLAB':       sheet_a[colab_col_a].astype(str).str.strip().str.upper(),
            'INITIAL_QTY': pd.to_numeric(sheet_a[initial_qty_col], errors='coerce').fillna(0),
            'TOTAL_QTY':   pd.to_numeric(sheet_a[total_qty_col], errors='coerce').fillna(0),
            'BALANCE':     pd.to_numeric(sheet_a[balance_col], errors='coerce').fillna(0)
        })

        sheet_a_unique = sheet_a_clean.drop_duplicates(subset=['COLAB'], keep='first')

        # Sheet B columns
        website_col    = find_column(sheet_b, ['WEBSITE', 'Website'])
        sku_col        = find_column(sheet_b, ['SKU', 'Sku'])
        size_col       = find_column(sheet_b, ['SIZE (US)', 'SIZE', 'Size US', 'Size'])
        qty_col        = find_column(sheet_b, ['QTY', 'Qty', 'Quantity'])
        order_date_col = find_column(sheet_b, ['ORDER RECV DATE', 'Order Recv Date', 'ORDER_DATE'])
        colab_col_b    = find_column(sheet_b, ['COLAB', 'Colab'])

        required_cols_b = {
            'WEBSITE': website_col, 'SIZE (US)': size_col, 'QTY': qty_col,
            'ORDER RECV DATE': order_date_col, 'COLAB': colab_col_b
        }
        missing_b = [k for k, v in required_cols_b.items() if v is None]
        if missing_b:
            st.error(f"❌ Missing required columns in Sheet B: {', '.join(missing_b)}")
            st.info("Available columns in Sheet B: " + ", ".join(sheet_b.columns))
            st.stop()

        # ── Build Sheet B RAW (one row per order line — DO NOT aggregate yet) ──
        sheet_b_raw = pd.DataFrame({
            'WEBSITE':    sheet_b[website_col].astype(str).str.strip().str.upper(),
            'SKU':        sheet_b[sku_col].astype(str).str.strip().str.upper() if sku_col else 'N/A',
            'SIZE_US':    sheet_b[size_col].astype(str).str.strip().str.upper(),
            'QTY':        pd.to_numeric(sheet_b[qty_col], errors='coerce').fillna(0),
            'ORDER_DATE': pd.to_datetime(sheet_b[order_date_col], errors='coerce'),
            'COLAB':      sheet_b[colab_col_b].astype(str).str.strip().str.upper()
        })

        # Add MONTH_YEAR label to raw Sheet B for filtering
        sheet_b_raw['MONTH_NUM']   = sheet_b_raw['ORDER_DATE'].dt.month
        sheet_b_raw['YEAR_NUM']    = sheet_b_raw['ORDER_DATE'].dt.year
        sheet_b_raw['MONTH_YEAR']  = sheet_b_raw['ORDER_DATE'].dt.strftime('%b-%y')

        # ── Aggregate Sheet B by COLAB for the merge (to get one row per COLAB) ──
        sheet_b_agg = sheet_b_raw.groupby('COLAB').agg(
            WEBSITE=('WEBSITE', 'first'),
            QTY=('QTY', 'sum'),
            ORDER_DATE=('ORDER_DATE', 'first')
        ).reset_index()

        # RIGHT JOIN: Sheet A unique → Sheet B aggregated
        merged_df = pd.merge(
            sheet_a_unique, sheet_b_agg,
            on='COLAB', how='right', suffixes=('_a', '_b')
        )

        for col in ['INITIAL_QTY', 'TOTAL_QTY', 'BALANCE', 'QTY']:
            merged_df[col] = merged_df[col].fillna(0)

        merged_df['SALES_PERCENTAGE'] = np.where(
            merged_df['INITIAL_QTY'] > 0,
            (merged_df['TOTAL_QTY'] / merged_df['INITIAL_QTY']) * 100, 0
        )
        merged_df['RETURN_PERCENTAGE'] = 0.0
        merged_df['MONTH_YEAR']  = merged_df['ORDER_DATE'].dt.strftime('%b-%y')
        merged_df['YEAR_MONTH']  = merged_df['ORDER_DATE'].dt.to_period('M')

        # Return raw Sheet B so charts can aggregate correctly
        return merged_df, sheet_a_unique, sheet_b_raw

    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        import traceback
        st.write("Detailed error:", traceback.format_exc())
        st.stop()


if uploaded_file is not None:
    try:
        with st.spinner('🔍 Loading and processing data...'):
            merged_df, sheet_a_unique, sheet_b_raw = load_and_process_data(uploaded_file)

        total_initial_qty = sheet_a_unique['INITIAL_QTY'].sum()
        total_qty_sold    = sheet_a_unique['TOTAL_QTY'].sum()
        total_balance     = sheet_a_unique['BALANCE'].sum()
        sales_pct         = (total_qty_sold / total_initial_qty * 100) if total_initial_qty > 0 else 0
        return_pct        = 0.0

        st.success(f"✅ Data loaded successfully! {len(merged_df):,} records processed")
        st.markdown("<hr style='border:none;border-top:1px solid rgba(62,175,189,0.15);margin:14px 0;'>", unsafe_allow_html=True)

        # ── Sidebar ────────────────────────────────────────────────────────────
        st.sidebar.markdown("<div class='sidebar-header'>🔍 Filter Options</div>", unsafe_allow_html=True)

        st.sidebar.markdown("### 📊 Sort Tables By")
        sort_column = st.sidebar.selectbox(
            "Select measure to sort by",
            ['Total Qty', 'Initial Qty', 'Balance', 'Sales%'],
            key='sort_measure'
        )
        sort_order = st.sidebar.radio("Sort order", ['Descending', 'Ascending'], horizontal=True)
        st.sidebar.markdown("---")

        # ── Dimension lists ────────────────────────────────────────────────────
        brands        = sorted(sheet_a_unique['BRAND'].dropna().unique())
        seasons       = sorted(sheet_a_unique['SEASON'].dropna().unique())
        categories    = sorted(sheet_a_unique['CATEGORY'].dropna().unique())
        subcategories = sorted(sheet_a_unique['SUBCATEGORY'].dropna().unique())
        colors        = sorted(sheet_a_unique['COLOR'].dropna().unique())
        colabs        = sorted(sheet_a_unique['COLAB'].dropna().unique())
        websites      = sorted(sheet_b_raw['WEBSITE'].dropna().unique())
        sizes         = sorted(
            sheet_b_raw[
                sheet_b_raw['SIZE_US'].notna() &
                (sheet_b_raw['SIZE_US'].str.strip() != '') &
                (~sheet_b_raw['SIZE_US'].str.upper().isin(['NAN', '0', 'NONE', 'N/A']))
            ]['SIZE_US'].unique()
        )

        # Month-Year ordered chronologically
        my_df = sheet_b_raw[sheet_b_raw['ORDER_DATE'].notna()].copy()
        if not my_df.empty:
            my_df_agg = (
                my_df.groupby(['MONTH_NUM', 'YEAR_NUM', 'MONTH_YEAR'])
                .size().reset_index()
                .sort_values(['YEAR_NUM', 'MONTH_NUM'])
            )
            month_years = my_df_agg['MONTH_YEAR'].tolist()
        else:
            month_years = []

        st.sidebar.markdown("### 🎯 Filter Data")

        # Sheet A dimensions
        selected_brands        = st.sidebar.multiselect("Brand",        ['All'] + brands,        default='All')
        selected_seasons       = st.sidebar.multiselect("Season",       ['All'] + seasons,       default='All')
        selected_categories    = st.sidebar.multiselect("Category",     ['All'] + categories,    default='All')
        selected_subcategories = st.sidebar.multiselect("Subcategory",  ['All'] + subcategories, default='All')
        selected_colors        = st.sidebar.multiselect("Color",        ['All'] + colors,        default='All')
        selected_colabs        = st.sidebar.multiselect("Colab",        ['All'] + colabs,        default='All')

        st.sidebar.markdown("---")

        # Sheet B dimensions
        selected_websites      = st.sidebar.multiselect("Website",      ['All'] + websites,      default='All')
        selected_sizes         = st.sidebar.multiselect("Size (US)",    ['All'] + sizes,         default='All')
        selected_month_years   = st.sidebar.multiselect("Month-Year",   ['All'] + month_years,   default='All')

        # ── Apply Sheet A filters → get valid COLABs ───────────────────────────
        filtered_colabs = sheet_a_unique.copy()
        if 'All' not in selected_brands        and selected_brands:
            filtered_colabs = filtered_colabs[filtered_colabs['BRAND'].isin(selected_brands)]
        if 'All' not in selected_seasons       and selected_seasons:
            filtered_colabs = filtered_colabs[filtered_colabs['SEASON'].isin(selected_seasons)]
        if 'All' not in selected_categories    and selected_categories:
            filtered_colabs = filtered_colabs[filtered_colabs['CATEGORY'].isin(selected_categories)]
        if 'All' not in selected_subcategories and selected_subcategories:
            filtered_colabs = filtered_colabs[filtered_colabs['SUBCATEGORY'].isin(selected_subcategories)]
        if 'All' not in selected_colors        and selected_colors:
            filtered_colabs = filtered_colabs[filtered_colabs['COLOR'].isin(selected_colors)]
        if 'All' not in selected_colabs        and selected_colabs:
            filtered_colabs = filtered_colabs[filtered_colabs['COLAB'].isin(selected_colabs)]

        valid_colabs = set(filtered_colabs['COLAB'].unique())

        # ── Apply Sheet B filters ──────────────────────────────────────────────
        filtered_b = sheet_b_raw[sheet_b_raw['COLAB'].isin(valid_colabs)].copy()
        if 'All' not in selected_websites    and selected_websites:
            filtered_b = filtered_b[filtered_b['WEBSITE'].isin(selected_websites)]
        if 'All' not in selected_sizes       and selected_sizes:
            filtered_b = filtered_b[filtered_b['SIZE_US'].isin(selected_sizes)]
        if 'All' not in selected_month_years and selected_month_years:
            filtered_b = filtered_b[filtered_b['MONTH_YEAR'].isin(selected_month_years)]

        # Filter merged_df (used for dimension tables only)
        filtered_df = merged_df[merged_df['COLAB'].isin(valid_colabs)].copy()
        if 'All' not in selected_websites    and selected_websites:
            filtered_df = filtered_df[filtered_df['WEBSITE'].isin(selected_websites)]

        st.sidebar.markdown("---")
        st.sidebar.info(f"""
        **Filter Summary:**
        - 📊 COLABs: {len(valid_colabs):,}
        - 🏷️ Brands: {filtered_colabs['BRAND'].nunique()}
        - 🌐 Websites: {filtered_b['WEBSITE'].nunique()}
        - 📏 Sizes: {filtered_b['SIZE_US'].nunique()}
        - 📅 Months: {filtered_b['MONTH_YEAR'].nunique()}
        """)

        if len(filtered_df) == 0:
            st.warning("⚠️ No data available for the selected filters.")
        else:
            # ── KPIs ────────────────────────────────────────────────────────────
            st.markdown("""
            <div style='font-family:"DM Serif Display",serif; font-size:22px; color:#3eafbd;
                 letter-spacing:-0.5px; margin-bottom:18px;'>
              📊 Key Performance Indicators
            </div>""", unsafe_allow_html=True)

            filtered_sheet_a = sheet_a_unique[sheet_a_unique['COLAB'].isin(valid_colabs)]
            f_init  = filtered_sheet_a['INITIAL_QTY'].sum()
            f_sold  = filtered_sheet_a['TOTAL_QTY'].sum()
            f_bal   = filtered_sheet_a['BALANCE'].sum()
            f_spct  = (f_sold / f_init * 100) if f_init > 0 else 0

            col1, col2, col3, col4, col5 = st.columns(5)
            kpis = [
                (col1, "📦", "Initial Qty",    f"{f_init:,.0f}"),
                (col2, "💰", "Total Qty Sold", f"{f_sold:,.0f}"),
                (col3, "⚖️", "Balance Qty",   f"{f_bal:,.0f}"),
                (col4, "🔄", "Return %",       f"{return_pct:.1f}%"),
                (col5, "📈", "Sales %",        f"{f_spct:.1f}%"),
            ]
            for col, icon, label, value in kpis:
                with col:
                    st.markdown(f"""
                    <div class='metric-card'>
                      <div class='metric-icon'>{icon}</div>
                      <div class='metric-label'>{label}</div>
                      <div class='metric-value'>{value}</div>
                    </div>""", unsafe_allow_html=True)

            st.markdown("<hr style='border:none;border-top:1px solid rgba(62,175,189,0.15);margin:24px 0;'>", unsafe_allow_html=True)

            # ── Helper ─────────────────────────────────────────────────────────
            def analyze_group(group_col, display_name):
                if group_col not in filtered_sheet_a.columns:
                    return pd.DataFrame()
                grouped = filtered_sheet_a.groupby(group_col, observed=True).agg(
                    INITIAL_QTY=('INITIAL_QTY', 'sum'),
                    TOTAL_QTY=('TOTAL_QTY', 'sum'),
                    BALANCE=('BALANCE', 'sum')
                ).reset_index()
                grouped['SALES_PERCENTAGE'] = np.where(
                    grouped['INITIAL_QTY'] > 0,
                    (grouped['TOTAL_QTY'] / grouped['INITIAL_QTY']) * 100, 0
                )
                sort_map = {
                    'Total Qty': 'TOTAL_QTY', 'Initial Qty': 'INITIAL_QTY',
                    'Balance': 'BALANCE', 'Sales%': 'SALES_PERCENTAGE'
                }
                grouped = grouped.sort_values(sort_map[sort_column], ascending=(sort_order == 'Ascending'))
                grouped.rename(columns={group_col: display_name}, inplace=True)
                return grouped

            # ── Distribution Tables ────────────────────────────────────────────
            st.markdown("""
            <div style='font-family:"DM Serif Display",serif; font-size:22px; color:#3eafbd;
                 letter-spacing:-0.5px; margin-bottom:18px;'>
              📋 Sales Distribution Tables
            </div>""", unsafe_allow_html=True)

            tables_config = [
                ('BRAND', 'Brand'), ('SEASON', 'Season'),
                ('CATEGORY', 'Category'), ('SUBCATEGORY', 'Subcategory'),
                ('COLOR', 'Color'), ('COLAB', 'Colab')
            ]

            for i in range(0, len(tables_config), 2):
                cols = st.columns(2)
                for j in range(2):
                    if i + j < len(tables_config):
                        col_name, display_name = tables_config[i + j]
                        with cols[j]:
                            st.markdown(f"""
                            <div class='table-container'>
                              <div class='table-title'>◈ {display_name} wise Distribution</div>
                            </div>""", unsafe_allow_html=True)

                            table_data = analyze_group(col_name, display_name)
                            if not table_data.empty:
                                col_cfg = {
                                    display_name:        st.column_config.TextColumn(display_name),
                                    'INITIAL_QTY':       st.column_config.NumberColumn('Initial Qty',    format="%d"),
                                    'TOTAL_QTY':         st.column_config.NumberColumn('Total Qty Sold', format="%d"),
                                    'BALANCE':           st.column_config.NumberColumn('Balance Qty',    format="%d"),
                                    'SALES_PERCENTAGE':  st.column_config.NumberColumn('Sales %',        format="%.1f%%"),
                                }
                                st.dataframe(table_data, column_config=col_cfg,
                                             hide_index=True, use_container_width=True, height=420)
                            else:
                                st.info(f"No data for {display_name}")

            st.markdown("<hr style='border:none;border-top:1px solid rgba(62,175,189,0.15);margin:24px 0;'>", unsafe_allow_html=True)

            # ── Visual Charts ──────────────────────────────────────────────────
            st.markdown("""
            <div style='font-family:"DM Serif Display",serif; font-size:22px; color:#3eafbd;
                 letter-spacing:-0.5px; margin-bottom:18px;'>
              📊 Visual Analytics
            </div>""", unsafe_allow_html=True)

            # ── CHART 1: Marketplace ───────────────────────────────────────────
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            st.markdown("<div class='chart-title'>🌐 Marketplace wise Qty Sold</div>", unsafe_allow_html=True)

            website_data = (
                filtered_b[filtered_b['WEBSITE'].notna() &
                           (filtered_b['WEBSITE'].str.strip() != '') &
                           (filtered_b['WEBSITE'].str.upper() != 'NAN')]
                .groupby('WEBSITE')['QTY'].sum()
                .reset_index()
                .sort_values('QTY', ascending=False)
            )

            if not website_data.empty:
                n = len(website_data)
                colors_bars = [TEAL_PALETTE[i % len(TEAL_PALETTE)] for i in range(n)]
                fig_ws = go.Figure(go.Bar(
                    x=website_data['WEBSITE'],
                    y=website_data['QTY'],
                    text=website_data['QTY'].apply(lambda v: f"{v:,.0f}"),
                    marker=dict(
                        color=colors_bars,
                        line=dict(color='rgba(255,255,255,0.06)', width=1),
                        cornerradius=6,
                    ),
                ))
                fig_ws.update_layout(title="Sales by Marketplace")
                fig_ws = _dark_layout(
                    fig_ws, "Marketplace", "Quantity Sold",
                    extra_xaxis={'categoryorder': 'array',
                                 'categoryarray': website_data['WEBSITE'].tolist()}
                )
                st.plotly_chart(fig_ws, use_container_width=True)
            else:
                st.info("No marketplace data available")
            st.markdown("</div>", unsafe_allow_html=True)

            # ── CHART 2: Size ──────────────────────────────────────────────────
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            st.markdown("<div class='chart-title'>📏 Size wise Qty Distribution</div>", unsafe_allow_html=True)

            size_data = (
                filtered_b[
                    filtered_b['SIZE_US'].notna() &
                    (filtered_b['SIZE_US'].str.strip() != '') &
                    (~filtered_b['SIZE_US'].str.upper().isin(['NAN', '0', 'NONE', 'N/A']))
                ]
                .groupby('SIZE_US')['QTY'].sum()
                .reset_index()
                .sort_values('QTY', ascending=False)
            )

            if not size_data.empty:
                size_data['SIZE_LABEL'] = size_data['SIZE_US'].astype(str).str.strip()
                all_size_labels = size_data['SIZE_LABEL'].tolist()

                import plotly.colors as pc
                gradient = pc.sample_colorscale(
                    [[0, "#3EAFBD"], [0.4, "#2A9D8F"], [0.8, "#52B788"], [1.0, "#E9C46A"]],
                    max(len(size_data), 2)
                )
                fig_sz = go.Figure(go.Bar(
                    x=size_data['SIZE_LABEL'],
                    y=size_data['QTY'],
                    text=size_data['QTY'].apply(lambda v: f"{v:,.0f}"),
                    marker=dict(
                        color=gradient,
                        line=dict(color='rgba(255,255,255,0.06)', width=1),
                        cornerradius=6,
                    ),
                ))
                fig_sz.update_layout(title="Sales by Size (US)")
                fig_sz = _dark_layout(
                    fig_sz, "Size (US)", "Quantity Sold",
                    extra_xaxis={
                        'type': 'category',
                        'categoryorder': 'array',
                        'categoryarray': all_size_labels,
                        'tickmode': 'array',
                        'tickvals': all_size_labels,
                        'ticktext': all_size_labels,
                    },
                    height=520
                )
                st.plotly_chart(fig_sz, use_container_width=True)
            else:
                st.info("No size data available")
            st.markdown("</div>", unsafe_allow_html=True)

            # ── CHART 3: Month-Year ────────────────────────────────────────────
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            st.markdown("<div class='chart-title'>📅 Month-Year wise Qty Distribution</div>", unsafe_allow_html=True)

            monthly_b = filtered_b[filtered_b['ORDER_DATE'].notna()].copy()

            if not monthly_b.empty:
                monthly_b['MONTH_NUM']   = monthly_b['ORDER_DATE'].dt.month
                monthly_b['YEAR_NUM']    = monthly_b['ORDER_DATE'].dt.year
                monthly_b['MONTH_LABEL'] = monthly_b['ORDER_DATE'].dt.strftime('%b-%y')

                monthly_agg = (
                    monthly_b.groupby(['MONTH_NUM', 'YEAR_NUM', 'MONTH_LABEL'])['QTY']
                    .sum()
                    .reset_index()
                    .sort_values(['YEAR_NUM', 'MONTH_NUM'])
                )
                ordered_labels = monthly_agg['MONTH_LABEL'].tolist()

                MONTH_COLORS = {
                    1: "#3EAFBD", 2: "#2E86AB", 3: "#2196A6",
                    4: "#2ECC9A", 5: "#52B788", 6: "#95D5B2",
                    7: "#E9C46A", 8: "#F4A261", 9: "#E76F51",
                    10: "#C1440E", 11: "#6D3A7C", 12: "#457B9D",
                }
                bar_colors = [MONTH_COLORS.get(m, "#3EAFBD") for m in monthly_agg['MONTH_NUM']]

                fig_mo = go.Figure(go.Bar(
                    x=monthly_agg['MONTH_LABEL'],
                    y=monthly_agg['QTY'],
                    text=monthly_agg['QTY'].apply(lambda v: f"{v:,.0f}"),
                    marker=dict(
                        color=bar_colors,
                        line=dict(color='rgba(255,255,255,0.06)', width=1),
                        cornerradius=5,
                    ),
                ))
                fig_mo.update_layout(title="Sales by Month-Year")
                fig_mo = _dark_layout(
                    fig_mo, "Month-Year", "Quantity Sold",
                    extra_xaxis={
                        'categoryorder': 'array',
                        'categoryarray': ordered_labels,
                    },
                    height=540
                )
                st.plotly_chart(fig_mo, use_container_width=True)
            else:
                st.info("No order date data available")
            st.markdown("</div>", unsafe_allow_html=True)

            # ── Raw data ───────────────────────────────────────────────────────
            with st.expander("🔍 View Raw Data"):
                st.dataframe(filtered_b, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error: {str(e)}")
        import traceback
        st.write("Detailed error:", traceback.format_exc())

else:
    st.markdown("""
    <div style='text-align:center; padding:40px 20px; color:#6ec6d0; font-size:15px;'>
      👆 Please upload an Excel file to begin analyzing your data.
    </div>""", unsafe_allow_html=True)

    with st.expander("📋 Required Excel File Structure"):
        st.markdown("""
        ### **Sheet A Columns (Required):**
        - `SEASON` · `BRAND` · `CATEGORY` · `Subcategory` · `COLOR` · `COLAB`
        - `INITIAL QTY` · `Total Qty` · `Balance`

        ### **Sheet B Columns (Required):**
        - `WEBSITE` · `SIZE (US)` · `QTY` · `ORDER RECV DATE` · `COLAB`

        ### **Sidebar Filters Available:**
        **From Sheet A:** Brand · Season · Category · Subcategory · Color · Colab
        
        **From Sheet B:** Website · Size (US) · Month-Year
        """)
