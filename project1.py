import streamlit as st
import pandas as pd
import os
import pypyodbc as odbc

import altair as alt

# Correct path based on script location
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
path = os.path.join(BASE_DIR, "static", "construction_materials.xlsx")

def get_connection():
    conn = odbc.connect(
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=MSI\SQLEXPRESS;"
        "DATABASE=constructionDB;"
        "Trusted_Connection=yes;"
    )

    return conn

@st.cache_data
def load_excel():
    xls = pd.ExcelFile(path)
    data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
    return data


data = load_excel()
df_materials = data.get("materials", pd.DataFrame())
df_sites = data.get("sites", pd.DataFrame())
df_vendors = data.get("vendors", pd.DataFrame())
df_movements = data.get("stock_movements", pd.DataFrame())
# sheet_name = st.selectbox('Select sheet : ', list(data.keys()))
# st.dataframe(data[sheet_name])

def overview():
    st.title("Overview")
    st.write("Construction Materials Management :)")

    st.subheader('Lists of Data')
    with st.expander("Data Schema", expanded=False):
        st.write(list(data.keys()))

    #Metrics of each
    col1,col2,col4 = st.columns(3)
    with col1 :
        st.metric('Total materials', len(df_materials))
    with col2:
        st.metric('Total Sites',len(df_sites))
    # with col3:
    #     st.metric('Total Vendors',len(df_vendors))
    with col4:
        st.metric('Total Movements',len(df_movements))

    st.divider()

    # ----- DAILY IN vs OUT: DYNAMIC CHART -----
    st.subheader("Quantity Movement")
    st.caption("Shows the quantity movement across all the sites")

    # Controls
    c1, c2, c3 = st.columns(3)
    with c1:
        granularity = st.selectbox("Time granularity", ["Day", "Week", "Month"])
    with c2:
        chart_type = st.selectbox("Chart type", ["Line", "Bar", "Area"])
    with c3:
        movement_filter = st.multiselect(
            "Movement types",
            ["IN", "OUT"],
            default=["IN", "OUT"],
        )

    # Prepare time column
    df_movements["movement_datetime"] = pd.to_datetime(df_movements["movement_datetime"])

    if granularity == "Day":
        df_movements["time_key"] = df_movements["movement_datetime"].dt.date
    elif granularity == "Week":
        # use year-week string
        df_movements["time_key"] = df_movements["movement_datetime"].dt.to_period("W").astype(str)
    elif granularity == "Month":
        # YYYY-MM
        df_movements["time_key"] = df_movements["movement_datetime"].dt.to_period("M").astype(str)

    # Filter by movement type
    if movement_filter:
        df_mov_filt = df_movements[df_movements["movement_type"].isin(movement_filter)].copy()
    else:
        df_mov_filt = df_movements.copy()

    # Group and pivot
    daily = (
        df_mov_filt.groupby(["time_key", "movement_type"], as_index=False)["quantity"]
        .sum()
    )

    if daily.empty:
        st.warning("No data for selected filters.")
    else:
        daily_pivot = (
            daily.pivot_table(index="time_key", columns="movement_type", values="quantity")
            .fillna(0)
            .sort_index()
        )

        if chart_type == "Line":
            st.line_chart(daily_pivot)
        elif chart_type =='Area':
            st.area_chart(daily_pivot)
        else:
            st.bar_chart(daily_pivot)

    # Optional: show base SQL for daily IN vs OUT (day-level)
    sql_inout_both = """
    SELECT 
        FORMAT(DATETRUNC(day, movement_datetime), 'dd-MM-yyyy') AS [date],
        movement_type,
        SUM(quantity) AS total_qty
    FROM dbo.construction_movements
    GROUP BY DATETRUNC(day, movement_datetime), movement_type
    ORDER BY [date];
    """
    show_inout_both = st.checkbox("show SQL code", value=False)

    if show_inout_both:
        st.code(sql_inout_both, language="sql")

    #####

        # ----- SECONDARY CHARTS: SIMPLE & INTERACTIVE -----

    left_col, right_col = st.columns(2)

    # 1) LEFT: Quantity by site (IN / OUT / Both)
    with left_col:
        st.subheader("**Quantity by site**")

        site_mov_type = st.radio(
            "Movement type (site)",
            ["IN", "OUT", "Both"],
            horizontal=True,
            key="site_mov_type_radio",
        )

        if site_mov_type == "Both":
            tmp_site = df_movements[df_movements["movement_type"].isin(["IN", "OUT"])].copy()
        else:
            tmp_site = df_movements[df_movements["movement_type"] == site_mov_type].copy()

        if tmp_site.empty:
            st.info(f"No {site_mov_type} data available for sites.")
        else:
            # group by site & movement_type
            site_group = (
                tmp_site.groupby(["site_id", "movement_type"], as_index=False)["quantity"]
                .sum()
                .rename(columns={"quantity": "total_qty"})
            )

            # join site names
            if not df_sites.empty:
                site_group = site_group.merge(
                    df_sites[["site_id", "site_name"]],
                    on="site_id",
                    how="left",
                )

            # pivot so we get columns IN, OUT (or just one if single type)
            site_pivot = (
                site_group.pivot_table(
                    index="site_name",
                    columns="movement_type",
                    values="total_qty",
                    fill_value=0,
                )
                .sort_values(by=list(site_group["movement_type"].unique()), ascending=False)
            )

            # table: show IN, OUT, and Total if both
            site_table = site_pivot.copy()
            if "IN" in site_table.columns and "OUT" in site_table.columns:
                site_table["Total"] = site_table["IN"] + site_table["OUT"]

            st.dataframe(site_table, use_container_width=True, height=220)

            # chart: bar chart; if Both → 2 bars per site, else → 1 bar per site
            st.bar_chart(site_pivot)

    # 2) RIGHT: Quantity by material (IN / OUT / Both)
    with right_col:
        st.subheader("**Quantity by material**")

        mat_mov_type = st.radio(
            "Movement type (material)",
            ["IN", "OUT", "Both"],
            horizontal=True,
            key="mat_mov_type_radio",
        )

        if mat_mov_type == "Both":
            tmp_mat = df_movements[df_movements["movement_type"].isin(["IN", "OUT"])].copy()
        else:
            tmp_mat = df_movements[df_movements["movement_type"] == mat_mov_type].copy()

        if tmp_mat.empty:
            st.info(f"No {mat_mov_type} data available for materials.")
        else:
            # group by material & movement_type
            mat_group = (
                tmp_mat.groupby(["material_id", "movement_type"], as_index=False)["quantity"]
                .sum()
                .rename(columns={"quantity": "total_qty"})
            )

            # join material names
            if not df_materials.empty:
                mat_group = mat_group.merge(
                    df_materials[["material_id", "material_name", "category"]],
                    on="material_id",
                    how="left",
                )

            # pivot: index = material_name, cols = IN/OUT
            mat_pivot = (
                mat_group.pivot_table(
                    index="material_name",
                    columns="movement_type",
                    values="total_qty",
                    fill_value=0,
                )
            )

            # sort by total across available movement types
            mat_pivot["Total"] = mat_pivot.sum(axis=1)
            mat_pivot = mat_pivot.sort_values("Total", ascending=False).head(10)

            # table: include category via merge back
            mat_table = mat_pivot.drop(columns=["Total"]).copy()
            mat_table = mat_table.merge(
                mat_group[["material_name", "category"]].drop_duplicates(),
                on="material_name",
                how="left",
            )
            # reorder columns: category first
            cols = ["category"] + [c for c in mat_table.columns if c not in ["category"]]
            mat_table = mat_table[cols]

            st.dataframe(mat_table, use_container_width=True, height=220)

            # chart: bar chart; if Both → 2 bars per material, else → 1
            mat_chart_data = mat_pivot.drop(columns=["Total"])
            st.bar_chart(mat_chart_data)


def masters():
    st.title("Master Data")
    # Master options
    master_options = []
    if "materials" in data:
        master_options.append("materials")
    if "sites" in data:
        master_options.append("sites")
    if "vendors" in data:
        master_options.append("vendors")
    if "activities" in data:
        master_options.append("activities")
    if not master_options:
        st.warning("No master sheets (materials/sites/vendors/activities) found in Excel.")
        return

    selected_master = st.selectbox("Select master table", master_options)

    df = data[selected_master].copy()

    st.subheader(f"{selected_master}")

    #Searching text
    search_text = st.text_input("Filter rows (any text search)", "")

    if search_text:
        mask = df.apply(
            lambda col: col.astype(str).str.contains(search_text, case=False, na=False)
        )
        df = df[mask.any(axis=1)]

    st.dataframe(df, use_container_width=True)

    st.title('SQL code here: ')
    user_query = st.text_area("Write your SQL query here", "SELECT TOP 10 * FROM dbo.construction_materials")

    if st.button("Run SQL"):
        conn = get_connection()
        df = pd.read_sql(user_query, conn)
        st.dataframe(df, use_container_width=True)
        conn.close()


# SIDE BAR
st.sidebar.title('Navigation')
page = st.sidebar.radio(
    'Go to',
    ('Overview','Master')
)

if page == 'Overview':
    overview()
elif page =='Master':
    masters()