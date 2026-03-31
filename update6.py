import math, folium, requests, time, io, os, sys
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium

# =========================================================
# EXE PATH RESOLVER
# =========================================================
def resolve_path(path):
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))

    candidates = [
        os.path.join(base_dir, path),
        os.path.join(base_dir, "_internal", path),
    ]

    if hasattr(sys, "_MEIPASS"):
        candidates.append(os.path.join(sys._MEIPASS, path))

    for p in candidates:
        if os.path.exists(p):
            return p

    return os.path.join(base_dir, path)

st.set_page_config(layout="wide", page_title="FTTH Smart Node Checker")

MAX_DISTANCE = 350
MAX_CAPACITY = 16
TOP_FALLBACK_NODES = 3

# =========================================================
# SESSION STATE
# =========================================================
for key in ["batch_done", "batch_summary_df", "batch_results", "single_res"]:
    if key not in st.session_state:
        st.session_state[key] = None if "df" in key or "res" in key else False

# =========================================================
# UTILS
# =========================================================
def clean_num(value):
    try:
        val = pd.to_numeric(str(value).replace("°", "").replace(",", "").strip(), errors="coerce")
        return val
    except:
        return None

@st.cache_data(ttl=3600)
def get_route_osrm(lat1, lon1, lat2, lon2):
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return (None, None)
    url = f"http://router.project-osrm.org/route/v1/driving/{lon1},{lat1};{lon2},{lat2}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=3).json()
        return (r['routes'][0]['distance'], r['routes'][0]['geometry']) if r['code'] == 'Ok' else (None, None)
    except:
        return (None, None)

def haversine(lat1, lon1, lat2, lon2):
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2):
        return 999999
    r = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi, dlambda = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlambda/2)**2
    return r * (2 * math.atan2(math.sqrt(a), math.sqrt(1-a)))

# =========================================================
# CORE LOGIC
# =========================================================
def analyze_one_customer(nodes_df, cust_name, cust_lat, cust_lon, connected_name, partner="-"):
    res = {
        "customer_name": cust_name, "cust_lat": cust_lat, "cust_lon": cust_lon, "partner": partner,
        "connected_name": connected_name or "-", "connected_status": "NOK",
        "connected_reason": "-", "connected_dist": "-",
        "connected_map_obj": None, "recommended_map_obj": None
    }
    if pd.isna(cust_lat) or pd.isna(cust_lon):
        res["connected_reason"] = "Location Error"
        return res

    search_name = str(connected_name).strip().upper() if connected_name and str(connected_name).strip() != "-" else None

    if search_name:
        node = nodes_df[nodes_df["node_name_upper"] == search_name]
        if not node.empty:
            n = node.iloc[0]
            s_dist = haversine(cust_lat, cust_lon, n["Latitude"], n["Longitude"])
            if s_dist > 800:
                stat, reas, d, g = "NOK", f"Over meter (Straight {int(s_dist)}m)", int(s_dist), None
            else:
                d, g = get_route_osrm(cust_lat, cust_lon, n["Latitude"], n["Longitude"])
                is_dist_ok = (d is not None and d <= MAX_DISTANCE)
                is_port_ok = (n["act"] < MAX_CAPACITY)
                if is_dist_ok and is_port_ok:
                    stat, reas = "Can Deploy", "OK"
                elif not is_dist_ok:
                    stat, reas = "NOK", f"Over meter {int(d) if d else 'Unknown'} m"
                else:
                    stat, reas = "NOK", f"Full Port ({int(n['act'])})"

            res.update({
                "connected_status": stat,
                "connected_reason": reas,
                "connected_dist": int(d) if d else "-"
            })
            res["connected_map_obj"] = {
                "name": n["node_name"],
                "lat": n["Latitude"],
                "lon": n["Longitude"],
                "dist": d,
                "geom": g,
                "status": stat,
                "reason": reas
            }
        else:
            res["connected_reason"] = "Node Not Found"

    if res["connected_status"] != "Can Deploy":
        temp_nodes = nodes_df.copy()
        temp_nodes["temp_dist"] = temp_nodes.apply(
            lambda r: haversine(cust_lat, cust_lon, r["Latitude"], r["Longitude"]), axis=1
        )
        candidates = temp_nodes[temp_nodes["temp_dist"] < 600].sort_values("temp_dist").head(TOP_FALLBACK_NODES)
        for _, n in candidates.iterrows():
            if search_name and n["node_name_upper"] == search_name:
                continue
            d, g = get_route_osrm(cust_lat, cust_lon, n["Latitude"], n["Longitude"])
            if d and d <= MAX_DISTANCE and n["act"] < MAX_CAPACITY:
                res["recommended_map_obj"] = {
                    "name": n["node_name"],
                    "lat": n["Latitude"],
                    "lon": n["Longitude"],
                    "dist": d,
                    "geom": g,
                    "status": "Can Deploy",
                    "reason": "OK"
                }
                break
    return res

def draw_map(res, conn=None, reco=None):
    if pd.isna(res["cust_lat"]) or pd.isna(res["cust_lon"]):
        return folium.Map(location=[16.8, 96.1], zoom_start=12)

    m = folium.Map(location=[res["cust_lat"], res["cust_lon"]], zoom_start=17)
    folium.Marker(
        [res["cust_lat"], res["cust_lon"]],
        tooltip=f"{res['customer_name']}",
        icon=folium.Icon(color="blue")
    ).add_to(m)

    for obj, color in [(conn, "red"), (reco, "green")]:
        if obj and obj.get("lat"):
            folium.Marker(
                [obj["lat"], obj["lon"]],
                tooltip=f"Node: {obj['name']}",
                icon=folium.Icon(color=color)
            ).add_to(m)
            if obj.get("geom"):
                folium.GeoJson(
                    obj["geom"],
                    style_function=lambda x, c=color: {"color": c, "weight": 5}
                ).add_to(m)
    return m

# =========================================================
# UI
# =========================================================
st.title("FTTH Smart Node Checker")
st.markdown("##### Powered by Zaw Min Htwe")
st.info("Survey Result အား Routing အကွာအဝေးကို လမ်းကြောင်းအတိုင်း တွက်ချက်ပေးပါသည်။")

t1, t2 = st.tabs(["Batch Check", "Single Check"])

with t1:
    st.caption("Fixed Data: nodes.csv, customers.xlsx, new_customers.xlsx")
    col_b1, col_b2 = st.columns(2)

    if col_b1.button("Run Batch", type="primary"):
        try:
            node_data = pd.read_csv(resolve_path("nodes.csv"))
            cust_data = pd.read_excel(resolve_path("customers.xlsx"))
            act_counts = cust_data.groupby("node_name").size().reset_index(name="act")
            nodes = node_data.merge(act_counts, on="node_name", how="left").fillna(0)
            nodes["node_name_upper"] = nodes["node_name"].astype(str).str.strip().str.upper()
            new_custs = pd.read_excel(resolve_path("new_customers.xlsx")).dropna(subset=['customer_name'])

            results, summary_rows = [], []
            prog = st.progress(0)

            for i, (_, r) in enumerate(new_custs.iterrows()):
                res = analyze_one_customer(
                    nodes,
                    r["customer_name"],
                    clean_num(r["lat"]),
                    clean_num(r["Long"]),
                    r.get("connected_node"),
                    str(r.get("Partner", "-"))
                )
                results.append(res)
                summary_rows.append({
                    "Customer Name": res["customer_name"],
                    "Partner": res["partner"],
                    "Lat": res["cust_lat"],
                    "Long": res["cust_lon"],
                    "Status": res["connected_status"],
                    "Reason": res["connected_reason"],
                    "Connected Node": res["connected_name"],
                    "Distance (m)": res["connected_dist"],
                    "Recommended": res["recommended_map_obj"]["name"] if res["recommended_map_obj"] else "-"
                })
                prog.progress((i + 1) / len(new_custs))

            st.session_state.batch_results = results
            st.session_state.batch_summary_df = pd.DataFrame(summary_rows)
            st.session_state.batch_done = True

        except Exception as e:
            st.error(f"Error: {e}")

    if col_b2.button("Clear Batch"):
        st.session_state.batch_done = False
        st.session_state.batch_summary_df = None
        st.session_state.batch_results = None
        st.rerun()

    if st.session_state.batch_done:
        st.markdown("### Filters")
        f_col1, f_col2, f_col3 = st.columns(3)
        df_full = st.session_state.batch_summary_df.copy()

        for c in ["Partner", "Status", "Reason"]:
            df_full[c] = df_full[c].astype(str).fillna("-")

        with f_col1:
            sel_p = st.multiselect("Partner", options=sorted(df_full["Partner"].unique()))
        with f_col2:
            sel_s = st.multiselect("Status", options=sorted(df_full["Status"].unique()))
        with f_col3:
            sel_r = st.multiselect("Reason", options=sorted(df_full["Reason"].unique()))

        f_df = df_full.copy()
        if sel_p:
            f_df = f_df[f_df["Partner"].isin(sel_p)]
        if sel_s:
            f_df = f_df[f_df["Status"].isin(sel_s)]
        if sel_r:
            f_df = f_df[f_df["Reason"].isin(sel_r)]

        st.dataframe(f_df, use_container_width=True)

        # Excel Download only
        xlsx_io = io.BytesIO()
        with pd.ExcelWriter(xlsx_io, engine='xlsxwriter') as writer:
            f_df.to_excel(writer, index=False, sheet_name='Report')
            for i, col in enumerate(f_df.columns):
                writer.sheets['Report'].set_column(i, i, 20)

        st.download_button("Download Excel Report", data=xlsx_io.getvalue(), file_name="Report.xlsx")

        st.divider()
        st.caption("Select Customer to View Detail")
        c_opts = f_df["Customer Name"].tolist()
        if c_opts:
            sel = st.selectbox("Select Customer to View Detail", options=c_opts, label_visibility="collapsed")
            res = next(r for r in st.session_state.batch_results if str(r["customer_name"]) == str(sel))

            st.markdown("## Customer Survey Result")
            st.divider()
            st.markdown(f"**Customer:** {res['customer_name']} | **Partner:** {res['partner']}")
            st.markdown(f"**Lat/Long:** {res['cust_lat']}, {res['cust_lon']}")
            st.markdown(f"**Connected Node:** {res['connected_name']}")
            st.markdown(f"**Status:** {res['connected_status']}")
            st.markdown(f"**Reason:** {res['connected_reason']}")
            st.markdown(f"**Distance Result:** {res['connected_dist']}m")
            st_folium(draw_map(res, res["connected_map_obj"], res["recommended_map_obj"]), height=500, width=1000, key=f"m_{sel}")

with t2:
    st.subheader("Single Customer Check")
    sl1, sl2 = st.columns(2)

    with sl1:
        s_na = st.text_input("Customer Name", key="sn_name")
        s_la = st.text_input("Latitude", key="sn_lat")

    with sl2:
        s_no = st.text_input("Connected Node", key="sn_node")
        s_lo = st.text_input("Longitude", key="sn_lon")

    sc1, sc2 = st.columns(2)

    if sc1.button("Run Single", type="primary"):
        if s_la and s_lo:
            cust_db = pd.read_excel(resolve_path("customers.xlsx"))
            n_s = pd.read_csv(resolve_path("nodes.csv")).merge(
                cust_db.groupby("node_name").size().reset_index(name="act"),
                on="node_name",
                how="left"
            ).fillna(0)
            n_s["node_name_upper"] = n_s["node_name"].astype(str).str.strip().str.upper()
            st.session_state.single_res = analyze_one_customer(
                n_s, s_na, clean_num(s_la), clean_num(s_lo), s_no
            )

    if sc2.button("Clear Single"):
        st.session_state.single_res = None
        st.rerun()

    if st.session_state.single_res:
        r = st.session_state.single_res
        st.markdown("## Customer Survey Result")
        st.divider()
        st.markdown(f"**Customer:** {r['customer_name']}")
        st.markdown(f"**Lat/Long:** {r['cust_lat']}, {r['cust_lon']}")
        st.markdown(f"**Connected Node:** {r['connected_name']}")
        st.markdown(f"**Status:** {r['connected_status']}")
        st.markdown(f"**Reason:** {r['connected_reason']}")
        st.markdown(f"**Distance Result:** {r['connected_dist']}m")
        st_folium(draw_map(r, r["connected_map_obj"], r["recommended_map_obj"]), height=500, width=1000, key="ms")
