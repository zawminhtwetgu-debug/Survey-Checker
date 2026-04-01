import math, folium, requests, time, io
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium
from fpdf import FPDF

st.set_page_config(layout="wide", page_title="FTTH Deployment Survey Tool")

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

def clean_for_pdf(text):
    if not isinstance(text, str):
        text = str(text)
    replacements = {"\u2013": "-", "\u2014": "-", "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"'}
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.encode("latin-1", "replace").decode("latin-1")

def haversine(lat1, lon1, lat2, lon2):
    if pd.isna(lat1) or pd.isna(lon1) or pd.isna(lat2) or pd.isna(lon2): return 999999
    r = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi, dlambda = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return r * (2 * math.atan2(math.sqrt(a), math.sqrt(1 - a)))

@st.cache_data(ttl=3600)
def get_shortest_path(lat1, lon1, lat2, lon2):
    direct_dist = haversine(lat1, lon1, lat2, lon2)
    url = f"https://router.project-osrm.org/route/v1/foot/{lon1},{lat1};{lon2},{lat2}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=5)
        data = r.json()
        if data.get("code") == "Ok" and data.get("routes"):
            route = data["routes"][0]
            path_dist = route["distance"]
            if path_dist > (direct_dist * 2.5):
                return direct_dist, {"type": "LineString", "coordinates": [[lon1, lat1], [lon2, lat2]]}
            return path_dist, route["geometry"]
    except:
        pass
    return direct_dist, {"type": "LineString", "coordinates": [[lon1, lat1], [lon2, lat2]]}

# =========================================================
# CORE LOGIC
# =========================================================
def analyze_one_customer(nodes_df, cust_name, cust_lat, cust_lon, connected_name, partner="-"):
    res = {
        "customer_name": cust_name, "cust_lat": cust_lat, "cust_lon": cust_lon,
        "partner": partner, "connected_name": connected_name or "-",
        "connected_status": "NOK", "connected_reason": "-",
        "connected_dist": "-", "connected_map_obj": None, "recommended_map_obj": None
    }
    if pd.isna(cust_lat) or pd.isna(cust_lon):
        res["connected_reason"] = "Location Error"
        return res

    search_name = str(connected_name).strip().upper() if connected_name and str(connected_name).strip() != "-" else None

    if search_name:
        node = nodes_df[nodes_df["node_name_upper"] == search_name]
        if not node.empty:
            n = node.iloc[0]
            dist, geom = get_shortest_path(cust_lat, cust_lon, n["Latitude"], n["Longitude"])
            is_dist_ok = dist <= MAX_DISTANCE
            is_port_ok = n["act"] < MAX_CAPACITY
            
            # Status and Reason mapping based on requested image
            if is_dist_ok and is_port_ok:
                stat, reas = "OK", "Can Deploy"
            elif not is_dist_ok:
                stat, reas = "Over Meter", f"Over meter {int(dist)} m"
            else:
                stat, reas = "Full Port", f"{MAX_CAPACITY} customers"

            res.update({"connected_status": stat, "connected_reason": reas, "connected_dist": int(dist)})
            res["connected_map_obj"] = {"name": n["node_name"], "lat": n["Latitude"], "lon": n["Longitude"], "dist": dist, "geom": geom, "status": stat, "reason": reas}
        else:
            res.update({"connected_status": "NOK", "connected_reason": "Node Not Found"})

    if res["connected_reason"] != "Can Deploy":
        temp_nodes = nodes_df.copy()
        temp_nodes["temp_dist"] = temp_nodes.apply(lambda r: haversine(cust_lat, cust_lon, r["Latitude"], r["Longitude"]), axis=1)
        candidates = temp_nodes[temp_nodes["temp_dist"] < 600].sort_values("temp_dist").head(TOP_FALLBACK_NODES)
        for _, n in candidates.iterrows():
            if search_name and n["node_name_upper"] == search_name: continue
            dist, geom = get_shortest_path(cust_lat, cust_lon, n["Latitude"], n["Longitude"])
            if dist <= MAX_DISTANCE and n["act"] < MAX_CAPACITY:
                res["recommended_map_obj"] = {"name": n["node_name"], "lat": n["Latitude"], "lon": n["Longitude"], "dist": dist, "geom": geom, "status": "OK", "reason": "Can Deploy"}
                break
    return res

def draw_map(res, conn=None, reco=None):
    if pd.isna(res["cust_lat"]) or pd.isna(res["cust_lon"]): return folium.Map(location=[16.8, 96.1], zoom_start=12)
    m = folium.Map(location=[res["cust_lat"], res["cust_lon"]], zoom_start=18)
    folium.Marker([res["cust_lat"], res["cust_lon"]], tooltip=res['customer_name'], icon=folium.Icon(color="blue")).add_to(m)
    for obj, color in [(conn, "red"), (reco, "green")]:
        if obj and obj.get("lat"):
            folium.Marker([obj["lat"], obj["lon"]], tooltip=f"Node: {obj['name']}", icon=folium.Icon(color=color)).add_to(m)
            if obj.get("geom"):
                folium.GeoJson(obj["geom"], style_function=lambda x, c=color: {"color": c, "weight": 5}).add_to(m)
    return m

# =========================================================
# UI
# =========================================================
st.title("FTTH Deployment Survey Tool")
st.markdown("##### Powered by Zaw Min Htwe")
st.info("Shortest Path Logic: လမ်းကြောင်းအတိုင်းတွက်ချက်ပြီး အဝေးကြီးပတ်မသွားစေရန် Detour Detection ထည့်သွင်းထားပါသည်။")

t1, t2 = st.tabs(["Batch Check", "Single Check"])
with t1:
    st.caption("Fixed Data: nodes.csv, NIMS.xlsx, new_customers.xlsx")
    col_b1, col_b2 = st.columns(2)
    if col_b1.button("Run Batch", type="primary"):
        try:
            # Changed filename to NIMS
            node_data, cust_data = pd.read_csv("nodes.csv"), pd.read_excel("NIMS.xlsx")
            act_counts = cust_data.groupby("node_name").size().reset_index(name="act")
            nodes = node_data.merge(act_counts, on="node_name", how="left").fillna(0)
            nodes["node_name_upper"] = nodes["node_name"].astype(str).str.strip().str.upper()
            new_custs = pd.read_excel("new_customers.xlsx").dropna(subset=["customer_name"])
            results, summary_rows = [], []
            prog = st.progress(0)
            for i, (_, r) in enumerate(new_custs.iterrows()):
                res = analyze_one_customer(nodes, r["customer_name"], clean_num(r["lat"]), clean_num(r["Long"]), r.get("connected_node"), str(r.get("Partner", "-")))
                results.append(res); summary_rows.append({"Customer Name": res["customer_name"], "Partner": res["partner"], "Lat": res["cust_lat"], "Long": res["cust_lon"], "Status": res["connected_status"], "Reason": res["connected_reason"], "Connected Node": res["connected_name"], "Distance (m)": res["connected_dist"], "Recommended": res["recommended_map_obj"]["name"] if res["recommended_map_obj"] else "-"})
                prog.progress((i + 1) / len(new_custs))
            st.session_state.batch_results, st.session_state.batch_summary_df, st.session_state.batch_done = results, pd.DataFrame(summary_rows), True
        except Exception as e: st.error(f"Error: {e}")

    if col_b2.button("Clear Batch"): st.session_state.batch_done = False; st.rerun()

    if st.session_state.batch_done:
        st.markdown("### Filters")
        f_col1, f_col2, f_col3 = st.columns(3)
        df_full = st.session_state.batch_summary_df.copy()
        with f_col1: sel_p = st.multiselect("Partner", options=sorted(df_full["Partner"].unique()))
        with f_col2: sel_s = st.multiselect("Status", options=sorted(df_full["Status"].unique()))
        with f_col3: sel_r = st.multiselect("Reason", options=sorted(df_full["Reason"].unique()))
        f_df = df_full.copy()
        if sel_p: f_df = f_df[f_df["Partner"].isin(sel_p)]
        if sel_s: f_df = f_df[f_df["Status"].isin(sel_s)]
        if sel_r: f_df = f_df[f_df["Reason"].isin(sel_r)]
        st.dataframe(f_df, use_container_width=True)
        
        # Export Excel with Fixed Width
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            f_df.to_excel(writer, index=False, sheet_name='Survey_Results')
            workbook = writer.book
            worksheet = writer.sheets['Survey_Results']
            for i, col in enumerate(f_df.columns):
                worksheet.set_column(i, i, 20)
        st.download_button(label="Export to Excel", data=output.getvalue(), file_name="Survey_Report.xlsx", mime="application/vnd.ms-excel")
        
        st.divider()
        c_opts = f_df["Customer Name"].tolist()
        if c_opts:
            sel = st.selectbox("Select Customer to View Detail", options=c_opts, label_visibility="collapsed")
            res = next(r for r in st.session_state.batch_results if str(r["customer_name"]) == str(sel))
            # Detail display formatted like requested images
            st.markdown("## Customer Survey Result")
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
    with sl1: s_na, s_la = st.text_input("Customer Name"), st.text_input("Latitude")
    with sl2: s_no, s_lo = st.text_input("Connected Node"), st.text_input("Longitude")
    if st.button("Run Single", type="primary"):
        if s_la and s_lo:
            n_s = pd.read_csv("nodes.csv")
            n_s["node_name_upper"] = n_s["node_name"].astype(str).str.strip().str.upper()
            n_s["act"] = 0 
            st.session_state.single_res = analyze_one_customer(n_s, s_na, clean_num(s_la), clean_num(s_lo), s_no)
    if st.session_state.single_res:
        r = st.session_state.single_res
        st.markdown(f"**Status:** {r['connected_status']} | **Reason:** {r['connected_reason']} | **Distance:** {r['connected_dist']}m")
        st_folium(draw_map(r, r["connected_map_obj"], r["recommended_map_obj"]), height=500, width=1000)
