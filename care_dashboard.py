import streamlit as st
import os
import subprocess
import sys
import threading
import webbrowser
import time

# -----------------------------
# Configuration
st.set_page_config(page_title="Care Labels Dashboard - ITL", layout="wide")

# -----------------------------
# Care Label Data
CARE_LABELS = [
    "LB 6751", "PWLB-165 C/1", "LB 6735 Angel Pink", "PWLB-171 C/1", 
    "LB 6745", "LB 07200 C/1", "LB 07202 C/1", "LB 2691", "LB 5735", 
]

# Button colors
BUTTON_COLORS = ["#6366f1", "#10b981", "#f59e0b", "#ef4444"]

# -----------------------------
# Custom CSS
st.markdown("""
<style>
body { background-color: #f3f4f6; font-family: 'Arial', sans-serif;width:100%; }
.header-container { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 2rem; border-radius: 1.5rem; margin-bottom: 2rem; color: white; text-align: center; box-shadow: 0 10px 25px rgba(0,0,0,0.1); }
.header-title { font-size: 2.5rem; font-weight: 800; margin: 0; }
.header-subtitle { font-size: 1.2rem; opacity: 0.9; margin-top: 0.5rem; }
.label-card { background: white; border-radius: 1rem; padding: 1.5rem; text-align: center; color: white; font-family: 'Arial', sans-serif; transition: all 0.3s ease; cursor: pointer; box-shadow: 0 4px 15px rgba(0,0,0,0.1); margin-bottom: 1rem; border: 3px solid transparent; }
.label-card:hover { transform: translateY(-5px); box-shadow: 0 8px 25px rgba(0,0,0,0.2); border-color: rgba(255,255,255,0.3); }
.label-title { font-size: 1.1rem; font-weight: 700; margin: 0; text-shadow: 0 2px 4px rgba(0,0,0,0.3); }
.back-btn-container { position: fixed; top: 80px; left: 0px; margin-left: 20px; z-index: 1000; }
.back-button { background: linear-gradient(135deg, #4f46e5, #7c3aed); color: white; border: none; padding: 12px 20px; border-radius: 25px; font-weight: 600; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; gap: 8px; box-shadow: 0 4px 15px rgba(79,70,229,0.4); transition: all 0.3s ease; }
.back-button:hover { transform: translateY(-2px); color:white; box-shadow: 0 6px 20px rgba(79,70,229,0.6); background: linear-gradient(135deg, #6366f1, #8b5cf6); }
.search-container { background: white; padding: 1rem; border-radius: 1rem; box-shadow: 0 4px 15px rgba(0,0,0,0.1); margin-bottom: 2rem; display: flex; gap: 0.5rem; align-items: center; }
.stTextInput > div > div > input { border-radius: 25px; border: 2px solid #e5e7eb; padding: 10px 15px; font-size: 1rem; transition: all 0.3s ease; flex:1; }
.stTextInput > div > div > input:focus { border-color: #6366f1; box-shadow: 0 0 0 3px rgba(99,102,241,0.1); }
</style>
""", unsafe_allow_html=True)

# -----------------------------
# Back Button
st.markdown("""
<div class="back-btn-container">
    <a href="http://localhost:8501" class="back-button">← Back</a>
</div>
""", unsafe_allow_html=True)

st.markdown("<div style='padding-top: 60px;'></div>", unsafe_allow_html=True)

# -----------------------------
# Header
st.markdown("""
<div class="header-container">
    <h1 class="header-title">Care Labels Dashboard</h1>
    <p class="header-subtitle">Select and manage your care label printing</p>
</div>
""", unsafe_allow_html=True)

# -----------------------------
# Search Bar + Button
st.markdown('<div class="search-container">', unsafe_allow_html=True)
search_query = st.text_input("🔍 Search Care Labels", placeholder="Enter label name or code...")
search_button = st.button("Search")
st.markdown('</div>', unsafe_allow_html=True)

# -----------------------------
# Function to run `new.py` in background and open browser
def open_new_script():
    base_path = os.path.dirname(os.path.abspath(__file__))
    new_py_path = os.path.join(base_path, "new.py")

    if os.path.exists(new_py_path):
        def run_script():
            subprocess.run([sys.executable, "-m", "streamlit", "run", new_py_path, "--server.port", "8506"])
        threading.Thread(target=run_script, daemon=True).start()
        time.sleep(3)  # Wait for server to start
        webbrowser.open_new_tab("http://localhost:8506")
    else:
        st.error("new.py not found!")

# -----------------------------
# Filter labels based on search
filtered_labels = CARE_LABELS
if search_button:
    if search_query:
        filtered_labels = [label for label in CARE_LABELS if search_query.lower() in label.lower()]
        if not filtered_labels:
            st.warning("No labels found matching your search.")
    else:
        st.info("Please enter a search term to filter labels.")

# -----------------------------
# Display label cards
cols_per_row = 4
for i in range(0, len(filtered_labels), cols_per_row):
    cols = st.columns(cols_per_row)
    for j, col in enumerate(cols):
        if i + j < len(filtered_labels):
            label = filtered_labels[i + j]
            color = BUTTON_COLORS[(i + j) % len(BUTTON_COLORS)]
            with col:
                st.markdown(
                    f"""
                    <div class="label-card" style="background: linear-gradient(135deg, {color}, {color}dd);">
                        <h3 class="label-title">{label}</h3>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
                button_key = f"btn_{label}_{i}_{j}"
                if st.button(f"Select {label}", key=button_key):
                    if label in ["LB 5735", "LB 5736"]:
                        open_new_script()
                    else:
                        st.info(f"You selected {label}")

# -----------------------------
# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #6b7280; padding: 1rem;">
    <p>ITL Care Labels Management System | Select a label to start processing</p>
</div>
""", unsafe_allow_html=True)
