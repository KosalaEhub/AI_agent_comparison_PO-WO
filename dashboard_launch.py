import streamlit as st
import os
import subprocess
import sys
import webbrowser
import threading
import importlib.util

# -----------------------------
# Folder path
base_path = os.path.dirname(os.path.abspath(__file__))

# Map buttons to python scripts
modules = {
    "Price Ticket": os.path.join(base_path, "CSAPP.py"),
    "Care Labels": os.path.join(base_path, "care_dashboard.py"),
    "Heat Transfer": os.path.join(base_path, "heat.py"),
    "RFID": os.path.join(base_path, "rfid.py"),
}

# -----------------------------
st.set_page_config(page_title="ITL", layout="wide")

# -----------------------------
# Tailwind-like card styling
st.markdown(
    """
    <style>
    body {
        background-color: #f3f4f6;
    }
    .module-card {
        border-radius: 1.5rem;
        padding: 2rem;
        text-align: center;
        color: white;
        font-family: 'Arial', sans-serif;
        transition: transform 0.3s, box-shadow 0.3s;
        cursor: pointer;
        box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        margin-bottom: 2rem;
    }
    .module-card:hover {
        transform: scale(1.05);
        box-shadow: 0 15px 30px rgba(0,0,0,0.2);
    }
    .module-title {
        font-size: 1.5rem;
        font-weight: bold;
        margin-bottom: 0.5rem;
    }
    .module-desc {
        font-size: 1rem;
        opacity: 0.9;
    }
    .header-title {
        font-size: 3rem;
        font-weight: 800;
        color: #4f46e5;
    }
    .header-subtitle {
        font-size: 1.2rem;
        color: #374151;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# -----------------------------
# Header
st.markdown(
    """
    <div style="text-align: center; margin-bottom: 3rem;">
        <h1 class="header-title">ITL PO & WO Checking System</h1>
        <p class="header-subtitle">Select a module to manage your apparel label printing</p>
    </div>
    """,
    unsafe_allow_html=True
)

# -----------------------------
col1, col2, col3, col4 = st.columns(4)

# -----------------------------
# Function to run external Streamlit script in new tab
def open_streamlit_script(script_path, port=8502):
    if not os.path.exists(script_path):
        st.error(f"{os.path.basename(script_path)} not found!")
        return

    def run_script():
        subprocess.run([sys.executable, "-m", "streamlit", "run", script_path, "--server.port", str(port)])

    threading.Thread(target=run_script, daemon=True).start()
    webbrowser.open(f"http://localhost:{port}")

# -----------------------------
# Module card function
def module_card(column, name, desc, port, bg_color):
    with column:
        st.markdown(
            f"""
            <div class="module-card" style="background-color: {bg_color};">
                <h2 class="module-title">{name}</h2>
                <p class="module-desc">{desc}</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        if st.button(f"Open {name}", key=f"btn_{name}"):
            # ALL modules now open in new tab
            open_streamlit_script(modules[name], port=port)

# -----------------------------
# Show module cards
module_card(col1, "Price Ticket", "Design and print price tickets for your products", 8503, "#6366f1")  # Indigo
module_card(col2, "Care Labels", "Create and manage care labels for apparel", 8502, "#10b981")        # Emerald Green
module_card(col3, "Heat Transfer", "Generate heat transfer labels quickly", 8504, "#f59e0b")         # Amber
module_card(col4, "RFID", "Manage RFID tags for your inventory", 8505, "#ef4444")                    # Red
