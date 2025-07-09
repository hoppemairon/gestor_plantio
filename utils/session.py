import streamlit as st
import json
import os

CONFIG_PATH = "config.json"

DEFAULTS = {
    "pess_receita": 15,
    "pess_despesas": 10,
    "otm_receita": 10,
    "otm_despesas": 10,
    **{f"inf_{i}": 4.0 for i in range(5)}
}

def carregar_configuracoes():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r") as f:
            config = json.load(f)
    else:
        config = DEFAULTS.copy()

    for key, value in config.items():
        if key not in st.session_state:
            st.session_state[key] = value
