import threading, subprocess, time
import streamlit as st

def iniciar_cron():
    def cron_loop():
        while True:
            try:
                subprocess.call(["/bin/bash", "ping.sh"], timeout=30)
            except Exception as e:
                print("Error en cron:", e)
            time.sleep(600)
    if "cron_started" not in st.session_state:
        threading.Thread(target=cron_loop, daemon=True).start()
        st.session_state.cron_started = True