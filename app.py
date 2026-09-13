import streamlit as st
import pandas as pd
import requests
from datetime import datetime

# ============================================================
# NUSTATYMAI
# ============================================================
N8N_BASE_URL = st.secrets.get("N8N_BASE_URL", "")
N8N_API_KEY = st.secrets.get("N8N_API_KEY", "")

ALLOWED_STATUSES = ["Laukia", "Patvirtinta"]
REQUEST_TIMEOUT = 20
# ============================================================

st.set_page_config(page_title="Sumų suvestinė", page_icon="📊", layout="wide")
st.title("📊 Sumų suvestinė — valdymo skydelis")

if not N8N_BASE_URL:
    st.error(
        "Nenustatytas N8N_BASE_URL. Programos nustatymuose (Settings → Secrets) "
        "įrašyk N8N_BASE_URL ir N8N_API_KEY."
    )
    st.stop()


def headers():
    return {"X-Api-Key": N8N_API_KEY}


@st.cache_data(ttl=15, show_spinner=False)
def fetch_data():
    """
    GET {N8N_BASE_URL}/webhook/gauti-duomenis
    Tikimasi n8n Webhook (GET), kuris perziurai.xlsx paverčia JSON ir
    grąžina per "Respond to Webhook" mazgą. Turinys — visų eilučių
    sąrašas (įskaitant "VISO:" eilutę), pvz.:
        [{"klientas": "...", "sutarties_nr": "...", "objektas": "...",
          "suma_be_pvm": 123.45, "suma_su_pvm": 149.37,
          "busena": "Laukia", "saltinio_failas": "...", ...}, ...]
    """
    resp = requests.get(
        f"{N8N_BASE_URL}/webhook/gauti-duomenis", headers=headers(), timeout=REQUEST_TIMEOUT
    )
    resp.raise_for_status()
    return pd.DataFrame(resp.json())


def post_status_updates(rows: list[dict]):
    """
    POST {N8N_BASE_URL}/webhook/atnaujinti-busena
    Body: {"updates": [{"saltinio_failas": "...", "busena": "Patvirtinta"}, ...]}
    n8n pusėje: pagal saltinio_failas surasti atitinkamą eilutę
    perziurai.xlsx faile ir perrašyti busena stulpelį.
    """
    resp = requests.post(
        f"{N8N_BASE_URL}/webhook/atnaujinti-busena",
        headers=headers(),
        json={"updates": rows},
        timeout=REQUEST_TIMEOUT,
    )
    resp.raise_for_status()


tab1, tab2, tab3, tab4 = st.tabs(
    ["🚀 Paleidimas", "📤 Įkelti failą", "✅ Patvirtinimas", "📈 Suvestinė"]
)

# ------------------------------------------------------------
# 1. PALEIDIMAS
# ------------------------------------------------------------
with tab1:
    st.subheader("Paleisti apdorojimą dabar")
    st.write(
        "Automatika ir taip veikia pati kas 10 minučių. "
        "Šis mygtukas leidžia paleisti apdorojimą iškart, nelaukiant."
    )
    if st.button("▶️ Apdoroti dabar", type="primary"):
        # POST {N8N_BASE_URL}/webhook/apdoroti-dabar
        # Webhook mazge "Respond" = Immediately, kad mygtukas nekabėtų
        # laukdamas, kol AI apdoros visus failus.
        with st.spinner("Siunčiama komanda į n8n..."):
            try:
                resp = requests.post(
                    f"{N8N_BASE_URL}/webhook/apdoroti-dabar",
                    headers=headers(),
                    timeout=REQUEST_TIMEOUT,
                )
                if resp.status_code in (200, 201):
                    st.success("Apdorojimas paleistas! Rezultatų ieškok skiltyje „Suvestinė“ po kelių minučių.")
                    fetch_data.clear()
                else:
                    st.error(f"n8n atsakė klaida (kodas {resp.status_code}). Patikrink Webhook mazgo nustatymus.")
            except requests.exceptions.RequestException as e:
                st.error(f"Nepavyko pasiekti n8n. Ar veikia tunelis ir n8n konteineris? Klaida: {e}")

# ------------------------------------------------------------
# 2. FAILŲ ĮKĖLIMAS
# ------------------------------------------------------------
with tab2:
    st.subheader("Įkelti naują dokumentą")
    st.write("Įkelk sutarties priedą, sąskaitą ar atliktų darbų aktą (PDF arba Word).")

    uploaded = st.file_uploader("Pasirink failą", type=["pdf", "docx", "doc"])
    if uploaded is not None:
        if st.button("⬆️ Įkelti į apdorojimo eilę"):
            # POST {N8N_BASE_URL}/webhook/ikelti-faila
            # Siunčiama kaip GRYNI dvejetainiai duomenys (raw body),
            # NE multipart/form-data — n8n Webhook mazge "Raw Body" = true,
            # "Field Name for Binary Data" = "data". Failo pavadinimas
            # perduodamas per X-Filename antraštę.
            with st.spinner("Keliama..."):
                try:
                    upload_headers = dict(headers())
                    upload_headers["Content-Type"] = "application/octet-stream"
                    upload_headers["X-Filename"] = uploaded.name
                    resp = requests.post(
                        f"{N8N_BASE_URL}/webhook/ikelti-faila",
                        headers=upload_headers,
                        data=uploaded.getvalue(),
                        timeout=60,
                    )
                    resp.raise_for_status()
                    result = resp.json()
                    status = result.get("status")
                    if status == "exists":
                        st.warning(f"„{uploaded.name}“ jau yra aplanke arba jau apdorotas anksčiau.")
                    elif status == "ok":
                        st.success(
                            f"„{uploaded.name}“ įkeltas. Bus apdorotas per artimiausius 10 min "
                            f"(arba paspausk „Apdoroti dabar“ skiltyje „Paleidimas“)."
                        )
                    else:
                        st.error(f"n8n grąžino netikėtą atsakymą: {result}")
                except requests.exceptions.RequestException as e:
                    st.error(f"Nepavyko įkelti failo. Klaida: {e}")

# ------------------------------------------------------------
# 3. PATVIRTINIMAS
# ------------------------------------------------------------
with tab3:
    st.subheader("Laukiančios patvirtinimo eilutės")

    try:
        df = fetch_data()
    except requests.exceptions.RequestException as e:
        df = None
        st.error(f"Nepavyko gauti duomenų iš n8n. Klaida: {e}")

    if df is not None:
        if df.empty:
            st.info("Duomenų dar nėra.")
        else:
            df_data = df[df["klientas"] != "VISO:"].copy()

            if "busena" not in df_data.columns:
                st.error("Stulpelio „busena“ duomenyse nėra — patikrink n8n endpoint'ą.")
            else:
                pending = df_data[df_data["busena"] == "Laukia"]

                if len(pending) == 0:
                    st.info("Nėra laukiančių patvirtinimo eilučių. ✅")
                else:
                    st.write(f"Iš viso laukia patvirtinimo: **{len(pending)}**")

                    edited = st.data_editor(
                        pending,
                        column_config={
                            "busena": st.column_config.SelectboxColumn(
                                "busena", options=ALLOWED_STATUSES, required=True
                            )
                        },
                        disabled=[c for c in pending.columns if c != "busena"],
                        use_container_width=True,
                        hide_index=True,
                        key="review_table",
                    )

                    if st.button("💾 Išsaugoti pakeitimus"):
                        changed = edited[edited["busena"] != pending["busena"]]
                        if changed.empty:
                            st.info("Nėra pakeitimų.")
                        else:
                            updates = [
                                {"saltinio_failas": row["saltinio_failas"], "busena": row["busena"]}
                                for _, row in changed.iterrows()
                            ]
                            try:
                                post_status_updates(updates)
                                fetch_data.clear()
                                st.success("Išsaugota!")
                                st.rerun()
                            except requests.exceptions.RequestException as e:
                                st.error(f"Nepavyko išsaugoti. Klaida: {e}")

# ------------------------------------------------------------
# 4. SUVESTINĖ
# ------------------------------------------------------------
with tab4:
    st.subheader("Būsenos suvestinė")

    try:
        df = fetch_data()
    except requests.exceptions.RequestException as e:
        df = None
        st.error(f"Nepavyko gauti duomenų iš n8n. Klaida: {e}")

    if df is not None and not df.empty:
        df_data = df[df["klientas"] != "VISO:"].copy()

        laukia = int((df_data["busena"] == "Laukia").sum()) if "busena" in df_data.columns else 0
        patvirtinta = int((df_data["busena"] == "Patvirtinta").sum()) if "busena" in df_data.columns else 0

        col1, col2, col3 = st.columns(3)
        col1.metric("Iš viso eilučių", len(df_data))
        col2.metric("Laukia", laukia)
        col3.metric("Patvirtinta", patvirtinta)

        viso_row = df[df["klientas"] == "VISO:"]
        if len(viso_row) > 0 and "suma_su_pvm" in viso_row.columns:
            st.metric("Bendra suma (su PVM)", f"{float(viso_row.iloc[0]['suma_su_pvm']):.2f} €")

        st.caption(f"Atnaujinta: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        if st.button("🔄 Atnaujinti"):
            fetch_data.clear()
            st.rerun()
    elif df is not None:
        st.warning("Dar nėra duomenų.")
