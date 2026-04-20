import streamlit as st
import pandas as pd
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
import smtplib
from email.message import EmailMessage

# ==============================
# PAGE CONFIG
# ==============================
st.set_page_config(page_title="EPP HSE AI Dashboard", layout="wide")

# ==============================
# HEADER (SAFE LOGO)
# ==============================
col_logo, col_title = st.columns([1, 5])

with col_logo:
    try:
        st.image("logo.png", width=80)
    except:
        st.write("🏢")

with col_title:
    st.markdown("## EPP Company - HSE AI Dashboard")

st.markdown("---")

# ==============================
# PDF FUNCTION
# ==============================
def create_pdf(actions, kpis):
    doc = SimpleDocTemplate("HSE_Report.pdf")
    styles = getSampleStyleSheet()
    content = []

    content.append(Paragraph("HSE Report", styles['Title']))
    content.append(Spacer(1, 10))

    for k, v in kpis.items():
        content.append(Paragraph(f"{k}: {v}", styles['Normal']))

    content.append(Spacer(1, 10))

    for a in actions:
        content.append(Paragraph(a, styles['Normal']))

    doc.build(content)
    return "HSE_Report.pdf"

# ==============================
# EMAIL FUNCTION
# ==============================
def send_email(receiver_email, file_path):
    sender_email = "your_email@gmail.com"
    app_password = "your_app_password"

    msg = EmailMessage()
    msg["Subject"] = "HSE Report"
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.set_content("Attached is the HSE report.")

    with open(file_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf", filename="HSE_Report.pdf")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(msg)

# ==============================
# FILE UPLOAD
# ==============================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file, engine='openpyxl', header=1)

    # ==============================
    # CLEAN DATA
    # ==============================
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

    df = df[df["Date"].notna()]
    df = df[df["Date"] != "Annual Planned"]

    df["Date"] = pd.to_datetime(df["Date"], errors='coerce')
    df["Year"] = df["Date"].dt.year
    df["Month"] = df["Date"].dt.month_name()

    df = df.dropna(subset=["Year"])

    # ==============================
    # SIDEBAR FILTERS
    # ==============================
    st.sidebar.header("🔎 Filters")

    years = sorted(df["Year"].unique())
    selected_year = st.sidebar.multiselect("Year", years, default=years)

    months = df["Month"].unique()
    selected_month = st.sidebar.multiselect("Month", months, default=months)

    df = df[df["Year"].isin(selected_year)]
    df = df[df["Month"].isin(selected_month)]

    # ==============================
    # COLUMN FINDER
    # ==============================
    def find_col(keyword):
        for col in df.columns:
            if keyword.lower() in col.lower():
                return col
        return None

    col_hours = find_col("EPP Total")
    col_training = find_col("Training")
    col_nearmiss = find_col("Near Miss")
    col_fatal = find_col("Fatal")
    col_ptw = find_col("PTW")
    col_lwdc = find_col("LWDC")
    col_mtc = find_col("MTC")
    col_fac = find_col("FAC")

    # ==============================
    # KPIs
    # ==============================
    st.subheader("📊 Key Metrics")

    k1, k2, k3, k4 = st.columns(4)
    k5, k6, k7, k8 = st.columns(4)

    k1.metric("Worked Hours", int(df[col_hours].sum()) if col_hours else 0)
    k2.metric("Training", int(df[col_training].sum()) if col_training else 0)
    k3.metric("Near Miss", int(df[col_nearmiss].sum()) if col_nearmiss else 0)
    k4.metric("Fatal", int(df[col_fatal].sum()) if col_fatal else 0)

    k5.metric("PTW", int(df[col_ptw].sum()) if col_ptw else 0)
    k6.metric("LWDC", int(df[col_lwdc].sum()) if col_lwdc else 0)
    k7.metric("MTC", int(df[col_mtc].sum()) if col_mtc else 0)
    k8.metric("FAC", int(df[col_fac].sum()) if col_fac else 0)

    # ==============================
    # GROUP DATA
    # ==============================
    yearly = df.groupby("Year").sum(numeric_only=True).reset_index()

    # ==============================
    # CHARTS
    # ==============================
    st.subheader("📈 Trends")

    if col_hours:
        st.plotly_chart(px.area(yearly, x="Year", y=col_hours, title="Worked Hours"))

    if col_training:
        st.plotly_chart(px.line(yearly, x="Year", y=col_training, markers=True, title="Training"))

    if col_nearmiss:
        st.plotly_chart(px.line(yearly, x="Year", y=col_nearmiss, markers=True, title="Near Miss"))

    # ==============================
    # INCIDENTS
    # ==============================
    st.subheader("📊 Incident Analysis")

    c1, c2, c3 = st.columns(3)

    if col_lwdc:
        c1.plotly_chart(px.line(yearly, x="Year", y=col_lwdc, markers=True, title="LWDC"))

    if col_mtc:
        c2.plotly_chart(px.line(yearly, x="Year", y=col_mtc, markers=True, title="MTC"))

    if col_fac:
        c3.plotly_chart(px.line(yearly, x="Year", y=col_fac, markers=True, title="FAC"))

    # ==============================
    # ACTION ENGINE
    # ==============================
    st.subheader("🎯 Recommended Actions")

    actions = []

    if col_training and yearly[col_training].iloc[-1] < yearly[col_training].mean():
        actions.append("Increase training by 20%.")

    if col_lwdc and yearly[col_lwdc].sum() > 0:
        actions.append("Investigate LWDC root causes.")

    if col_fac and yearly[col_fac].sum() > 0:
        actions.append("Immediate management review for FAC.")

    if col_nearmiss:
        actions.append("Encourage near miss reporting.")

    if not actions:
        actions.append("Safety performance stable.")

    for a in actions:
        st.markdown(f"- {a}")

    # ==============================
    # EXPORT & EMAIL
    # ==============================
    st.subheader("📤 Export & Send")

    kpis = {
        "Worked Hours": int(df[col_hours].sum()) if col_hours else 0,
        "Training": int(df[col_training].sum()) if col_training else 0,
        "Near Miss": int(df[col_nearmiss].sum()) if col_nearmiss else 0,
        "Fatal": int(df[col_fatal].sum()) if col_fatal else 0,
    }

    if st.button("📄 Generate PDF"):
        create_pdf(actions, kpis)
        st.success("PDF created")

    email = st.text_input("Enter email")

    if st.button("📧 Send Email"):
        if email:
            send_email(email, "HSE_Report.pdf")
            st.success("Email sent")
        else:
            st.warning("Enter email first")

else:
    st.info("Upload Excel file")

# ==============================
# FOOTER
# ==============================
st.markdown("---")
st.markdown("© 2026 EPP Company | HSE AI Dashboard")
