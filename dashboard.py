# -*- coding: utf-8 -*-
"""Unified Protobazar sales analytics and pricing dashboard."""
from __future__ import annotations
from pathlib import Path
import streamlit as st
from sales_analytics_module import show_sales_page
from pricing_module import show_pricing_page, LOGO_B64

st.set_page_config(page_title="پلتفرم پروتوبازار", page_icon=str(Path(__file__).resolve().parent / "logo.png"), layout="wide", initial_sidebar_state="expanded")

GLOBAL_CSS = r"""
:root{--navy:#111c3a;--navy2:#202e55;--brand:#4f46e5;--cyan:#06b6d4;--ink:#17243b;--muted:#64748b;--line:#e2e8f0}
html,body,[class*="css"],.stApp,button,input,textarea,select{font-family:Tahoma,'Segoe UI',sans-serif!important}
.stApp{background:radial-gradient(ellipse at 10% 0%,#4f46e50b,transparent 30%),linear-gradient(180deg,#f8fafc,#f4f7fb);color:var(--ink)}
.block-container{max-width:1540px;padding-top:1.4rem!important;padding-bottom:3.5rem!important}
h1,h2,h3,h4{color:var(--navy)!important;font-weight:800!important;direction:rtl}
.stMarkdown,.stCaption,.stAlert,[data-testid="stWidgetLabel"],label{direction:rtl;text-align:right}
section[data-testid="stSidebar"]{background:linear-gradient(165deg,#101a36 0%,#1d2a50 55%,#111c3a 100%);border-left:1px solid #ffffff14;box-shadow:-12px 0 36px #0f172a16}
section[data-testid="stSidebar"] *{font-family:Tahoma,'Segoe UI',sans-serif!important;color:#edf3ff}
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p{color:#c2cde1}
section[data-testid="stSidebar"] h1,section[data-testid="stSidebar"] h2,section[data-testid="stSidebar"] h3{color:#fff!important}

/* Keep values legible inside light sidebar form controls. */
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea,
section[data-testid="stSidebar"] [data-baseweb="input"] input,
section[data-testid="stSidebar"] [data-testid="stNumberInput"] input {
    color:#172033!important;
    -webkit-text-fill-color:#172033!important;
    background-color:#fff!important;
    caret-color:#172033!important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    color:#172033!important;
    background-color:#fff!important;
}
section[data-testid="stSidebar"] [data-baseweb="select"] input {
    color:#172033!important;
    -webkit-text-fill-color:#172033!important;
}
section[data-testid="stSidebar"] [data-testid="stNumberInput"] button {
    color:#172033!important;
    background-color:#f8fafc!important;
}
.sidebar-brand-card{display:flex;align-items:center;gap:12px;padding:13px;margin:7px 0 16px;border-radius:16px;background:linear-gradient(130deg,#ffffff16,#ffffff08);border:1px solid #ffffff25;box-shadow:0 8px 22px #050b1c28;direction:rtl}
.sidebar-brand-card img{width:48px;height:48px;object-fit:contain;background:#fff;border-radius:12px;padding:4px;flex:0 0 auto}
.sidebar-brand-card strong{display:block;color:#fff;font-size:1rem}.sidebar-brand-card small{display:block;color:#c5d1e5;font-size:.75rem;margin-top:3px}
.sidebar-nav-title{font-size:.75rem;color:#aebbd3;letter-spacing:.04em;margin:5px 4px 8px;font-weight:700}
section[data-testid="stSidebar"] .stButton>button{width:100%;display:flex;justify-content:flex-start;text-align:right;direction:rtl;white-space:normal;line-height:1.5;padding:10px 14px;border-radius:11px;border:1px solid #ffffff20;background:#ffffff0a;color:#edf3ff!important;font-weight:700;min-height:44px;transition:background .16s ease,border-color .16s ease}
section[data-testid="stSidebar"] .stButton>button:hover{background:#ffffff18;border-color:#8ea5d2;transform:translateX(-2px)}
section[data-testid="stSidebar"] .stButton>button[kind="primary"]{background:linear-gradient(110deg,#4f46e5,#3730a3);border-color:#7770f0;box-shadow:0 5px 15px #4338ca50}
.stButton>button,[data-testid="stDownloadButton"] button{border-radius:11px!important;background:var(--brand)!important;border:1px solid var(--brand)!important;color:#fff!important;font-weight:700!important;transition:all .16s ease}
.stButton>button:hover,[data-testid="stDownloadButton"] button:hover{background:#312e81!important;transform:translateY(-1px)}
[data-testid="stFileUploader"]{border:1px dashed #cbd5e1;border-radius:14px;background:#fff;padding:8px;box-shadow:0 4px 15px #17243b09}
[data-testid="stFileUploaderDropzone"]{background:#f8fafc!important;border:0!important;border-radius:10px!important}
[data-testid="stFileUploaderDropzone"] *{font-family:Tahoma,'Segoe UI',sans-serif!important}
[data-testid="stFileUploaderDropzone"] button{background:#eef2ff!important;color:#312e81!important;border:1px solid #c7d2fe!important;border-radius:9px!important;font-size:inherit!important}
[data-testid="stMetric"],div[data-testid="stDataFrame"],div[data-testid="stPlotlyChart"],[data-testid="stExpander"]{border-radius:15px!important}
[data-testid="stMetric"]{background:#fff;border:1px solid var(--line);padding:15px 17px;box-shadow:0 6px 20px #17243b0b}
.stTabs [aria-selected="true"]{color:var(--brand)!important;border-bottom-color:var(--cyan)!important}
[data-testid="stSidebarCollapseButton"],[data-testid="stSidebarCollapsedControl"],[data-testid="collapsedControl"]{background:#18264a!important;border:1px solid #ffffff38!important;border-radius:10px!important;color:transparent!important;font-size:0!important;text-indent:-9999px!important;overflow:hidden!important;width:2.5rem!important;height:2.5rem!important;position:relative!important}
[data-testid="stSidebarCollapseButton"]::before,[data-testid="stSidebarCollapsedControl"]::before,[data-testid="collapsedControl"]::before{content:"‹";font:700 24px Arial,sans-serif!important;color:#fff!important;display:block;text-indent:0!important;text-align:center;line-height:1.1}
.home-hero{direction:rtl;display:flex;align-items:center;gap:24px;padding:30px;background:linear-gradient(120deg,#111c3a,#283d70 65%,#176878);border-radius:22px;color:#fff;box-shadow:0 15px 40px #14213d24;margin:3px 0 22px}
.home-hero img{height:88px;width:88px;object-fit:contain;background:#fff;border-radius:18px;padding:7px;flex:none}
.home-hero h1{color:#fff!important;margin:0 0 8px;font-size:1.7rem}.home-hero p{color:#e3eaf7;margin:0;line-height:1.9}
.home-card{height:100%;min-height:190px;padding:22px;border-radius:18px;background:#fff;border:1px solid #e2e8f0;box-shadow:0 8px 24px #14213d0b;direction:rtl;text-align:right}
.home-card .icon{font-size:1.7rem}.home-card h3{margin:12px 0 7px;font-size:1.1rem}.home-card p{color:#64748b;font-size:.91rem;line-height:1.8;margin:0}
@media(max-width:700px){.home-hero{padding:20px;gap:14px}.home-hero img{width:62px;height:62px}.home-hero h1{font-size:1.2rem}}
"""
st.markdown(f"<style>{GLOBAL_CSS}</style>", unsafe_allow_html=True)

if "dashboard_page" not in st.session_state:
    st.session_state.dashboard_page = "home"

page_labels = {"home": "خانه", "sales": "تحلیل فروش", "pricing": "سیستم قیمت‌گذاری"}
with st.sidebar:
    st.markdown(f'<div class="sidebar-brand-card"><img src="data:image/webp;base64,{LOGO_B64}" alt="Protobazar"><div><strong>پروتو بازار</strong><small>سامانه هوشمند تحلیل و قیمت‌گذاری</small></div></div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-nav-title">ناوبری</div>', unsafe_allow_html=True)
    current = st.session_state.dashboard_page
    for key, label in page_labels.items():
        # Plain Persian labels avoid icon-font fallback artifacts in RTL layouts.
        if st.button(label, key=f"nav_{key}", type="primary" if current == key else "secondary", use_container_width=True):
            st.session_state.dashboard_page = key
            st.rerun()
    if current != "home":
        st.divider()
        if st.button("←  بازگشت به خانه", key="back_home", use_container_width=True):
            st.session_state.dashboard_page = "home"
            st.rerun()

if st.session_state.dashboard_page == "sales":
    show_sales_page()
elif st.session_state.dashboard_page == "pricing":
    show_pricing_page()
else:
    st.markdown(f'<section class="home-hero"><img src="data:image/webp;base64,{LOGO_B64}" alt="Protobazar"><div><h1>سامانه هوشمند تحلیل و قیمت‌گذاری پروتو بازار</h1><p>ابزارهای یکپارچه برای پایش عملکرد فروش، تحلیل گزارش‌ها و مدیریت هوشمند قیمت محصولات.</p></div></section>', unsafe_allow_html=True)
    st.markdown("### از کجا شروع کنیم؟")
    c1,c2=st.columns(2)
    with c1:
        st.markdown('<div class="home-card"><div class="icon">📊</div><h3>تحلیل فروش</h3><p>گزارش‌های فروش را بارگذاری کنید و شاخص‌ها، روندها و عملکرد فروش را بررسی کنید.</p></div>', unsafe_allow_html=True)
        if st.button("رفتن به تحلیل فروش  ←",key="home_sales",use_container_width=True):
            st.session_state.dashboard_page="sales"; st.rerun()
    with c2:
        st.markdown('<div class="home-card"><div class="icon">🏷️</div><h3>سیستم قیمت‌گذاری</h3><p>قیمت‌های محصولات و سابقه خرید را تحلیل کنید، تغییرات گروهی انجام دهید و خروجی اکسل بگیرید.</p></div>', unsafe_allow_html=True)
        if st.button("رفتن به سیستم قیمت‌گذاری  ←",key="home_pricing",use_container_width=True):
            st.session_state.dashboard_page="pricing"; st.rerun()
