import io
import os
import secrets
import hashlib
import hmac
from collections import Counter

import pandas as pd
import qrcode
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

from storage import ExcelStore

st.set_page_config(page_title="ZEN ART RIOT | College Fest", page_icon="🎉", layout="wide", initial_sidebar_state="collapsed")

EVENTS = ["Singing", "Reel Roulette", "Film Making", "Dance", "Antakshari", "Mr Sau Ms Sau", "Star Night", "Fashion Show"]
DEPARTMENTS = ["Computer Science", "Information Technology", "Mechanical", "Civil", "Electrical", "Electronics", "Commerce", "Arts", "Science", "Other"]


def config(key, default=""):
    try:
        return st.secrets.get(key, os.environ.get(key, default))
    except (FileNotFoundError, KeyError):
        return os.environ.get(key, default)


@st.cache_resource
def store():
    return ExcelStore(config("AZURE_STORAGE_CONNECTION_STRING"), config("AZURE_CONTAINER", "zen-event"), config("AZURE_BLOB_NAME", "college_fest_registrations.xlsx"))


st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
:root{--ink:#11113d;--coral:#ff735b;--cream:#fbf7f2;--muted:#696a83}
html,body,.stApp,[data-testid="stAppViewContainer"]{background:#fbf7f2!important;color:var(--ink);font-family:'DM Sans',sans-serif}
.block-container{max-width:1500px;padding:1rem 2.5rem 4rem}
[data-testid="stHeader"]{background:transparent}[data-testid="stDecoration"]{display:none}
h1,h2,h3{font-family:'Plus Jakarta Sans',sans-serif;color:var(--ink)}
.zen-nav{display:flex;align-items:center;justify-content:space-between;gap:15px;padding:5px 8px 18px}
.zen-logo{font-family:'Plus Jakarta Sans',sans-serif;font-size:40px;font-weight:800;letter-spacing:-3px;line-height:1;color:var(--ink)}
.zen-logo span{color:var(--coral)}.zen-logo small{display:block;letter-spacing:5px;font-size:11px;font-weight:700;margin-top:7px}
.nav-links{display:flex;align-items:center;gap:32px;padding:17px 28px;border-radius:50px;background:#ffffffb8;box-shadow:0 10px 35px #2820400a;font-size:15px;font-weight:700}
.nav-links a{color:var(--ink)!important;text-decoration:none}.nav-links a:hover{color:var(--coral)!important}
.nav-pill{border:2px solid #14143f;border-radius:30px;padding:12px 22px;font-weight:800;font-size:12px}
.hero-wrap{background:#fffaf5;display:grid;grid-template-columns:44% 56%;min-height:550px;overflow:hidden;border-radius:18px;position:relative}
.hero-copy{padding:55px 0 38px 34px;position:relative;z-index:2}
.eyebrow{letter-spacing:5px;font-size:13px;font-weight:800;color:#535676}
.hero-title{font-family:'Plus Jakarta Sans',sans-serif;font-size:clamp(3.4rem,6.0vw,6.7rem);line-height:.98;letter-spacing:-.075em;font-weight:800;margin:26px 0 16px;color:var(--ink)}
.hero-title em{color:var(--coral);font-style:normal}
.hero-desc{font-size:19px;color:#686b83;line-height:1.6;max-width:520px}
.hero-stats{display:flex;gap:24px;flex-wrap:wrap;margin:35px 0 32px}
.hero-stats>div{padding-right:20px;border-right:1px solid #cfc9ca}.hero-stats>div:last-child{border:0}
.hero-stats strong{display:block;font-family:'Plus Jakarta Sans',sans-serif;font-size:29px;color:var(--ink)}.hero-stats span{font-size:12px;color:#727386}
.hero-cta{display:inline-block;background:var(--coral);color:white!important;text-decoration:none;border-radius:50px;padding:17px 29px;font-size:16px;font-weight:800;box-shadow:0 12px 24px #ff735b24}
.hero-art{background-size:contain;background-repeat:no-repeat;background-position:center center;min-height:550px}
.quote-strip{text-align:center;font-size:15px;font-weight:800;color:#4d4b69;padding:14px 10px 24px}
.zen-section{display:flex;justify-content:space-between;align-items:end;gap:10px;margin:20px 0}
.section-kicker{font-size:12px;letter-spacing:3px;font-weight:800;color:var(--coral)}
.section-title{font-family:'Plus Jakarta Sans',sans-serif;font-size:35px;font-weight:800;letter-spacing:-.055em;color:var(--ink);margin:7px 0}
.section-sub{font-size:14px;color:#737389}
[data-testid="stButton"] button[kind="secondary"]{background:#ffc9bd!important;border:1px solid #f0c7c0!important;border-radius:20px!important;min-height:128px!important;justify-content:flex-start!important;text-align:left!important;white-space:pre-wrap!important;font-family:'Plus Jakarta Sans',sans-serif!important;font-size:17px!important;font-weight:800!important;color:#10103b!important;box-shadow:0 10px 22px #11113d10!important;transition:transform .2s,box-shadow .2s!important;padding:20px!important}
[data-testid="stButton"] button[kind="secondary"]:hover{transform:translateY(-4px);box-shadow:0 18px 32px #11113d1c!important}
[data-testid="stHorizontalBlock"]:has([data-testid="stButton"]) [data-testid="stColumn"]:nth-child(4n+1) button[kind="secondary"]{background:#ffd0c4!important}
[data-testid="stHorizontalBlock"]:has([data-testid="stButton"]) [data-testid="stColumn"]:nth-child(4n+2) button[kind="secondary"]{background:#b8edf2!important}
[data-testid="stHorizontalBlock"]:has([data-testid="stButton"]) [data-testid="stColumn"]:nth-child(4n+3) button[kind="secondary"]{background:#ffe09a!important}
[data-testid="stHorizontalBlock"]:has([data-testid="stButton"]) [data-testid="stColumn"]:nth-child(4n+4) button[kind="secondary"]{background:#c8e8d7!important}
[data-testid="stTabs"] [data-baseweb="tab-list"]{gap:20px;border-bottom:1px solid #e8dfd9}
[data-testid="stTabs"] button{font-family:'DM Sans',sans-serif;font-weight:800}
[data-testid="stForm"]{background:white;border:1px solid #eee5df;border-radius:18px;padding:22px}
[data-testid="stTextInput"] input{background:#fff8f1!important;border-radius:12px!important}
[data-testid="stFormSubmitButton"] button{background:#ff735b!important;color:white!important;border:0!important;border-radius:12px!important;min-height:49px;font-weight:800!important}
.footer{text-align:center;padding:36px 0 12px;color:#7a7a8a;font-size:13px}
@media(max-width:1000px){.hero-wrap{grid-template-columns:1fr}.hero-copy{padding:30px 24px}.hero-art{min-height:380px}.nav-links{display:none}.block-container{padding:1rem 1rem 3rem}}
@media(max-width:540px){.hero-title{font-size:3.4rem}.hero-copy{padding:24px 16px}.hero-art{min-height:245px}.hero-stats{gap:12px}.hero-stats>div{padding-right:12px}.section-title{font-size:28px}.nav-pill{font-size:10px;padding:10px}.zen-logo{font-size:31px}[data-testid="stButton"] button[kind="secondary"]{min-height:105px!important;font-size:13px!important;padding:12px!important}}

/* Festival collage design refresh */
.zen-nav{padding:10px 10px 22px}.nav-pill{background:#fffdfa;box-shadow:0 6px 20px #11113d0b}
.hero-wrap{grid-template-columns:42% 58%;background:#fffaf6;min-height:525px;box-shadow:0 14px 55px #24173a0a}
.hero-copy{padding:57px 0 36px 36px}
.hero-title{font-size:clamp(3.2rem,5.5vw,6rem);letter-spacing:-.08em;margin:20px 0 17px}
.hero-art{min-height:525px;background-position:center center;background-size:cover;mask-image:linear-gradient(to right,transparent 0%,black 10%);-webkit-mask-image:linear-gradient(to right,transparent 0%,black 10%)}
.hero-cta:hover{background:#f45f49}
.zen-section{margin-top:29px}
[data-testid="stButton"] button[kind="secondary"]{min-height:145px!important;border:0!important;border-radius:24px!important;font-size:19px!important;line-height:1.55!important}
[data-testid="stButton"] button[kind="secondary"]:focus{outline:2px solid #14143f}
@media(max-width:1000px){.hero-wrap{grid-template-columns:1fr}.hero-art{min-height:340px;background-size:contain;mask-image:none;-webkit-mask-image:none}.hero-copy{padding:30px 25px}}
@media(max-width:540px){.hero-art{min-height:220px}.hero-title{font-size:3.2rem}.hero-copy{padding:25px 18px}.hero-stats strong{font-size:23px}[data-testid="stButton"] button[kind="secondary"]{min-height:105px!important;font-size:13px!important}}
</style>""",unsafe_allow_html=True)

import base64
from pathlib import Path
_art = base64.b64encode((Path(__file__).parent / 'assets' / 'festival_collage.jpg').read_bytes()).decode('ascii')
st.markdown(f"""<div class="zen-nav"><div class="zen-logo">ZEN<span>.</span><small>ART RIOT</small></div>
<div class="nav-links"><a href="#home">Home</a><a href="#events">Events</a><a href="#registration">Register</a></div>
<div class="nav-pill">✦ COLLEGE FEST 2026</div></div>
<div id="home" class="hero-wrap"><div class="hero-copy"><div class="eyebrow">COLLEGE FEST 2026</div>
<div class="hero-title">BREAK THE<br><em>BORING.</em></div>
<div class="hero-desc">A collision of culture, creativity, competition and pure campus energy.</div>
<div class="hero-stats"><div><strong>08</strong><span>Epic events</span></div><div><strong>∞</strong><span>Memories</span></div><div><strong>01</strong><span>Campus celebration</span></div></div>
<a class="hero-cta" href="#events">Explore the Events →</a></div>
<div class="hero-art" style="background-image:url('data:image/jpeg;base64,{_art}')"></div></div>
<div class="quote-strip">✦ Same Campus. Different Stories. One Art Riot. ✦ &nbsp;&nbsp; GOOD VIBES ONLY</div>
<div id="events" class="zen-section"><div><div class="section-kicker">THE LINEUP / 08 EXPERIENCES</div><div class="section-title">Our Live Events</div><div class="section-sub">Music. Art. Stories. People. You. — Select an event to register.</div></div></div>""",unsafe_allow_html=True)

EVENT_ART = [("🎤", "Singing"),("📸", "Reel Roulette"),("🎬", "Film Making"),("💃", "Dance"),("🎶", "Antakshari"),("👑", "Mr Sau Ms Sau"),("🪩", "Star Night"),("👗", "Fashion Show")]
EVENT_TAGLINES = {"Singing":"Voices that unite", "Reel Roulette":"Spin. Shoot. Shine.","Film Making":"Stories beyond frames", "Dance":"The Fake Wedding", "Antakshari":"Songs. Teams. Fun.","Mr Sau Ms Sau":"Personality with purpose", "Star Night":"Bollywood Night", "Fashion Show":"Style beyond limits"}
for start_idx in (0,4):
    cols=st.columns(4,gap="medium")
    for col,(emoji,event) in zip(cols,EVENT_ART[start_idx:start_idx+4]):
        with col:
            if st.button(f"{emoji}  {event} ↗\n{EVENT_TAGLINES[event]}",key=f"event_{event}",use_container_width=True):
                st.session_state["selected_event"]=event
                st.session_state.pop("last_registration",None)
                st.rerun()

def qr_png(url):
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_H, box_size=12, border=4)
    qr.add_data(url)
    qr.make(fit=True)
    buf = io.BytesIO()
    qr.make_image(fill_color="#20204b", back_color="white").convert("RGB").save(buf, format="PNG")
    return buf.getvalue()


def poster_png(url):
    image = Image.new("RGB", (900, 1300), "#fff4de")
    draw = ImageDraw.Draw(image)
    try:
        bold = ImageFont.truetype("DejaVuSans-Bold.ttf", 74)
        mid = ImageFont.truetype("DejaVuSans-Bold.ttf", 38)
        small = ImageFont.truetype("DejaVuSans.ttf", 24)
    except OSError:
        bold = mid = small = ImageFont.load_default()
    navy = "#20204b"
    draw.rounded_rectangle((80, 90, 820, 270), radius=30, fill="#ffb92e", outline=navy, width=9)
    for x in range(103, 805, 32):
        draw.ellipse((x, 103, x+11, 114), fill="white")
        draw.ellipse((x, 246, x+11, 257), fill="white")
    for y in range(135, 245, 32):
        draw.ellipse((95, y, 106, y+11), fill="white")
        draw.ellipse((795, y, 806, y+11), fill="white")
    draw.text((450, 177), "ZEN ART RIOT", font=bold, anchor="mm", fill=navy)
    draw.text((450, 340), "Scan to register", font=mid, anchor="mm", fill="#d11965")
    draw.rounded_rectangle((140, 390, 760, 1020), radius=26, fill="white", outline=navy, width=6)
    qr = Image.open(io.BytesIO(qr_png(url))).convert("RGB").resize((560, 560))
    image.paste(qr, (170, 425))
    draw.text((450, 1085), "Singing  •  Reel Roulette  •  Film Making  •  Dance", font=small, anchor="mm", fill=navy)
    draw.text((450, 1130), "Antakshari  •  Mr Sau Ms Sau  •  Star Night", font=small, anchor="mm", fill=navy)
    draw.text((450, 1175), "Fashion Show", font=small, anchor="mm", fill=navy)
    draw.text((450, 1240), url[:70], font=small, anchor="mm", fill="#656580")
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return buf.getvalue()


def password_ok(password):
    expected = config("ADMIN_PASSWORD")
    if not expected:
        return False
    return hmac.compare_digest(hashlib.sha256(password.encode()).digest(), hashlib.sha256(str(expected).encode()).digest())


EVENT_TITLES = {
    "Singing": "🎤 SINGING", "Reel Roulette": "📱 REEL ROULETTE",
    "Film Making": "🎬 FILM MAKING", "Dance": "💃 DANCE – THE FAKE WEDDING",
    "Antakshari": "🎶 ANTAKSHARI", "Mr Sau Ms Sau": "👑 MR SAU MS SAU",
    "Star Night": "✨ STAR NIGHT – BOLLYWOOD NIGHT",
    "Fashion Show": "👗 FASHION SHOW",
}

st.markdown("""<style>
[role="dialog"]{border:1px solid #eae6e3!important;border-radius:24px!important;overflow:auto;max-height:90vh;box-shadow:0 24px 80px #10103225!important}
[role="dialog"] [data-testid="stForm"]{border:0!important;box-shadow:none!important;padding:5px 0!important}
[role="dialog"] [data-testid="stTextInput"] input{background:#fbf9f5!important;border:1px solid #e4e0da!important;border-radius:11px!important;min-height:47px}
[role="dialog"] [data-testid="stFormSubmitButton"] button{background:#ff735d!important;color:white!important}
.zen-modal-title{background:linear-gradient(115deg,#ffd2c5,#fff1de);padding:22px;border-radius:17px;font-family:'Plus Jakarta Sans',sans-serif;font-size:23px;font-weight:800;letter-spacing:-.05em;color:#101032}
</style>""",unsafe_allow_html=True)

@st.dialog("Event registration", width="small")
def event_registration_dialog(event):
    st.markdown(f'<div class="zen-modal-title">{EVENT_TITLES[event]}</div>', unsafe_allow_html=True)
    with st.form(f"modal_form_{event}", clear_on_submit=True):
        category = ""
        if event in ("Singing", "Dance"):
            category = st.radio("Category *", ["Single", "Duet", "Mix"], horizontal=True, key=f"cat_{event}")
        name = st.text_input("Name *", max_chars=100, placeholder="Enter your full name")
        batch = st.text_input("Batch *", max_chars=80, placeholder="e.g. 2024 or BBA 2nd year")
        mobile = st.text_input("Mobile number *", max_chars=10, placeholder="10-digit number")
        consent = st.checkbox("I agree to the organizers storing my registration details.", value=False)
        submit = st.form_submit_button("Register", use_container_width=True)
    if st.button("Cancel", key=f"cancel_{event}", use_container_width=True):
        st.session_state.pop("selected_event", None)
        st.rerun()
    if submit:
        if not consent:
            st.error("Please accept the consent checkbox.")
        else:
            try:
                registration_id, error = store().register({
                    "name": name, "batch": batch, "mobile": mobile,
                    "event": event, "category": category,
                    "department": "", "year": "", "college": "",
                })
                if error:
                    st.error(error)
                else:
                    st.session_state["last_registration"] = (event, registration_id)
                    st.session_state.pop("selected_event", None)
                    st.rerun()
            except Exception as exc:
                import logging
                logging.exception("ZEN EVENT save failure")
                st.error("Registration could not be saved. Please contact the organizer.")
                # Diagnostic details are shown locally; avoid exposing cloud credentials or internals to students.
                if not config("AZURE_STORAGE_CONNECTION_STRING"):
                    st.caption(f"Local Excel error: {type(exc).__name__}: {exc}")

register_tab, admin_tab = st.tabs(["Student Registration", "Admin Dashboard"])
with register_tab:
    st.markdown('<div id="registration" class="zen-section"><div><div class="section-kicker">READY TO JOIN?</div><div class="section-title">Your moment starts here.</div><p class="section-sub">Tap any event card above, complete the registration form and get your registration ID.</p></div></div>',unsafe_allow_html=True)
    if st.session_state.get("last_registration"):
        registered_event, reg_id = st.session_state["last_registration"]
        st.success(f"Registered for {registered_event}! Your registration ID is {reg_id}.")
        if st.button("Close success message"):
            st.session_state.pop("last_registration", None)
            st.rerun()

with admin_tab:
    st.subheader("Organizer Dashboard")
    if not config("ADMIN_PASSWORD"):
        st.warning("Admin access is disabled until ADMIN_PASSWORD is configured in Streamlit secrets.")
    elif not st.session_state.get("admin_logged_in", False):
        with st.form("admin_login"):
            password = st.text_input("Admin password", type="password")
            login = st.form_submit_button("Log in")
        if login:
            if password_ok(password):
                st.session_state["admin_logged_in"] = True
                st.rerun()
            else:
                st.error("Incorrect password.")
    else:
        if st.button("Log out"):
            st.session_state["admin_logged_in"] = False
            st.rerun()
        try:
            data = store().rows()
            df = pd.DataFrame(data, columns=["Registration ID", "Student Name", "Mobile Number", "Department", "Year", "College", "Event", "Registration Date", "Registration Time", "Category", "Batch"])
            unique_students = df["Mobile Number"].astype(str).nunique() if not df.empty else 0
            a, b, c = st.columns(3)
            a.metric("Unique mobile numbers", unique_students)
            b.metric("Total registrations", len(df))
            c.metric("Events", len(EVENTS))
            selected = st.selectbox("Filter by event", ["All Events"] + EVENTS)
            view = df if selected == "All Events" else df[df["Event"] == selected]
            search = st.text_input("Search by name or mobile")
            if search:
                view = view[view["Student Name"].astype(str).str.contains(search, case=False, regex=False) | view["Mobile Number"].astype(str).str.contains(search, regex=False)]
            st.dataframe(view, use_container_width=True, hide_index=True)
            if not df.empty:
                st.bar_chart(df["Event"].value_counts())
            st.download_button("Download Excel workbook", data=store().read_bytes(), file_name="college_fest_registrations.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            url = str(config("REGISTRATION_URL")).strip()
            if url.startswith("https://"):
                st.subheader("QR Code for Students")
                st.image(qr_png(url), width=260)
                st.download_button("Download QR code", qr_png(url), "zen_event_qr.png", "image/png")
                st.download_button("Download printable poster", poster_png(url), "zen_event_poster.png", "image/png")
            else:
                st.info("Set REGISTRATION_URL to your deployed https:// Streamlit URL to generate a QR code and poster.")
        except Exception:
            st.error("Unable to read the Excel workbook. Check the storage configuration.")

if st.session_state.get("selected_event") in EVENTS:
    event_registration_dialog(st.session_state["selected_event"])

st.markdown('<div class="footer">zen. &nbsp; · &nbsp; Made for unforgettable college memories</div>', unsafe_allow_html=True)
