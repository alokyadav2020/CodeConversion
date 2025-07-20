import streamlit as st
import os
from pathlib import Path
# from .streamlit.marcos_ import main_vba_code_converter


if "role" not in st.session_state:
    st.session_state.role = None

ROLES = [None,"Admin"]

PASSWORD = st.secrets["password"] 
USER = st.secrets["user"]
def login():

    # st.header("Log in")
    # role = st.selectbox("Choose your role", ROLES)
    with st.form("Login Form", clear_on_submit=True):
        st.write("Login Form")
        
        # Username and password inputs
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        
        # Submit button
        submit_button = st.form_submit_button("Login",use_container_width=True)
        
        # Check if the form was submitted
        if submit_button:

            if USER == username and PASSWORD == password:
                st.session_state.role = "Admin"
                st.success("Login successful!")
                st.session_state["logged_in"] = True
                st.rerun()
           


def logout():
    st.session_state.role = None
    st.rerun()


role = st.session_state.role


logout_page = st.Page(logout, title="Log out", icon=":material/logout:")



admin_1 = st.Page(os.path.join("streamlit","marcos_consersion.py"),title="marcos extractor",icon=":material/security:",default=(role == "Admin"),)
admin_2 = st.Page(os.path.join("streamlit","find_controls.py"),title="find_controls",icon=":material/security:",)
admin_3 = st.Page(os.path.join("streamlit","find_controls._1.py"),title="find_controls 01",icon=":material/security:",)
# admin_2 = st.Page(os.path.join("streamlit","Page_3_about_industry.py"), title="About Company",icon=":material/person_add:" )


account_pages = [logout_page]


admin_pages = [admin_1,admin_2, admin_3]



page_dict = {}

if st.session_state.role == "Admin":
    page_dict["Admin"] = admin_pages

if len(page_dict) > 0:
    pg = st.navigation({"Account": account_pages} | page_dict)
else:
    pg = st.navigation([st.Page(login)])

pg.run()