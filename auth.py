import hmac
import streamlit as st

def check_password() -> bool:
    """Returns `True` if the user had the correct password."""
    
    # Handle missing password gracefully
    if "APP_PASSWORD" not in st.secrets:
        st.error("Application password has not been configured.")
        st.stop()

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if hmac.compare_digest(st.session_state["password"], st.secrets["APP_PASSWORD"]):
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False

    # Return True if the password is validated.
    if st.session_state.get("password_correct", False):
        return True

    # Show input for password.
    st.title("SEO Toolkit")
    st.text_input(
        "Please enter the password to access the toolkit:", 
        type="password", 
        on_change=password_entered, 
        key="password"
    )
    if "password_correct" in st.session_state:
        st.error("😕 Password incorrect")
    return False

def render_logout_button():
    """Renders a logout button in the sidebar."""
    if st.sidebar.button("Logout"):
        st.session_state["password_correct"] = False
        st.rerun()
