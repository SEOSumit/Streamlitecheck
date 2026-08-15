import hmac
import streamlit as st
import extra_streamlit_components as stx

@st.cache_resource(experimental_allow_widgets=True)
def get_manager():
    return stx.CookieManager()

def check_password() -> bool:
    """Returns `True` if the user had the correct password."""
    
    # Handle missing password gracefully
    if "APP_PASSWORD" not in st.secrets:
        st.error("Application password has not been configured.")
        st.stop()
        
    cookie_manager = get_manager()

    # Check cookie first
    if cookie_manager.get("auth_token") == "authenticated":
        return True

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if hmac.compare_digest(st.session_state["password"], st.secrets["APP_PASSWORD"]):
            st.session_state["password_correct"] = True
            cookie_manager.set("auth_token", "authenticated", max_age=30*24*60*60)
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
    cookie_manager = get_manager()
    if st.sidebar.button("Logout"):
        st.session_state["password_correct"] = False
        cookie_manager.delete("auth_token")
        st.rerun()
