import streamlit as st

st.set_page_config(
    page_title="Presentation Builder",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def main():
    st.title("📊 Presentation Builder")
    st.markdown("""
    Welcome to the Presentation Builder! Use this tool to create professional presentations
    using data from various sources.
    
    ### How to use:
    1. Configure your API settings in the Config page
    2. Build your slides in the Slide Builder
    3. Preview and download your presentation
    """)
    
    st.markdown("### Navigation")
    st.page_link("pages/1_Slide_Builder.py", label="1️⃣ Slide Builder", icon="📝")
    st.page_link("pages/2_Config.py", label="⚙️ API Configuration", icon="⚙️")
    st.page_link("pages/3_Preview.py", label="👁️ Preview & Download", icon="👁️")

if __name__ == "__main__":
    main()
