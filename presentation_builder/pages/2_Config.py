import streamlit as st
import os
from dotenv import load_dotenv

# Load existing .env file if it exists
load_dotenv()

def save_config():
    """Save configuration to .env file"""
    with open('.env', 'w') as f:
        f.write("# S&P Global Market Intelligence (CIQ) API Configuration\n")
        f.write(f"CIQ_USER={st.session_state.ciq_user}\n" if 'ciq_user' in st.session_state else "")
        f.write(f"CIQ_PASS={st.session_state.ciq_pass}\n" if 'ciq_pass' in st.session_state else "")
    st.success("Configuration saved successfully!")

def main():
    st.title("⚙️ API Configuration")
    
    st.markdown("""
    Configure your S&P Global Market Intelligence (CIQ) API credentials here. 
    These will be saved in the `.env` file in your project directory.
    """)
    
    # CIQ API Section
    st.subheader("S&P Global Market Intelligence (CIQ) API")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.text_input("CIQ Username", 
                     value=os.getenv('CIQ_USER', ''), 
                     key='ciq_user',
                     help="Your S&P Global Market Intelligence username")
    
    with col2:
        st.text_input("CIQ Password", 
                     value=os.getenv('CIQ_PASS', ''), 
                     type="password", 
                     key='ciq_pass',
                     help="Your S&P Global Market Intelligence password")
    
    # Save button
    if st.button("💾 Save Configuration", use_container_width=True):
        save_config()
    
    # Display current configuration
    with st.expander("Current Configuration", expanded=True):
        config = {
            "CIQ_USER": os.getenv('CIQ_USER', 'Not set'),
            "CIQ_PASS": "********" if os.getenv('CIQ_PASS') else "Not set"
        }
        
        st.json(config)
        
        if os.getenv('CIQ_USER') and os.getenv('CIQ_PASS'):
            st.success("✅ CIQ API credentials are configured")
        else:
            st.warning("⚠️ Please provide your CIQ API credentials")
    
    # Instructions
    st.markdown("""
    ### Configuration Instructions:
    1. Enter your S&P Global Market Intelligence (CIQ) username and password
    2. Click 'Save Configuration' to store your credentials
    3. The credentials will be saved in a `.env` file in your project directory
    
    ### Security Notes:
    - Never commit the `.env` file to version control
    - Keep your credentials private and secure
    - The password is stored in plaintext in the `.env` file
    - Restart the app after changing configuration
    
    ### Getting Help:
    - Contact S&P Global Market Intelligence support for API access
    - Refer to the documentation for more information
    """)

if __name__ == "__main__":
    main()
