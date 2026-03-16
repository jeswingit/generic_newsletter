#!/usr/bin/env python3
"""
streamlit_app.py

Streamlit web application for the newsletter generator.
Host this on Streamlit Cloud, Heroku, or any platform that supports Streamlit.
"""

import streamlit as st
from pathlib import Path
import tempfile
import os
from datetime import datetime

# Import functions from generate_newsletter.py
from generate_newsletter import (
    read_excel_rows,
    build_html_email,
    build_eml_message,
    _load_image_part,
    EMAIL_CONFIG,
    DEFAULT_OUT,
    DEFAULT_MONTH,
)

# Page configuration
st.set_page_config(
    page_title="ADIA EMEA Newsletter Generator",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1a1a1a;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #555555;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<div class="main-header">📧 ADIA EMEA Newsletter Generator</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Generate professional HTML email newsletters from Excel data</div>', unsafe_allow_html=True)
    
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Month selection
        months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]
        selected_month = st.selectbox(
            "Select Month:",
            months,
            index=months.index(DEFAULT_MONTH) if DEFAULT_MONTH in months else 2
        )
        
        # Email configuration
        st.subheader("Email Settings")
        from_email = st.text_input(
            "From Email:",
            value=EMAIL_CONFIG["from"],
            help="Sender email address"
        )
        to_email = st.text_input(
            "To Email:",
            value=EMAIL_CONFIG["to"],
            help="Recipient email address"
        )
        
        # Subject (optional)
        use_custom_subject = st.checkbox("Use custom subject", value=False)
        custom_subject = ""
        if use_custom_subject:
            custom_subject = st.text_input(
                "Subject:",
                value="",
                help="Leave empty to use default format"
            )
        else:
            current_year = datetime.now().year
            default_subject = EMAIL_CONFIG["subject"].format(
                month=selected_month, year=current_year
            )
            st.info(f"Default subject: **{default_subject}**")
        
        # Output filename
        output_filename = st.text_input(
            "Output Filename:",
            value=DEFAULT_OUT,
            help="Name for the generated EML file"
        )
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📊 Upload Excel File")
        
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload an Excel file with columns: Type, Data, Title, Creator, Image"
        )
        
        if uploaded_file is not None:
            # Display file info
            st.success(f"✅ File uploaded: **{uploaded_file.name}**")
            st.info(f"File size: {uploaded_file.size / 1024:.2f} KB")
            
            # Preview button
            if st.button("📋 Preview Excel Data", use_container_width=True):
                with st.spinner("Reading Excel file..."):
                    try:
                        # Save uploaded file temporarily
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                            tmp_file.write(uploaded_file.getvalue())
                            tmp_path = Path(tmp_file.name)
                        
                        # Read and display data
                        grouped = read_excel_rows(tmp_path)
                        
                        st.subheader("Excel Data Preview")
                        for type_key, rows in grouped.items():
                            with st.expander(f"**{type_key}** ({len(rows)} row(s))"):
                                for i, row in enumerate(rows, 1):
                                    st.write(f"**Row {i}:**")
                                    st.json(row)
                        
                        # Clean up temp file
                        os.unlink(tmp_path)
                        
                    except Exception as e:
                        st.error(f"Error reading Excel file: {str(e)}")
            
            # Generate button
            st.markdown("---")
            generate_button = st.button(
                "🚀 Generate Newsletter",
                type="primary",
                use_container_width=True,
                help="Generate the newsletter EML file"
            )
            
            if generate_button:
                generate_newsletter(
                    uploaded_file,
                    selected_month,
                    from_email,
                    to_email,
                    custom_subject if use_custom_subject else None,
                    output_filename
                )
    
    with col2:
        st.header("ℹ️ Instructions")
        st.markdown("""
        ### How to Use:
        
        1. **Upload Excel File**
           - File must have columns: Type, Data, Title, Creator, Image
        
        2. **Select Month**
           - Choose the month for the newsletter
        
        3. **Configure Email**
           - Set From/To addresses
           - Optionally customize subject
        
        4. **Generate**
           - Click "Generate Newsletter"
           - Download the EML file
        
        ### Content Types:
        - **Month News**: Bullet list items
        - **Save the Date**: Event announcements
        - **Product**: Product spotlights with images
        - **General**: Informational blocks
        
        ### Tips:
        - Images should be relative paths from Excel file location
        - Preview data before generating
        - Check the status messages for any issues
        """)
        
        st.markdown("---")
        st.header("📝 Status")
        if 'status' in st.session_state:
            st.info(st.session_state.status)


def generate_newsletter(uploaded_file, month, from_email, to_email, custom_subject, output_filename):
    """Generate the newsletter from uploaded Excel file."""
    
    status_container = st.container()
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    try:
        # Step 1: Save uploaded file temporarily
        status_text.info("📥 Saving uploaded file...")
        progress_bar.progress(10)
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = Path(tmp_file.name)
        
        # Step 2: Read Excel data
        status_text.info("📊 Reading Excel data...")
        progress_bar.progress(30)
        
        grouped = read_excel_rows(tmp_path)
        
        # Display data summary
        data_summary = "**Data Summary:**\n"
        for type_key, rows in grouped.items():
            data_summary += f"- {type_key}: {len(rows)} row(s)\n"
        status_text.info(data_summary)
        progress_bar.progress(40)
        
        # Step 3: Prepare image attachments
        status_text.info("🖼️ Preparing image attachments...")
        progress_bar.progress(50)
        
        image_cids: dict[str, str] = {}
        image_parts: dict[str, object] = {}
        
        product_rows = grouped.get("Product", [])
        image_count = 0
        for row in product_rows:
            if row.get("image") and row["image"] not in image_cids:
                img_part, cid = _load_image_part(row["image"], tmp_path.parent)
                if img_part is not None:
                    image_cids[row["image"]] = cid
                    image_parts[row["image"]] = img_part
                    image_count += 1
        
        if image_count > 0:
            status_text.info(f"✅ Prepared {image_count} image(s)")
        progress_bar.progress(60)
        
        # Step 4: Build HTML content
        status_text.info("🏗️ Building HTML email structure...")
        progress_bar.progress(70)
        
        html = build_html_email(grouped, month, EMAIL_CONFIG, image_cids)
        status_text.info(f"✅ Built HTML with {len(grouped)} section type(s)")
        progress_bar.progress(80)
        
        # Step 5: Build EML message
        status_text.info("📧 Building EML message...")
        progress_bar.progress(85)
        
        if custom_subject and custom_subject.strip():
            subject = custom_subject.strip()
        else:
            current_year = datetime.now().year
            subject = EMAIL_CONFIG["subject"].format(month=month, year=current_year)
        
        msg = build_eml_message(html, from_email, to_email, subject)
        
        # Attach images
        for image_path, img_part in image_parts.items():
            msg.attach(img_part)
        
        progress_bar.progress(90)
        
        # Step 6: Prepare download
        status_text.info("💾 Preparing download...")
        progress_bar.progress(95)
        
        eml_bytes = msg.as_bytes()
        
        # Clean up temp file
        os.unlink(tmp_path)
        
        progress_bar.progress(100)
        progress_bar.empty()
        
        # Success message and download button
        st.success("✅ Newsletter generated successfully!")
        st.markdown(f'<div class="success-box"><strong>Subject:</strong> {subject}</div>', unsafe_allow_html=True)
        
        st.download_button(
            label="📥 Download EML File",
            data=eml_bytes,
            file_name=output_filename,
            mime="message/rfc822",
            type="primary",
            use_container_width=True
        )
        
        status_text.empty()
        st.session_state.status = f"✅ Newsletter generated: {output_filename}"
        
    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Error generating newsletter: {str(e)}")
        st.exception(e)
        st.session_state.status = f"❌ Error: {str(e)}"


if __name__ == "__main__":
    main()
