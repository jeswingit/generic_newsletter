#!/usr/bin/env python3
"""
streamlit_app.py

Streamlit web application for the newsletter generator.
Renders output with newsletter_renderer (same pipeline as the Flask builder).

Host this on Streamlit Cloud, Heroku, or any platform that supports Streamlit.
"""

import streamlit as st
from pathlib import Path
import tempfile
import os
from datetime import datetime

from streamlit_sortables import sort_items

from generate_newsletter import (
    read_excel_rows,
    build_eml_message,
    _load_image_part,
    EMAIL_CONFIG,
    DEFAULT_OUT,
    DEFAULT_MONTH,
    excel_to_newsletter_config,
    embed_local_images_in_config,
)
from newsletter_renderer import render_newsletter
from template_generator import create_excel_template

# Page configuration
st.set_page_config(
    page_title="Newsletter Generator",
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
    .preview-container {
        border: 2px solid #dddddd;
        border-radius: 0.5rem;
        padding: 1rem;
        background-color: #ffffff;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<div class="main-header">📧 Newsletter Generator</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">Build newsletters from Excel using the same HTML engine as the Flask '
        "Newsletter Builder (<code>newsletter_renderer</code>)</div>",
        unsafe_allow_html=True,
    )
    
    # Download Template Button (at the top)
    col_template, _ = st.columns([1, 3])
    with col_template:
        if st.button("📥 Download Excel Template", use_container_width=True, help="Download a sample Excel template with example data"):
            try:
                template_bytes = create_excel_template()
                st.download_button(
                    label="⬇️ Click to Download Template",
                    data=template_bytes.getvalue(),
                    file_name="newsletter_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.success("✅ Template ready! Click the download button above.")
            except Exception as e:
                st.error(f"Error generating template: {str(e)}")
    
    st.markdown("---")
    
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

        org_name = st.text_input(
            "Organization (header/footer):",
            value="Aon",
            help="Shown in the email header and footer",
        )
        tagline = st.text_input(
            "Tagline:",
            value="Newsletter",
            help="Subtitle under the organization name in the header",
        )

        st.subheader("Layout")

        available_blocks = [
            "Month News",
            "Save the Date",
            "General Information",
            "General",
        ]

        default_enabled_blocks = [b for b in available_blocks if b != "Save the Date"] + ["Save the Date"]
        enabled_blocks = st.multiselect(
            "Enabled blocks:",
            options=available_blocks,
            default=default_enabled_blocks,
            help="Select which blocks to include in the newsletter. Drag to reorder below.",
        )

        # Keep a stable list for sortables to avoid runtime add/remove glitches
        if "layout_blocks" not in st.session_state:
            st.session_state.layout_blocks = enabled_blocks[:]
        # If enabled blocks changed, reset ordering to the enabled set in current order
        if set(st.session_state.layout_blocks) != set(enabled_blocks):
            st.session_state.layout_blocks = enabled_blocks[:]

        if st.session_state.layout_blocks:
            st.caption("Drag and drop to reorder blocks.")
            ordered_blocks = sort_items(
                st.session_state.layout_blocks,
                direction="vertical",
                key="newsletter_layout_sort",
            )
            # `sort_items` returns a new ordered list; persist it
            st.session_state.layout_blocks = ordered_blocks
        else:
            ordered_blocks = []

        st.subheader("Block Background Colors")
        default_bg = {
            "Month News": EMAIL_CONFIG["colors"].get("white", "#ffffff"),
            "Save the Date": EMAIL_CONFIG["colors"].get("save_date_bg", "#E5EFF0"),
            "General Information": EMAIL_CONFIG["colors"].get("white", "#ffffff"),
            "General": EMAIL_CONFIG["colors"].get("white", "#ffffff"),
        }
        block_bg_colors: dict[str, str] = {}
        for block_id in available_blocks:
            if block_id not in enabled_blocks:
                continue
            block_bg_colors[block_id] = st.color_picker(
                f"{block_id} background",
                value=default_bg.get(block_id, "#ffffff"),
                key=f"bg_{block_id}",
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
            
            # Generate and Preview buttons
            st.markdown("---")
            col_gen, col_prev = st.columns(2)
            
            with col_gen:
                generate_button = st.button(
                    "🚀 Generate Newsletter",
                    type="primary",
                    use_container_width=True,
                    help="Generate the newsletter EML file"
                )
            
            with col_prev:
                preview_button = st.button(
                    "👁️ Preview Email",
                    use_container_width=True,
                    help="Preview how the newsletter will look"
                )
            
            if preview_button:
                preview_email(
                    uploaded_file,
                    selected_month,
                    from_email,
                    to_email,
                    custom_subject if use_custom_subject else None,
                    ordered_blocks,
                    block_bg_colors,
                    org_name,
                    tagline,
                )

            if generate_button:
                generate_newsletter(
                    uploaded_file,
                    selected_month,
                    from_email,
                    to_email,
                    custom_subject if use_custom_subject else None,
                    output_filename,
                    ordered_blocks,
                    block_bg_colors,
                    org_name,
                    tagline,
                )
    
    with col2:
        st.header("ℹ️ Instructions")
        st.markdown("""
        ### How to Use:
        
        1. **Download Template** (optional) — sample Excel with the right columns  
        2. **Upload Excel** — columns: Type, Data, Title, Creator, Image  
        3. **Month** — used in the default subject line  
        4. **Email & branding** — From/To, optional custom subject, organization & tagline  
        5. **Layout** — enable and reorder blocks; set per-block background colors  
        6. **Preview / Generate** — same HTML as `python app.py` / `newsletter_renderer`  
        
        ### Content Types (Type column):
        - **Month News** → bullet list (“What’s Going On”)  
        - **Save the Date** → event list (headings styled like other sections)  
        - **General Information** → **Product** rows as product cards (images via Image column paths)  
        - **General** → text blocks with title + body  
        
        ### Tips:
        - Image paths are relative to the Excel file folder (or absolute)  
        - Preview and HTML download embed images as data URIs; EML uses inline CID attachments  
        """)
        
        st.markdown("---")
        st.header("📝 Status")
        if 'status' in st.session_state:
            st.info(st.session_state.status)


def _build_meta_and_subject(
    month: str,
    from_email: str,
    to_email: str,
    custom_subject: str | None,
    org_name: str,
    tagline: str,
) -> tuple[dict, str]:
    if custom_subject and custom_subject.strip():
        subject = custom_subject.strip()
    else:
        current_year = datetime.now().year
        subject = EMAIL_CONFIG["subject"].format(month=month, year=current_year)
    meta = {
        "newsletterName": "Newsletter",
        "subject": subject,
        "from": from_email,
        "to": to_email,
        "orgName": org_name,
        "tagline": tagline,
    }
    return meta, subject


def generate_newsletter(
    uploaded_file,
    month,
    from_email,
    to_email,
    custom_subject,
    output_filename,
    ordered_blocks: list[str],
    block_bg_colors: dict[str, str],
    org_name: str,
    tagline: str,
):
    """Generate the newsletter from uploaded Excel file."""
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        status_text.info("📥 Saving uploaded file...")
        progress_bar.progress(10)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_path = Path(tmp_file.name)

        status_text.info("📊 Reading Excel data...")
        progress_bar.progress(25)

        grouped = read_excel_rows(tmp_path)

        data_summary = "**Data Summary:**\n"
        for type_key, rows in grouped.items():
            data_summary += f"- {type_key}: {len(rows)} row(s)\n"
        status_text.info(data_summary)
        progress_bar.progress(35)

        meta, subject = _build_meta_and_subject(
            month, from_email, to_email, custom_subject, org_name, tagline
        )
        config = excel_to_newsletter_config(
            grouped,
            email_config=EMAIL_CONFIG,
            ordered_blocks=ordered_blocks if ordered_blocks else None,
            block_bg_colors=block_bg_colors,
            meta=meta,
            bullet_heading="What's Going On",
        )

        status_text.info("🖼️ Preparing image attachments...")
        progress_bar.progress(45)

        image_cids: dict[str, str] = {}
        image_parts: dict[str, object] = {}
        image_count = 0
        for row in grouped.get("Product", []):
            img = row.get("image")
            if img and img not in image_cids:
                img_part, cid = _load_image_part(img, tmp_path.parent)
                if img_part is not None:
                    image_cids[img] = cid
                    image_parts[img] = img_part
                    image_count += 1

        if image_count > 0:
            status_text.info(f"✅ Prepared {image_count} image(s) for EML")
        progress_bar.progress(55)

        status_text.info("🏗️ Rendering HTML (newsletter_renderer)...")
        progress_bar.progress(65)

        html_eml = render_newsletter(config, image_cids=image_cids)
        html_embedded = render_newsletter(
            embed_local_images_in_config(config, tmp_path.parent)
        )
        status_text.info("✅ HTML ready")
        progress_bar.progress(75)

        status_text.info("📧 Building EML message...")
        progress_bar.progress(80)

        msg = build_eml_message(html_eml, from_email, to_email, subject)
        for img_part in image_parts.values():
            msg.attach(img_part)

        progress_bar.progress(90)
        status_text.info("💾 Preparing download...")
        progress_bar.progress(95)

        eml_bytes = msg.as_bytes()
        os.unlink(tmp_path)

        progress_bar.progress(100)
        progress_bar.empty()

        st.success("✅ Newsletter generated successfully!")
        st.markdown(
            f'<div class="success-box"><strong>Subject:</strong> {subject}</div>',
            unsafe_allow_html=True,
        )

        col_dl_eml, col_dl_html = st.columns(2)

        with col_dl_eml:
            st.download_button(
                label="📥 Download EML File",
                data=eml_bytes,
                file_name=output_filename,
                mime="message/rfc822",
                type="primary",
                use_container_width=True,
                key="download_eml",
            )

        with col_dl_html:
            st.download_button(
                label="📥 Download HTML File",
                data=html_embedded,
                file_name=output_filename.replace(".eml", ".html"),
                mime="text/html",
                use_container_width=True,
                key="download_html",
            )

        status_text.empty()
        st.session_state.status = f"✅ Newsletter generated: {output_filename}"

    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Error generating newsletter: {str(e)}")
        st.exception(e)
        st.session_state.status = f"❌ Error: {str(e)}"


def preview_email(
    uploaded_file,
    month,
    from_email,
    to_email,
    custom_subject,
    ordered_blocks: list[str],
    block_bg_colors: dict[str, str],
    org_name: str,
    tagline: str,
):
    """Preview the newsletter HTML before generating (embedded images for iframe)."""
    try:
        with st.spinner("Generating preview..."):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = Path(tmp_file.name)

            grouped = read_excel_rows(tmp_path)

            meta, subject = _build_meta_and_subject(
                month, from_email, to_email, custom_subject, org_name, tagline
            )
            config = excel_to_newsletter_config(
                grouped,
                email_config=EMAIL_CONFIG,
                ordered_blocks=ordered_blocks if ordered_blocks else None,
                block_bg_colors=block_bg_colors,
                meta=meta,
                bullet_heading="What's Going On",
            )

            html = render_newsletter(
                embed_local_images_in_config(config, tmp_path.parent)
            )

            os.unlink(tmp_path)

            st.markdown("---")
            st.subheader("📧 Email Preview")
            st.info(f"**Subject:** {subject}")
            st.info(f"**From:** {from_email} | **To:** {to_email}")

            st.markdown('<div class="preview-container">', unsafe_allow_html=True)
            st.components.v1.html(html, height=800, scrolling=True)
            st.markdown("</div>", unsafe_allow_html=True)

            st.session_state.preview_html = html
            st.session_state.preview_subject = subject

            st.download_button(
                label="📥 Download HTML Version",
                data=html,
                file_name="newsletter_preview.html",
                mime="text/html",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"❌ Error generating preview: {str(e)}")
        st.exception(e)


if __name__ == "__main__":
    main()
