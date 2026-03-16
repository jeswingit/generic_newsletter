# Newsletter Generator

A Python script that generates newsletter EML files from Excel data files.

## Description

This tool reads an Excel file with newsletter content and generates a professionally formatted HTML email newsletter in EML format. It supports multiple content types including month news, save-the-date events, product spotlights, and general information sections.

## Features

- Reads newsletter content from Excel files (columns: Type, Data, Title, Creator, Image)
- Generates HTML email newsletters with embedded images
- Supports multiple content sections:
  - **Month News**: Bullet list of monthly updates
  - **Save the Date**: Event announcements
  - **Product**: Product spotlight cards with images
  - **General**: Informational blocks

## Requirements

- Python 3.x
- openpyxl library
- streamlit (for web interface)

## Installation

Install all dependencies:
```bash
pip install -r requirements.txt
```

Or install individually:
```bash
pip install openpyxl streamlit
```

## Usage

### 🌐 Web Interface (Streamlit) - Recommended

**Run locally:**
```bash
streamlit run streamlit_app.py
```

**Host on Streamlit Cloud:**
1. Push your code to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Sign in with GitHub
4. Click "New app"
5. Select your repository and set main file to `streamlit_app.py`
6. Click "Deploy"

The web interface provides:
- Drag-and-drop Excel file upload
- Month selection dropdown
- Email configuration sidebar
- Real-time preview of Excel data
- One-click newsletter generation
- Direct EML file download

### 🖥️ Desktop GUI (Tkinter)

Launch the desktop GUI application:
```bash
python newsletter_gui.py
```

The GUI provides an easy-to-use interface for:
- Selecting Excel files via file browser
- Choosing the month from a dropdown menu
- Configuring email addresses and subject
- Selecting output file location
- Viewing real-time generation status

### Command Line Interface

Basic usage:
```bash
python generate_newsletter.py
```

With custom options:
```bash
python generate_newsletter.py --xlsx data.xlsx --out output.eml --month "April"
```

### Command-line Arguments

- `--xlsx`: Path to Excel data file (default: `EmailData (2).xlsx`)
- `--out`: Output EML file path (default: `newsletter_output.eml`)
- `--month`: Month name for the newsletter (default: `March`)
- `--from`: From email address
- `--to`: To email address
- `--subject`: Email subject line

## Excel File Format

The Excel file should contain the following columns:
- **Type**: Content type (`Month News`, `Save the Date`, `Product`, or `General`)
- **Data**: Main content text
- **Title**: Title/heading for the content
- **Creator**: Author/creator name (optional)
- **Image**: Path to image file (optional, for Product type)

## Deployment Options

### Streamlit Cloud (Free)
- **URL**: [share.streamlit.io](https://share.streamlit.io)
- **Pros**: Free, easy setup, automatic HTTPS, no server management
- **Limitations**: Public apps are free, private apps require paid plan

### Other Hosting Options
- **Heroku**: Use Procfile with `web: streamlit run streamlit_app.py --server.port=$PORT`
- **AWS/Azure/GCP**: Deploy as containerized app
- **Local Network**: Run `streamlit run streamlit_app.py --server.address=0.0.0.0`

## License

This project is provided as-is for internal use.
