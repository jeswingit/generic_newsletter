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

## Installation

```bash
pip install openpyxl
```

## Usage

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

## License

This project is provided as-is for internal use.
