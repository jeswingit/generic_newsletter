"""
newsletter_renderer.py

Generic HTML email renderer. Takes a newsletter JSON config dict and produces
a complete, table-based HTML email string compatible with major email clients.

Each section renderer returns one or more <tr> elements (no outer wrapper).
render_newsletter() assembles the outer HTML shell.
"""

import html as html_module
import re
import uuid

# Matches [link text](https://url) markdown-style links
_LINK_RE = re.compile(r'\[([^\]]+)\]\((https?://[^)]+|mailto:[^)]+)\)')


def _render_text(text: str) -> str:
    """Escape text and convert [text](url) markdown links to HTML <a> tags."""
    parts = []
    last = 0
    for m in _LINK_RE.finditer(text):
        parts.append(html_module.escape(text[last:m.start()]))
        link_text = html_module.escape(m.group(1))
        url = m.group(2)
        parts.append(
            f'<a href="{url}" style="color:#EB0017;text-decoration:underline;"'
            f' target="_blank" rel="noopener noreferrer">{link_text}</a>'
        )
        last = m.end()
    parts.append(html_module.escape(text[last:]))
    return "".join(parts)
from datetime import datetime


# ---------------------------------------------------------------------------
# Section renderers — each returns HTML <tr> string(s)
# ---------------------------------------------------------------------------

def render_header(props: dict, theme: dict, image_cids: dict = None, **_) -> str:
    org_name = html_module.escape(props.get("orgName", "Aon"))
    tagline = html_module.escape(props.get("tagline", "Newsletter"))
    bg = props.get("backgroundColor") or theme.get("primaryColor", "#E31837")
    text_color = props.get("textColor", "#ffffff")
    logo_url = props.get("logoUrl", "")
    title_bg = _lighten_bg(bg)

    # Resolve logo: use CID for EML, direct URL for preview/HTML
    logo_html = ""
    if logo_url:
        if image_cids and logo_url in image_cids:
            logo_src = f"cid:{image_cids[logo_url]}"
        else:
            logo_src = logo_url
        logo_html = (
            f"<img src=\"{logo_src}\" alt=\"{org_name}\" "
            "style=\"max-height:48px;max-width:160px;height:auto;display:block;"
            "margin-bottom:8px;\" />"
        )

    return (
        "        <!-- HEADER -->\n"
        "        <tr>\n"
        "          <td style=\"padding:0;background-color:#ffffff;\">\n"
        "            <table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n"
        "              <tr>\n"
        f"                <td width=\"6\" style=\"background-color:{bg};\">&nbsp;</td>\n"
        "                <td valign=\"middle\" style=\"padding:18px 16px 18px 14px;"
        f"background-color:{title_bg};\">\n"
        + (f"                  {logo_html}\n" if logo_html else "")
        + f"                  <div style=\"font-size:22px;font-weight:bold;color:#1a1a1a;"
        "letter-spacing:0.5px;line-height:1.2;\">"
        f"{org_name}</div>\n"
        f"                  <div style=\"font-size:13px;color:#555555;margin-top:4px;"
        f"font-style:italic;\">{tagline}</div>\n"
        "                </td>\n"
        f"                <td style=\"background-color:{bg};\">&nbsp;</td>\n"
        "              </tr>\n"
        "            </table>\n"
        "          </td>\n"
        "        </tr>\n"
        "        <tr><td style=\"height:1px;background-color:#E0E0E0;"
        "font-size:0;line-height:0;\">&nbsp;</td></tr>"
    )


def render_footer(props: dict, theme: dict, **_) -> str:
    org_name = html_module.escape(props.get("orgName", "Aon"))
    year = html_module.escape(str(props.get("year", datetime.now().year)))
    bg = props.get("backgroundColor", "#1A1A1A")
    text_color = props.get("textColor", "#aaaaaa")
    link_color = theme.get("primaryColor", "#E31837")

    return (
        "\n"
        "        <!-- FOOTER -->\n"
        "        <tr>\n"
        f"          <td style=\"padding:16px 28px;background-color:{bg};text-align:center;\">\n"
        f"            <div style=\"font-size:11px;color:{text_color};line-height:1.5;\">"
        f"&copy; {year} {org_name} &bull; All rights reserved</div>\n"
        f"            <div style=\"font-size:11px;color:{text_color};margin-top:4px;\">\n"
        f"              <a href=\"#\" style=\"color:{link_color};text-decoration:none;\">Unsubscribe</a>\n"
        "              &nbsp;&bull;&nbsp;\n"
        f"              <a href=\"#\" style=\"color:{text_color};text-decoration:none;\">View in browser</a>\n"
        "            </div>\n"
        "          </td>\n"
        "        </tr>"
    )


def render_bullet_list(props: dict, theme: dict, **_) -> str:
    heading = html_module.escape(props.get("heading", ""))
    items = props.get("items", [])
    bullet_color = props.get("bulletColor") or theme.get("primaryColor", "#E31837")
    bg = props.get("backgroundColor", "#ffffff")
    bg_style = f"background-color:{bg};" if bg else ""
    divider_color = "#eeeeee"

    rows_html = "\n".join(
        "              <tr>\n"
        f"                <td width=\"12\" valign=\"top\" style=\"padding:0 0 12px 0;color:{bullet_color};"
        "font-size:18px;line-height:1;\">&bull;</td>\n"
        "                <td style=\"padding:0 0 12px 8px;font-size:13px;font-weight:bold;"
        "color:#1a1a1a;line-height:1.5;\">"
        f"{_render_text(str(item))}</td>\n"
        "              </tr>"
        for item in items if str(item).strip()
    )

    if not rows_html:
        return ""

    return (
        "\n"
        f"        <!-- BULLET LIST: {html_module.escape(heading)} -->\n"
        "        <tr>\n"
        f"          <td style=\"padding:24px 28px 8px 28px;{bg_style}\">\n"
        f"            <div style=\"font-size:16px;font-weight:bold;color:#1a1a1a;"
        f"margin-bottom:14px;padding-bottom:6px;border-bottom:1px solid {divider_color};\">"
        f"{heading}</div>\n"
        "            <table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n"
        + rows_html + "\n"
        "            </table>\n"
        "          </td>\n"
        "        </tr>"
    )


def render_event_list(props: dict, theme: dict, **_) -> str:
    heading = html_module.escape(props.get("heading", "Save the Date!"))
    items = props.get("items", [])
    bg = props.get("backgroundColor", "#EEF6F7")
    bg_style = f"background-color:{bg};" if bg else ""
    divider_color = "#eeeeee"

    rows_html = "\n".join(
        "              <tr>\n"
        "                <td width=\"12\" valign=\"top\" style=\"padding:0 0 10px 0;"
        "color:#1a1a1a;font-size:18px;line-height:1;\">&bull;</td>\n"
        "                <td style=\"padding:0 0 10px 8px;font-size:13px;"
        "color:#1a1a1a;line-height:1.5;\">"
        f"{_render_text(str(item))}</td>\n"
        "              </tr>"
        for item in items if str(item).strip()
    )

    if not rows_html:
        return ""

    return (
        "\n"
        "        <!-- EVENT LIST -->\n"
        "        <tr>\n"
        f"          <td style=\"padding:24px 28px 8px 28px;{bg_style}\">\n"
        f"            <div style=\"font-size:16px;font-weight:bold;color:#1a1a1a;"
        f"margin-bottom:14px;padding-bottom:6px;border-bottom:1px solid {divider_color};\">"
        f"{heading}</div>\n"
        "            <table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\">\n"
        + rows_html + "\n"
        "            </table>\n"
        "          </td>\n"
        "        </tr>"
    )


def render_text_block(props: dict, theme: dict, **_) -> str:
    heading = html_module.escape(props.get("heading", ""))
    body = _render_text(props.get("body", ""))
    bg = props.get("backgroundColor", "#ffffff")
    bg_style = f"background-color:{bg};" if bg else ""

    return (
        "\n"
        "        <!-- TEXT BLOCK -->\n"
        "        <tr>\n"
        f"          <td style=\"padding:22px 28px;{bg_style}\">\n"
        + (
            f"            <div style=\"font-size:14px;font-weight:bold;color:#1a1a1a;"
            f"margin-bottom:8px;\">{heading}</div>\n"
            if heading else ""
        )
        + f"            <div style=\"font-size:13px;color:#444444;line-height:1.6;\">{body}</div>\n"
        "          </td>\n"
        "        </tr>"
    )


def render_product_card(props: dict, theme: dict, image_cids: dict = None, **_) -> str:
    title = html_module.escape(props.get("title", ""))
    body = _render_text(props.get("body", ""))
    creator = html_module.escape(props.get("creator", ""))
    image_url = props.get("imageUrl", "")
    bg = props.get("backgroundColor", "#ffffff")
    bg_style = f"background-color:{bg};" if bg else ""

    # Resolve image source: prefer CID for EML, direct URL for preview
    img_html = ""
    if image_url:
        if image_cids and image_url in image_cids:
            src = f"cid:{image_cids[image_url]}"
        else:
            src = image_url
        img_html = (
            f"            <img src=\"{src}\" alt=\"{html_module.escape(props.get('title', ''))}\" "
            "style=\"max-width:100%;height:auto;display:block;margin-bottom:10px;\">\n"
        )

    creator_html = (
        f"            <div style=\"font-size:11px;color:#888888;font-style:italic;"
        f"margin-top:6px;\">By {creator}</div>\n"
        if creator else ""
    )

    return (
        "\n"
        "        <!-- PRODUCT CARD -->\n"
        "        <tr>\n"
        f"          <td style=\"padding:16px 28px 12px 28px;{bg_style}"
        "border-bottom:1px solid #eeeeee;\">\n"
        + (
            f"            <div style=\"font-size:16px;font-weight:bold;color:#1a1a1a;"
            f"margin-bottom:10px;\">{title}</div>\n"
            if title else ""
        )
        + img_html
        + f"            <div style=\"font-size:13px;color:#444444;line-height:1.6;\">{body}</div>\n"
        + creator_html
        + "          </td>\n"
        "        </tr>"
    )


def render_image_block(props: dict, theme: dict, image_cids: dict = None, **_) -> str:
    image_url = props.get("imageUrl", "")
    alt = html_module.escape(props.get("altText", ""))
    caption = _render_text(props.get("caption", ""))
    bg = props.get("backgroundColor", "#ffffff")
    bg_style = f"background-color:{bg};" if bg else ""

    if not image_url:
        return ""

    if image_cids and image_url in image_cids:
        src = f"cid:{image_cids[image_url]}"
    else:
        src = image_url

    caption_html = (
        f"            <div style=\"font-size:11px;color:#767676;text-align:center;"
        f"margin-top:8px;font-style:italic;\">{caption}</div>\n"
        if caption else ""
    )

    return (
        "\n"
        "        <!-- IMAGE BLOCK -->\n"
        "        <tr>\n"
        f"          <td style=\"padding:16px 28px;{bg_style}\">\n"
        f"            <img src=\"{src}\" alt=\"{alt}\" "
        "style=\"max-width:100%;height:auto;display:block;\">\n"
        + caption_html
        + "          </td>\n"
        "        </tr>"
    )


def render_divider(props: dict, theme: dict, **_) -> str:
    color = props.get("color", "#E0E0E0")
    spacing = int(props.get("spacing", 20))

    return (
        "\n"
        "        <!-- DIVIDER -->\n"
        f"        <tr><td style=\"height:{spacing}px;font-size:0;line-height:0;\">&nbsp;</td></tr>\n"
        f"        <tr><td style=\"height:1px;background-color:{color};"
        "font-size:0;line-height:0;\">&nbsp;</td></tr>\n"
        f"        <tr><td style=\"height:{spacing}px;font-size:0;line-height:0;\">&nbsp;</td></tr>"
    )


# ---------------------------------------------------------------------------
# Section registry
# ---------------------------------------------------------------------------

SECTION_RENDERERS = {
    "header":       render_header,
    "footer":       render_footer,
    "bullet_list":  render_bullet_list,
    "event_list":   render_event_list,
    "text_block":   render_text_block,
    "product_card": render_product_card,
    "image_block":  render_image_block,
    "divider":      render_divider,
}

SECTION_LABELS = {
    "header":       "Header",
    "footer":       "Footer",
    "bullet_list":  "Bullet List",
    "event_list":   "Event / Save the Date",
    "text_block":   "Text Block",
    "product_card": "Product Card",
    "image_block":  "Image",
    "divider":      "Divider",
}


# ---------------------------------------------------------------------------
# Top-level renderer
# ---------------------------------------------------------------------------

def render_newsletter(config: dict, image_cids: dict = None) -> str:
    """
    Render a complete HTML email string from a newsletter JSON config.

    Args:
        config: Newsletter config dict with keys: meta, theme, sections
        image_cids: Optional mapping of imageUrl -> CID string for EML embedding.
                    If None, imageUrl values are used directly in <img src>.

    Returns:
        Complete HTML string.
    """
    if image_cids is None:
        image_cids = {}

    theme = config.get("theme", {})
    sections = config.get("sections", [])
    meta = config.get("meta", {})

    bg = theme.get("backgroundColor", "#F5F5F5")
    font = theme.get("fontFamily", "Arial,Helvetica,sans-serif")
    width = theme.get("tableWidth", 700)
    title = html_module.escape(meta.get("newsletterName", "Newsletter"))

    parts = []
    for section in sections:
        section_type = section.get("type", "")
        props = section.get("props", {})
        renderer = SECTION_RENDERERS.get(section_type)
        if renderer:
            rendered = renderer(props=props, theme=theme, image_cids=image_cids)
            if rendered:
                parts.append(rendered)

    inner_html = "\n".join(parts)

    return (
        "<!DOCTYPE html>\n"
        "<html lang=\"en\">\n"
        "<head>\n"
        "<meta charset=\"UTF-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n"
        f"<title>{title}</title>\n"
        "</head>\n"
        f"<body style=\"margin:0;padding:0;background-color:{bg};font-family:{font};\">\n"
        f"<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" "
        f"style=\"background-color:{bg};\">\n"
        "  <tr>\n"
        f"    <td align=\"center\" valign=\"top\" style=\"padding:20px 10px;"
        f"background-color:{bg};\">\n"
        f"      <table width=\"{width}\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\" border=\"0\" "
        f"bgcolor=\"#ffffff\" style=\"background-color:#ffffff !important;"
        f"border:1px solid #E0E0E0;margin:0 auto;\">\n"
        + inner_html + "\n"
        "      </table>\n"
        "    </td>\n"
        "  </tr>\n"
        "</table>\n"
        "</body>\n"
        "</html>"
    )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _lighten_bg(hex_color: str) -> str:
    """Return a very light tint of the given hex color for use as title area bg."""
    try:
        h = hex_color.lstrip("#")
        if len(h) == 3:
            h = h[0]*2 + h[1]*2 + h[2]*2
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        # Mix 90% white
        r2 = int(r * 0.1 + 255 * 0.9)
        g2 = int(g * 0.1 + 255 * 0.9)
        b2 = int(b * 0.1 + 255 * 0.9)
        return f"#{r2:02x}{g2:02x}{b2:02x}"
    except Exception:
        return "#f9f9f9"
