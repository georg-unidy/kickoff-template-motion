import json, os, io, requests
from pptx import Presentation
from pptx.util import Emu
from pptx.dml.color import RGBColor

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Kick-Off Template.pptx")

# Default brand color used in template (YBB yellow / placeholder)
DEFAULT_FILL_COLOR = "F9CC11"
# Default text accent color used in template
DEFAULT_TEXT_COLOR = "0055FF"

def replace_run_text(shape, old, new):
    if not shape.has_text_frame:
        return False
    changed = False
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if old in run.text:
                run.text = run.text.replace(old, new)
                changed = True
    return changed

def replace_across_runs(shape, old_parts, new_text):
    if not shape.has_text_frame:
        return False
    for para in shape.text_frame.paragraphs:
        runs = para.runs
        for i in range(len(runs)):
            if i + len(old_parts) > len(runs):
                continue
            if all(runs[i+j].text == old_parts[j] for j in range(len(old_parts))):
                runs[i].text = new_text
                for j in range(1, len(old_parts)):
                    runs[i+j].text = ""
                return True
    return False

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip("#")
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))

def replace_fill_color(prs, old_hex, new_hex):
    """Replace all shape fill colors matching old_hex with new_hex."""
    new_rgb = hex_to_rgb(new_hex)
    old_hex = old_hex.upper()
    count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            try:
                current = str(shape.fill.fore_color.rgb).upper()
                if current == old_hex:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = new_rgb
                    count += 1
            except:
                pass
    return count

def replace_text_color(prs, old_hex, new_hex):
    """Replace all run font colors matching old_hex with new_hex."""
    new_rgb = hex_to_rgb(new_hex)
    old_hex = old_hex.upper()
    count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    try:
                        current = str(run.font.color.rgb).upper()
                        if current == old_hex:
                            run.font.color.rgb = new_rgb
                            count += 1
                    except:
                        pass
    return count

def insert_logo(slide, logo_url, left_emu, top_emu, width_emu, height_emu):
    """Download logo from URL and insert at given position, removing placeholder text."""
    # Remove "Add Brand Logo here" placeholder
    for shape in slide.shapes:
        replace_run_text(shape, "Add Brand Logo here", "")

    # Download image
    response = requests.get(logo_url, timeout=15)
    response.raise_for_status()
    img_data = io.BytesIO(response.content)

    slide.shapes.add_picture(img_data, left_emu, top_emu, width_emu, height_emu)

def generate(config: dict, output_path: str):
    brand        = config.get("brand_name", "Brand")
    date         = config.get("kickoff_date", "TBD")
    csm_name     = config.get("csm_name", "")
    csm_email    = config.get("csm_email", "service@unidy.de")
    csm_phone    = config.get("csm_phone", "")
    goals        = config.get("goals", ["Goal 1", "Goal 2", "Goal 3", "Goal 4"])
    integrations = config.get("integrations_v1", [])
    go_live_date = config.get("go_live_date", "")
    notion_url   = config.get("notion_url", "")
    logo_url     = config.get("logo_url", "")
    # CI colors — hex strings without #, e.g. "F9CC11"
    ci_primary   = config.get("ci_primary_color", "").lstrip("#").upper()
    ci_text      = config.get("ci_text_color", "").lstrip("#").upper()

    while len(goals) < 4:
        goals.append("")

    parts = date.split("/")
    dd   = parts[0] if len(parts) > 0 else ""
    mm   = parts[1] if len(parts) > 1 else ""
    yyyy = parts[2] if len(parts) > 2 else ""

    prs = Presentation(TEMPLATE_PATH)

    # ── Slide 1: Cover ──────────────────────────────────────────
    for shape in prs.slides[0].shapes:
        replace_across_runs(shape, ["Brand", " ID"], f"{brand} ID")
        replace_run_text(shape, "dd",   dd)
        replace_run_text(shape, "mm",   mm)
        replace_run_text(shape, "yyyy", yyyy)

    # Insert logo at same position as YBB (top right, ~4.47 x 4.47 cm)
    if logo_url:
        insert_logo(
            prs.slides[0],
            logo_url,
            left_emu   = Emu(10246150),
            top_emu    = Emu(315337),
            width_emu  = Emu(1610900),
            height_emu = Emu(1610900),
        )
    else:
        for shape in prs.slides[0].shapes:
            replace_run_text(shape, "Add Brand Logo here", "")

    # ── Slide 9: CSM Contact ─────────────────────────────────────
    for shape in prs.slides[8].shapes:
        if csm_name:
            replace_run_text(shape, "Infos hier:", csm_name)
        replace_run_text(shape, "service@unidy.de", csm_email)
        if csm_phone:
            replace_run_text(shape, "Handynummer für Notfälle", csm_phone)

    # ── Slide 12: Setup title ────────────────────────────────────
    for shape in prs.slides[11].shapes:
        replace_across_runs(shape, ['\u201cB', 'rand', '\u201d Setup'],
                            f'\u201c{brand}\u201d Setup')

    # ── Slide 13: Goals ──────────────────────────────────────────
    for shape in prs.slides[12].shapes:
        for i, goal in enumerate(goals[:4], 1):
            if goal:
                replace_run_text(shape, f"Goal {i}", goal)
        replace_run_text(shape, "Hier noch JGS Goals einbauen", "")
        replace_run_text(shape,
            "https://www.notion.so/unidy-gmbh/Customer-Outcomes-and-Customer-Personas-31e5a77a0d338015b860eca1bda9a136", "")

    # ── Slide 16: Next Steps ─────────────────────────────────────
    for shape in prs.slides[15].shapes:
        if integrations:
            integ_str = " · ".join(integrations[:3])
            replace_run_text(shape,
                "Kunden kontaktiert alle Integrationspartner mit ",
                f"Integrationen V1: {integ_str} – Partner kontaktieren mit ")
        if go_live_date:
            replace_run_text(shape,
                "Go-Live Checklist on Customer Page with Customer To Dos",
                f"Go-Live Checklist · Ziel: {go_live_date}")
        if notion_url:
            replace_run_text(shape, "Erwähnung Notion Page", f"Notion: {notion_url}")

    # ── CI Colors (global, applied last) ─────────────────────────
    # Shape fill colors (setup diagram boxes etc.)
    if ci_primary:
        replace_fill_color(prs, DEFAULT_FILL_COLOR, ci_primary)

    # Text accent color (brand name on cover, section labels etc.)
    if ci_text:
        replace_text_color(prs, DEFAULT_TEXT_COLOR, ci_text)
    elif ci_primary:
        # If only primary given, use it for text too
        replace_text_color(prs, DEFAULT_TEXT_COLOR, ci_primary)

    prs.save(output_path)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    with open(args.config) as f:
        config = json.load(f)
    generate(config, args.output)
