import json, sys, os
from pptx import Presentation

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Kick-Off Template.pptx")

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

def generate(config: dict, output_path: str):
    brand = config.get("brand_name", "Brand")
    date = config.get("kickoff_date", "TBD")
    csm_name = config.get("csm_name", "")
    csm_email = config.get("csm_email", "service@unidy.de")
    csm_phone = config.get("csm_phone", "")
    goals = config.get("goals", ["Goal 1", "Goal 2", "Goal 3", "Goal 4"])
    integrations_v1 = config.get("integrations_v1", [])
    go_live_date = config.get("go_live_date", "")
    notion_url = config.get("notion_url", "")

    while len(goals) < 4:
        goals.append("")

    parts = date.split("/")
    dd = parts[0] if len(parts) > 0 else ""
    mm = parts[1] if len(parts) > 1 else ""
    yyyy = parts[2] if len(parts) > 2 else ""

    prs = Presentation(TEMPLATE_PATH)

    # Slide 1: Cover
    for shape in prs.slides[0].shapes:
        replace_across_runs(shape, ["Brand", " ID"], f"{brand} ID")
        replace_run_text(shape, "dd", dd)
        replace_run_text(shape, "mm", mm)
        replace_run_text(shape, "yyyy", yyyy)
        replace_run_text(shape, "Add Brand Logo here", "")

    # Slide 9: CSM Contact
    for shape in prs.slides[8].shapes:
        if csm_name:
            replace_run_text(shape, "Infos hier:", csm_name)
        replace_run_text(shape, "service@unidy.de", csm_email)
        if csm_phone:
            replace_run_text(shape, "Handynummer für Notfälle", csm_phone)

    # Slide 12: Setup title
    for shape in prs.slides[11].shapes:
        replace_across_runs(shape, ['\u201cB', 'rand', '\u201d Setup'],
                            f'\u201c{brand}\u201d Setup')

    # Slide 13: Goals
    for shape in prs.slides[12].shapes:
        for i, goal in enumerate(goals[:4], 1):
            if goal:
                replace_run_text(shape, f"Goal {i}", goal)
        replace_run_text(shape, "Hier noch JGS Goals einbauen", "")
        replace_run_text(shape,
            "https://www.notion.so/unidy-gmbh/Customer-Outcomes-and-Customer-Personas-31e5a77a0d338015b860eca1bda9a136", "")

    # Slide 16: Next Steps
    for shape in prs.slides[15].shapes:
        if integrations_v1:
            integ_str = " · ".join(integrations_v1[:3])
            replace_run_text(shape,
                "Kunden kontaktiert alle Integrationspartner mit ",
                f"Integrationen V1: {integ_str} – Partner kontaktieren mit ")
        if go_live_date:
            replace_run_text(shape,
                "Go-Live Checklist on Customer Page with Customer To Dos",
                f"Go-Live Checklist · Ziel: {go_live_date}")
        if notion_url:
            replace_run_text(shape, "Erwähnung Notion Page", f"Notion: {notion_url}")

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
