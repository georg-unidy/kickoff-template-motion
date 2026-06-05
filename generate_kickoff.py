#!/usr/bin/env python3
"""
Kick-Off PPTX Generator
Nimmt das Template und füllt kundenspezifische Daten ein.
Usage: python generate_kickoff.py --config config.json --output output.pptx
"""

import json
import sys
import os
import shutil
import re
import subprocess
from pathlib import Path

SKILLS_DIR = "/mnt/skills/public/pptx/scripts"
TEMPLATE_PATH = "/home/claude/slides_input/Kick-Off Template.pptx"
WORK_DIR = "/home/claude/kickoff_work"

def run(cmd, cwd=None):
    r = subprocess.run(cmd, shell=True, capture_output=True, text=True, cwd=cwd)
    if r.returncode != 0:
        print(f"ERROR: {cmd}\n{r.stderr}", file=sys.stderr)
        sys.exit(1)
    return r.stdout

def xml_escape(text):
    return (text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;"))

def replace_in_file(path, old, new):
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    if old in content:
        content = content.replace(old, new)
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    return False

def edit_slide(slide_num, replacements, work_dir):
    """Apply text replacements to a slide XML file."""
    slide_path = f"{work_dir}/ppt/slides/slide{slide_num}.xml"
    if not os.path.exists(slide_path):
        print(f"  WARNING: slide{slide_num}.xml not found, skipping")
        return
    for old, new in replacements.items():
        replace_in_file(slide_path, old, xml_escape(new))

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
    
    # Ensure 4 goals (pad or trim)
    while len(goals) < 4:
        goals.append("")
    
    # Setup/reset work directory
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)
    
    # Unpack template
    print("📦 Entpacke Template...")
    run(f"python {SKILLS_DIR}/office/unpack.py '{TEMPLATE_PATH}' '{WORK_DIR}'")
    
    # --- Slide 1: Cover (Brand ID, Date) ---
    print("✏️  Slide 1: Cover")
    s1 = f"{WORK_DIR}/ppt/slides/slide1.xml"
    with open(s1, 'r', encoding='utf-8') as f:
        content = f.read()
    # Replace "Brand" + " ID" → "{brand} ID" (it's split across two runs)
    # The first run has "Brand" as the text
    content = content.replace('<a:t>Brand</a:t>', f'<a:t>{xml_escape(brand)}</a:t>', 1)
    # Replace date parts: dd, mm, yyyy
    content = content.replace('<a:t>dd</a:t>', '<a:t>' + xml_escape(date.split('/')[0] if '/' in date else date[:2]) + '</a:t>', 1)
    content = content.replace('<a:t>mm</a:t>', '<a:t>' + xml_escape(date.split('/')[1] if '/' in date and len(date.split('/')) > 1 else '') + '</a:t>', 1)
    content = content.replace('<a:t>yyyy</a:t>', '<a:t>' + xml_escape(date.split('/')[2] if '/' in date and len(date.split('/')) > 2 else '') + '</a:t>', 1)
    with open(s1, 'w', encoding='utf-8') as f:
        f.write(content)
    
    # --- Slide 9: CSM Contact ---
    print("✏️  Slide 9: CSM Kontakt")
    s9 = f"{WORK_DIR}/ppt/slides/slide9.xml"
    with open(s9, 'r', encoding='utf-8') as f:
        content = f.read()
    # Replace "Infos hier:" placeholder area - replace contact details
    content = content.replace('<a:t>service@unidy.de</a:t>', f'<a:t>{xml_escape(csm_email)}</a:t>')
    if csm_phone:
        content = content.replace('<a:t>Handynummer für Notfälle</a:t>', f'<a:t>{xml_escape(csm_phone)}</a:t>')
    if csm_name:
        content = content.replace('<a:t>Infos hier:</a:t>', f'<a:t>{xml_escape(csm_name)}</a:t>')
    with open(s9, 'w', encoding='utf-8') as f:
        f.write(content)
    
    # --- Slide 12: Setup / Integrations ---
    print("✏️  Slide 12: Setup")
    s12 = f"{WORK_DIR}/ppt/slides/slide12.xml"
    with open(s12, 'r', encoding='utf-8') as f:
        content = f.read()
    # "Brand" in title is split: "&#x201C;B" + "rand" + "&#x201D; Setup"
    # Replace 'rand' that's part of "Brand" in title
    content = content.replace('<a:t>&#x201C;B</a:t>', f'<a:t>&#x201C;{xml_escape(brand[0])}</a:t>', 1)
    content = content.replace('<a:t>rand</a:t>', f'<a:t>{xml_escape(brand[1:])}</a:t>', 1)
    with open(s12, 'w', encoding='utf-8') as f:
        f.write(content)
    
    # --- Slide 13: Goals ---
    print("✏️  Slide 13: Goals")
    s13 = f"{WORK_DIR}/ppt/slides/slide13.xml"
    with open(s13, 'r', encoding='utf-8') as f:
        content = f.read()
    for i, goal in enumerate(goals[:4], 1):
        content = content.replace(f'<a:t>Goal {i}</a:t>', f'<a:t>{xml_escape(goal)}</a:t>')
    with open(s13, 'w', encoding='utf-8') as f:
        f.write(content)
    
    # --- Slide 16: Next Steps ---
    print("✏️  Slide 16: Next Steps")
    s16 = f"{WORK_DIR}/ppt/slides/slide16.xml"
    with open(s16, 'r', encoding='utf-8') as f:
        content = f.read()
    if go_live_date:
        content = content.replace(
            '<a:t>Go-Live Checklist on Customer Page with Customer To Dos</a:t>',
            f'<a:t>Go-Live Checklist on Customer Page · Ziel: {xml_escape(go_live_date)}</a:t>'
        )
    if notion_url:
        content = content.replace(
            '<a:t>Erwähnung Notion Page</a:t>',
            f'<a:t>Notion Page: {xml_escape(notion_url)}</a:t>'
        )
    if integrations_v1:
        integ_str = " · ".join(integrations_v1[:3])
        content = content.replace(
            '<a:t>Kunden kontaktiert alle Integrationspartner mit </a:t>',
            f'<a:t>Integrationen V1: {xml_escape(integ_str)} – Partner kontaktieren mit </a:t>'
        )
    with open(s16, 'w', encoding='utf-8') as f:
        f.write(content)
    
    # Clean and pack
    print("🧹 Bereinige...")
    run(f"python {SKILLS_DIR}/clean.py '{WORK_DIR}'")
    
    print("📦 Packe PPTX...")
    run(f"python {SKILLS_DIR}/office/pack.py '{WORK_DIR}' '{output_path}' --original '{TEMPLATE_PATH}'")
    
    print(f"\n✅ Fertig! Gespeichert unter: {output_path}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True, help="JSON config file")
    parser.add_argument("--output", required=True, help="Output PPTX path")
    args = parser.parse_args()
    
    with open(args.config) as f:
        config = json.load(f)
    
    generate(config, args.output)
