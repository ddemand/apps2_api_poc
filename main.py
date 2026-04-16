"""
An exploratory run of the OneVizion API on APPS2.
Author: Greg Tiffany, David Demand
Email: ddemand@onevizion.com
Date: 3/27/2026
"""

# ====== Install initial modules and dependencies
import subprocess
import sys
import glob
import json
import io
import os
subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'python_dependencies.ini'])

# ====== Import the required modules
import pandas as pd
import requests
from openai import OpenAI
import markdown
import re
from html import escape
from reportlab.platypus import Image
import logging
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, ListFlowable, ListItem
from reportlab.lib.units import inch
from reportlab.lib.colors import HexColor
from datetime import datetime,  timezone
from onevizion import Trackor
import urllib.parse

# ====== Set variables and other environmental items
pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", None)
pd.set_option("display.max_colwidth", None)

# Static assets
logo_path = "assets/logo.png"

# ====== Logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
    handlers=[
        logging.FileHandler("apps2_exec_report.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ====== Get keys
def get_key(key: str) -> str | None:
    file_path = "settings.json"
    try:
        with open(file_path, "r") as f:
            data = json.load(f)
        for api in data.get("keys", []):
            if api.get("key") == key:
                return api.get("value")
    except Exception as e:
        logger.exception("Failed loading API key: %s", key)
    return None

# ====== Get API key's, endpoints and URL's
apps2_api_key = get_key("apps2_api_key")
azure_foundry_api_key = get_key("azure_foundry")
azure_url = get_key("AZURE_URL")
data_url = get_key("DATA_URL")


url_subdomain = get_key("URL_SUBDOMAIN")
destination_field = get_key("DESTINATION_FIELD")
trackor_type = get_key("TRACKOR_TYPE")
csv_dimension_column = get_key("CSV_DIMENSION_COLUMN")
base_url = f'https://{url_subdomain}.onevizion.com/api/v3/trackor/'

if not all([apps2_api_key, azure_foundry_api_key, azure_url, url_subdomain, base_url]):
    raise RuntimeError("One or more required API keys are missing")

# ====== Get the data from the OneVizion API
logger.info("Fetching project data from OneVizion API")
response = requests.get(data_url, headers= {"Authorization": apps2_api_key, "Accept": "text/csv"})
response.raise_for_status()

csv_text = response.content.decode("utf-8-sig")
df = pd.read_csv(io.StringIO(csv_text))

logger.info("Loaded %d records", len(df))

# ====== Identify the unique records and their trackors
program_ids = df[csv_dimension_column].dropna().unique().tolist()
logger.info(f"Discovered %d unique {csv_dimension_column}", len(program_ids))

# ====== Create URL-friendly versions and build the dictionary
program_ids_url = [urllib.parse.quote(str(pid)) for pid in program_ids]   # ensure string
program_id_to_url: dict[str, str] = dict(zip(program_ids, program_ids_url))
logger.info("Created program_id_to_url dictionary with %d entries", len(program_id_to_url))

# ====== Use the API to get Trackor information
response = requests.get('https://apps2.onevizion.com/api/v3/trackor_types', headers= {
    "Authorization": apps2_api_key,
    "Accept": "application/json"}
                        )
response.raise_for_status()
trackor_list = json.loads(response.content.decode("utf-8-sig"))

# Extract the id and name for the selected trackor
try:
    program_trackor = next(
        item for item in trackor_list
        if isinstance(item, dict) and item.get("name") == trackor_type
    )
    trackor_type_id = program_trackor["id"]
    trackor_type_name    = program_trackor["name"]
    trackor_type_label   = program_trackor.get("label", "")
    trackor_type_prefix  = program_trackor.get("prefix", "")
    logger.info(f"{trackor_type} loaded → ID: {trackor_type_id}, Name: {trackor_type_name}, "
                f"Label: {trackor_type_label}, Prefix: {trackor_type_prefix}")
except StopIteration:
    raise ValueError(f"{trackor_type} not found in the API response") from None

# ====== Use an AI model to analyze the data (Instance name: 'apps2-azure-analysis')
client = OpenAI(
    base_url=azure_url,
    api_key=azure_foundry_api_key
)
deployment_name = "gpt-4.1"

# ====== Helper: Markdown → ReportLab-safe text
def md_to_reportlab(text: str) -> str:
    """
    Converts markdown **bold** into <b> tags and escapes everything else.
    """
    if not text:
        return ""

    text = escape(text)
    text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", text)
    return text

styles = getSampleStyleSheet()

styles.add(ParagraphStyle(
    name="TitleStyle",
    fontSize=18,
    leading=22,
    spaceAfter=16,
    textColor=HexColor("#2c3e50"),
    alignment=TA_LEFT
))

styles.add(ParagraphStyle(
    name="BodyStyle",
    fontSize=10.5,
    leading=14,
    spaceAfter=6
))

styles.add(ParagraphStyle(
    name="BulletStyle",
    fontSize=10.5,
    leading=14,
    leftIndent=12
))

styles.add(ParagraphStyle(
    name="H1",
    fontSize=16,
    leading=20,
    spaceBefore=20,
    spaceAfter=10,
    fontName="Helvetica-Bold",
    textColor=HexColor("#2c3e50")
))

styles.add(ParagraphStyle(
    name="H2",
    fontSize=14,
    leading=18,
    spaceBefore=18,
    spaceAfter=8,
    fontName="Helvetica-Bold",
    textColor=HexColor("#2c3e50")
))

styles.add(ParagraphStyle(
    name="H3",
    fontSize=12.5,
    leading=16,
    spaceBefore=16,
    spaceAfter=6,
    fontName="Helvetica-Bold",
    textColor=HexColor("#34495e")
))

# ====== Program-level analysis & PDF generation
for program_id in program_ids:
    logger.info(f"Processing {csv_dimension_column}: %s", program_id)

    program_df = df[df[csv_dimension_column] == program_id].copy()
    if program_df.empty:
        logger.warning(f"No data for {csv_dimension_column} %s", program_id)
        continue

    # Get URL-encoded key and resolve TRACKOR_ID
    if program_id not in program_id_to_url:
        logger.error(f"No URL mapping found for {csv_dimension_column}: %s", program_id)
        continue

    encoded_key = program_id_to_url[program_id]

    try:
        trackor_response = requests.get(
            f'https://apps2.onevizion.com/api/v3/trackor_types/{trackor_type}/trackors?XITOR_KEY={encoded_key}',
            headers={
                "Authorization": apps2_api_key,
                "Accept": "application/json"
            }
        )
        trackor_response.raise_for_status()
        trackor_data = trackor_response.json()

        if not trackor_data or len(trackor_data) == 0:
            logger.error(f"No Trackor found for {csv_dimension_column}: %s (encoded: %s)", program_id, encoded_key)
            continue

        # Take the first match (should usually be only one)
        trackor_id = trackor_data[0]["TRACKOR_ID"]
        logger.info(f"Resolved TRACKOR_ID %s for {csv_dimension_column} %s", trackor_id, program_id)

    except requests.exceptions.RequestException as e:
        logger.error("Failed to resolve Trackor ID for Program %s: %s", program_id, e)
        continue
    except (KeyError, IndexError, TypeError) as e:
        logger.error("Unexpected response format when resolving Trackor for %s: %s", program_id, e)
        continue

    # AI/LLM Analysis
    analysis_prompt = f"""
        You are an expert Program Director overseeing telecom network deployments.
        
        Analyze ONLY {csv_dimension_column}: {program_id}
        
        Data (JSON):
        {program_df.to_dict(orient="records")}
        
        Use the following markdown structure exactly:
        
        # Summary Details
        # Overall Program Health Status
        - Bullet points
        
        # Key Projects Overview and Milestones
        # Risks & Mitigations
        # Strategic Recommendations
        
        Be concise, executive-focused, data-driven, and actionable.
        Highlight owners, risks, and critical dates using **bold** text.
        """

    try:
        completion = client.chat.completions.create(
            model=deployment_name,
            messages=[{"role": "user", "content": analysis_prompt}],
            temperature=0.3,   # optional: lower for more consistent executive tone
        )

        executive_summary = completion.choices[0].message.content.strip()

    except Exception as e:
        logger.error("LLM call failed for Program %s: %s", program_id, e)
        continue

    # PDF Generation
    pdf_filename = f"Executive_Project_Summary_{program_id.replace(' ', '_')}.pdf"

    doc = SimpleDocTemplate(
        pdf_filename,
        pagesize=LETTER,
        rightMargin=50,
        leftMargin=50,
        topMargin=50,
        bottomMargin=50
    )

    elements = []

    # Logo
    if os.path.exists(logo_path):
        elements.append(Image(logo_path, width=2.44 * inch, height=0.49 * inch, hAlign="LEFT"))
        elements.append(Spacer(1, 0.25 * inch))

    # Title
    elements.append(Paragraph(
        f"Executive Summary — Program {program_id}",
        styles["TitleStyle"]
    ))

    # Updated Date with UTC Timestamp
    now_utc = datetime.now(timezone.utc)
    date_str = now_utc.strftime('%B %d, %Y')
    time_str = now_utc.strftime('%H:%M:%S UTC')

    elements.append(Paragraph(
        f"<b>{csv_dimension_column}:</b> {program_id}<br/>"
        f"<b>Date:</b> {date_str} at {time_str}",
        styles["BodyStyle"]
    ))

    elements.append(Spacer(1, 0.3 * inch))

    # Parse markdown into ReportLab elements
    bullet_buffer = []

    for line in executive_summary.splitlines():
        raw = line.strip()
        if not raw:
            continue

        if raw.startswith("#"):
            # Flush any pending bullets
            if bullet_buffer:
                elements.append(ListFlowable(
                    [ListItem(Paragraph(b, styles["BulletStyle"])) for b in bullet_buffer],
                    bulletType="bullet"
                ))
                bullet_buffer = []

            level = len(raw) - len(raw.lstrip("#"))
            text = raw.lstrip("#").strip()
            style_name = "H1" if level == 1 else "H2" if level == 2 else "H3"
            elements.append(Paragraph(md_to_reportlab(text), styles[style_name]))
            continue

        if raw.startswith("- "):
            bullet_buffer.append(md_to_reportlab(raw.lstrip("- ").strip()))
            continue

        # Flush bullets if we hit normal text
        if bullet_buffer:
            elements.append(ListFlowable(
                [ListItem(Paragraph(b, styles["BulletStyle"])) for b in bullet_buffer],
                bulletType="bullet"
            ))
            bullet_buffer = []

        elements.append(Paragraph(md_to_reportlab(raw), styles["BodyStyle"]))

    # Don't forget any trailing bullets
    if bullet_buffer:
        elements.append(ListFlowable(
            [ListItem(Paragraph(b, styles["BulletStyle"])) for b in bullet_buffer],
            bulletType="bullet"
        ))

    # Build the PDF
    try:
        doc.build(elements)
        logger.info("Successfully generated PDF: %s", pdf_filename)
    except Exception as e:
        logger.error("Failed to build PDF for Program %s: %s", program_id, e)
        continue

    # Upload PDF to the current Trackor
    try:
        with open(pdf_filename, "rb") as f:
            files = {"file": (pdf_filename, f, "application/pdf")}

            upload_url = f'https://apps2.onevizion.com/api/v3/trackor/{trackor_id}/file/{destination_field}'

            upload_response = requests.post(
                upload_url,
                headers={
                    "Accept": "*/*",
                    "Authorization": apps2_api_key
                },
                params={"file_name": pdf_filename},
                files=files
            )

        logger.info(
            "Upload result for Program %s (Trackor ID %s): %s - %s",
            program_id, trackor_id, upload_response.status_code, upload_response.text[:500]
        )

        if upload_response.status_code not in (200, 201, 204):
            logger.warning("Upload may have failed for %s", program_id)

    except Exception as e:
        logger.error("Failed to upload PDF for Program %s: %s", program_id, e)