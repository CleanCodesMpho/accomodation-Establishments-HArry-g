from fastapi import FastAPI, Request
from arcgis.gis import GIS
from arcgis.features import FeatureLayer
import os
import qrcode
from datetime import datetime

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm

# =========================
# APP INIT
# =========================
app = FastAPI()

# =========================
# AGOL AUTH
# =========================
AGOL_USERNAME = os.getenv("AGOL_USERNAME")
AGOL_PASSWORD = os.getenv("AGOL_PASSWORD")

if not AGOL_USERNAME or not AGOL_PASSWORD:
    raise Exception("AGOL credentials not set in environment variables")

gis = GIS("https://www.arcgis.com", AGOL_USERNAME, AGOL_PASSWORD)

# =========================
# FEATURE LAYER
# =========================
SURVEY_LAYER_URL = "https://services6.arcgis.com/345WScIubRHps95b/arcgis/rest/services/service_bd43c481ce0345febf5fc02b8ec3b09f/FeatureServer/FeatureServer/0"
layer = FeatureLayer(SURVEY_LAYER_URL, gis=gis)

# =========================
# TEMPLATE
# =========================
TEMPLATE_PATH = "accomodation_establishments_Report.docx"

if not os.path.exists(TEMPLATE_PATH):
    raise Exception(f"{TEMPLATE_PATH} not found in project root")

# =========================
# TEMP PAYLOAD STORAGE
# =========================
LAST_PAYLOAD = {}
LAST_ERROR = None

# =========================
# HEALTH CHECK
# =========================
@app.get("/")
def home():
    return {"status": "running"}

# =========================
# DEBUG ENDPOINT
# =========================
@app.get("/debug")
def debug():
    return {
        "template_exists": os.path.exists(TEMPLATE_PATH),
        "username_set": bool(AGOL_USERNAME),
        "password_set": bool(AGOL_PASSWORD),
        "layer_url": SURVEY_LAYER_URL
    }

# =========================
# LAST PAYLOAD
# =========================
@app.get("/last-payload")
def last_payload():
    return {
        "last_error": LAST_ERROR,
        "payload": LAST_PAYLOAD
    }

# =========================
# TEST QUERY
# =========================
@app.get("/test-query/{objectid}")
def test_query(objectid: int):
    result = layer.query(where=f"OBJECTID={objectid}", out_fields="*")
    return {
        "found": len(result.features),
        "attributes": result.features[0].attributes if result.features else None
    }

# =========================
# TEST UPDATE
# =========================
@app.get("/test-update/{objectid}")
def test_update(objectid: int):
    result = layer.edit_features(updates=[{
        "attributes": {
            "OBJECTID": objectid,
            "report_status": "test_ok",
            "report_url": "https://example.com/test.docx"
        }
    }])
    return {"edit_result": result}

# =========================
# HELPER: EXTRACT OBJECTID
# =========================
def extract_objectid(payload):
    if "submittedRecord" in payload:
        attrs = payload["submittedRecord"].get("attributes", {})
        if "OBJECTID" in attrs:
            return attrs["OBJECTID"]

    if "serverResponse" in payload:
        sr = payload["serverResponse"]
        if isinstance(sr, dict):
            if "objectId" in sr:
                return sr["objectId"]
            if "editResults" in sr and sr["editResults"]:
                first = sr["editResults"][0]
                if "objectId" in first:
                    return first["objectId"]

    if "feature" in payload:
        feature = payload["feature"]
        if isinstance(feature, dict):
            attrs = feature.get("attributes", {})
            if "OBJECTID" in attrs:
                return attrs["OBJECTID"]
            result = feature.get("result", {})
            if "objectId" in result:
                return result["objectId"]

    if "features" in payload and payload["features"]:
        first = payload["features"][0]
        attrs = first.get("attributes", {})
        if "OBJECTID" in attrs:
            return attrs["OBJECTID"]

    if isinstance(payload, dict):
        for key, value in payload.items():
            if key in ("OBJECTID", "objectId"):
                return value
            found = extract_objectid(value)
            if found is not None:
                return found

    if isinstance(payload, list):
        for item in payload:
            found = extract_objectid(item)
            if found is not None:
                return found

    return None

# =========================
# QR GENERATOR
# =========================
def generate_qr(url, path):
    img = qrcode.make(url)
    img.save(path)

# =========================
# UPLOAD REPORT TO AGOL
# =========================
def upload_report_to_agol(file_path, objectid):
    root_folder = gis.content.folders.get()

    item_properties = {
        "title": f"Report_{objectid}",
        "type": "Microsoft Word",
        "tags": ["survey123", "report", "automation"],
        "snippet": f"Automatically generated report for Survey123 submission {objectid}"
    }

    report_item = root_folder.add(
        item_properties=item_properties,
        file=file_path
    ).result()

    report_item.sharing.sharing_level = "EVERYONE"

    return f"https://www.arcgis.com/home/item.html?id={report_item.itemid}"

# =========================
# REPORT GENERATION
# =========================
def generate_report(attributes, objectid):
    os.makedirs("output", exist_ok=True)

    docx_file = os.path.join("output", f"report_{objectid}.docx")
    qr_file = os.path.join("output", f"qr_{objectid}.png")

    # Temporary QR target for first render
    temp_url = f"https://www.arcgis.com/home/item.html?id=temp-{objectid}"
    generate_qr(temp_url, qr_file)

    edit_date = attributes.get("EditDate")
    if edit_date:
        edit_date = datetime.fromtimestamp(edit_date / 1000).strftime("%Y-%m-%d %H:%M:%S")
    else:
        edit_date = "N/A"

    doc = DocxTemplate(TEMPLATE_PATH)
    qr_image = InlineImage(doc, qr_file, width=Mm(25))

    context = {
        "municipality": attributes.get("municipality", "N/A"),
        "precommise_name": attributes.get("precommise_name", "N/A"),
        "address": attributes.get("address", "N/A"),
        "Precommise_Type": attributes.get("Precommise_Type", "N/A"),
        "owner_name": attributes.get("owner_name", "N/A"),
        "contact": attributes.get("contact", "N/A"),
        "inspection_date": edit_date,
        "EHP": attributes.get("EHP", "N/A"),

        "Q1": attributes.get("Q1", "N/A"),
        "recomm1": attributes.get("recomm1", "N/A"),
        "Q2": attributes.get("Q2", "N/A"),
        "recomm2": attributes.get("recomm2", "N/A"),
        "Q3": attributes.get("Q3", "N/A"),
        "recomm3": attributes.get("recomm3", "N/A"),
        "Q4": attributes.get("Q4", "N/A"),
        "recomm4": attributes.get("recomm4", "N/A"),
        "Q5": attributes.get("Q5", "N/A"),
        "recomm5": attributes.get("recomm5", "N/A"),

        "Q6": attributes.get("Q6", "N/A"),
        "recomm6": attributes.get("recomm6", "N/A"),
        "Q7": attributes.get("Q7", "N/A"),
        "recomm7": attributes.get("recomm7", "N/A"),
        "Q8": attributes.get("Q8", "N/A"),
        "recomm8": attributes.get("recomm8", "N/A"),
        "Q9": attributes.get("Q9", "N/A"),
        "recomm9": attributes.get("recomm9", "N/A"),
        "Q10": attributes.get("Q10", "N/A"),
        "recomm10": attributes.get("recomm10", "N/A"),

        "Q11": attributes.get("Q11", "N/A"),
        "recomm11": attributes.get("recomm11", "N/A"),
        "Q12": attributes.get("Q12", "N/A"),
        "recomm12": attributes.get("recomm12", "N/A"),
        "Q13": attributes.get("Q13", "N/A"),
        "recomm13": attributes.get("recomm13", "N/A"),
        "Q14": attributes.get("Q14", "N/A"),
        "recomm14": attributes.get("recomm14", "N/A"),
        "Q15": attributes.get("Q15", "N/A"),
        "recomm15": attributes.get("recomm15", "N/A"),

        "Q16": attributes.get("Q16", "N/A"),
        "recomm16": attributes.get("recomm16", "N/A"),
        "Q17": attributes.get("Q17", "N/A"),
        "recomm17": attributes.get("recomm17", "N/A"),
        "Q18": attributes.get("Q18", "N/A"),
        "recomm18": attributes.get("recomm18", "N/A"),
        "Q19": attributes.get("Q19", "N/A"),
        "recomm19": attributes.get("recomm19", "N/A"),

        "Q20": attributes.get("Q20", "N/A"),
        "recomm20": attributes.get("recomm20", "N/A"),
        "Q21": attributes.get("Q21", "N/A"),
        "recomm21": attributes.get("recomm21", "N/A"),
        "Q22": attributes.get("Q22", "N/A"),
        "recomm22": attributes.get("recomm22", "N/A"),
        "Q23": attributes.get("Q23", "N/A"),
        "recomm23": attributes.get("recomm23", "N/A"),

        "Q24": attributes.get("Q24", "N/A"),
        "recomm24": attributes.get("recomm24", "N/A"),
        "Q25": attributes.get("Q25", "N/A"),
        "recomm25": attributes.get("recomm25", "N/A"),
        "Q26": attributes.get("Q26", "N/A"),
        "recomm26": attributes.get("recomm26", "N/A"),
        "Q27": attributes.get("Q27", "N/A"),
        "recomm27": attributes.get("recomm27", "N/A"),

        "Q28": attributes.get("Q28", "N/A"),
        "recomm28": attributes.get("recomm28", "N/A"),
        "Q29": attributes.get("Q29", "N/A"),
        "recomm29": attributes.get("recomm29", "N/A"),
        
        "compliance": attributes.get("compliance", "N/A"),
        "recommedations_": attributes.get("recommedations_", "N/A"),
        "person_incharge": attributes.get("person_incharge", "N/A"),
        "contacts": attributes.get("contacts", "N/A"),
        "ehp_email_address": attributes.get("ehp_email_address", "N/A"),
        "additional_pictures": attributes.get("additional_pictures", "N/A"),
        "risk_rating": attributes.get("risk_rating", "N/A"),
        "action_taken": attributes.get("action_taken", "N/A"),
        "additional_pictures": attributes.get("additional_pictures", "N/A"),
        "EHP": attributes.get("EHP", "N/A"),
      
        "qr_code": qr_image
    }

    doc.render(context)
    doc.save(docx_file)

    real_url = upload_report_to_agol(docx_file, objectid)
    return real_url

# =========================
# UPDATE FEATURE
# =========================
def update_feature(objectid, url, status):
    result = layer.edit_features(updates=[{
        "attributes": {
            "OBJECTID": objectid,
            "report_url": url,
            "report_status": status
        }
    }])
    return result

# =========================
# WEBHOOK ENDPOINT
# =========================
@app.post("/webhook/survey123")
async def survey_webhook(request: Request):
    global LAST_PAYLOAD, LAST_ERROR

    payload = await request.json()
    LAST_PAYLOAD = payload
    LAST_ERROR = None
    objectid = None

    try:
        objectid = extract_objectid(payload)

        if objectid is None:
            LAST_ERROR = f"OBJECTID not found. Payload keys: {list(payload.keys()) if isinstance(payload, dict) else 'not a dict'}"
            return {
                "status": "failed",
                "error": LAST_ERROR
            }

        update_feature(objectid, "webhook_received", "received")

        result = layer.query(where=f"OBJECTID={objectid}", out_fields="*")

        if not result.features:
            update_feature(objectid, "query_failed", "failed")
            LAST_ERROR = f"No feature found for OBJECTID {objectid}"
            return {
                "status": "failed",
                "error": LAST_ERROR
            }

        attributes = result.features[0].attributes

        update_feature(objectid, "query_ok", "queried")

        report_url = generate_report(attributes, objectid)

        edit_result = update_feature(objectid, report_url, "completed")

        return {
            "status": "success",
            "objectid": objectid,
            "report_url": report_url,
            "edit_result": str(edit_result)
        }

    except Exception as e:
        LAST_ERROR = str(e)
        if objectid is not None:
            try:
                update_feature(objectid, f"ERROR: {str(e)}", "failed")
            except Exception:
                pass

        return {
            "status": "failed",
            "objectid": objectid,
            "error": str(e)
        }
