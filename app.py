import os
import json
import tempfile
import re
import google.generativeai as genai
import gspread
import mimetypes
import streamlit as st
from google.oauth2.service_account import Credentials
from openpyxl.utils import get_column_letter

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

DEFAULT_SHEET_ID = "17tMHStXQYXaIQHQIA4jdUyHaYt_tuoNCEEuJCstWEuw"

COMPANY_NAME_ROW = 1
CONTACT_INFO_ROW = 2
HEADER_ROW = 3
ITEM_MASTER_LIST_COL = 2
COLUMNS_PER_SUPPLIER = 4

def extract_sheet_id_from_url(url):
    if not url:
        return None
    if "/" not in url and " " not in url and len(url) > 20:
        return url
    m = re.search(r"spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    return m.group(1) if m else None

def extract_json_from_text(text):
    start = text.find("{")
    end = text.rfind("}") + 1
    if start >= 0 and end > start:
        json_str = text[start:end]
        cleaned_json = re.sub(r",\s*}", "}", json_str)
        cleaned_json = re.sub(r",\s*]", "]", cleaned_json)
        try:
            return json.loads(cleaned_json)
        except json.JSONDecodeError:
            return None
    return None

def extract_contact_info(text):
    if not text:
        return ""
    phone_pattern = r"(?<!\w)((0\d{1,2}[-\s]?\d{3}[-\s]?\d{3,4})|(0\d{2}[-\s]?\d{7})|(0\d{2}[-\s]?\d{3}[-\s]?\d{4}))(?!\w)"
    email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    phone_matches = re.findall(phone_pattern, text)
    email_matches = re.findall(email_pattern, text)
    phone_numbers = [m[0] for m in phone_matches] if phone_matches else []
    contact_parts = []
    if email_matches:
        contact_parts.append(f"Email: {', '.join(email_matches)}")
    if phone_numbers:
        formatted_phones = []
        for phone in phone_numbers:
            clean_phone = re.sub(r"\s", "", phone)
            if len(clean_phone) >= 9:
                formatted_phones.append(clean_phone)
        contact_parts.append(f"Phone: {', '.join(formatted_phones)}")
    return ", ".join(contact_parts)

def clean_product_name(name):
    if not name:
        return "Unknown Product"
    return re.sub(r"^\d+\.\s*", "", name.strip())

def validate_json_data(json_data):
    if not json_data:
        return {
            "company": "Unknown Company",
            "contact": "",
            "products": [],
            "totalPrice": 0,
            "totalVat": 0,
            "totalPriceIncludeVat": 0,
            "priceGuaranteeDay": 0,
            "deliveryTime": "",
            "paymentTerms": "",
            "otherNotes": "",
        }
    if not json_data.get("company"):
        json_data["company"] = "Unknown Company"
    if "contact" in json_data:
        if isinstance(json_data["contact"], dict):
            contact_parts = []
            if "email" in json_data["contact"]:
                contact_parts.append(f"Email: {json_data['contact']['email']}")
            if "phone" in json_data["contact"]:
                contact_parts.append(f"Phone: {json_data['contact']['phone']}")
            json_data["contact"] = ", ".join(contact_parts)
        else:
            json_data["contact"] = extract_contact_info(str(json_data["contact"]))
    else:
        json_data["contact"] = ""
    if not json_data.get("products"):
        json_data["products"] = []
    for product in json_data.get("products", []):
        if product.get("name"):
            product["name"] = clean_product_name(product["name"])
        else:
            product["name"] = "Unknown Product"
        if not product.get("quantity"):
            product["quantity"] = 1
        else:
            product["quantity"] = (
                float(str(product["quantity"]).replace(",", ""))
                if str(product["quantity"])
                .replace(",", "")
                .replace(".", "", 1)
                .isdigit()
                else 1
            )
        if not product.get("unit"):
            product["unit"] = "ชิ้น"
        if not product.get("pricePerUnit"):
            product["pricePerUnit"] = 0
        else:
            product["pricePerUnit"] = (
                float(str(product["pricePerUnit"]).replace(",", ""))
                if str(product["pricePerUnit"])
                .replace(",", "")
                .replace(".", "", 1)
                .isdigit()
                else 0
            )
        if not product.get("totalPrice"):
            product["totalPrice"] = round(
                product["quantity"] * product["pricePerUnit"], 2
            )
        else:
            product["totalPrice"] = (
                float(str(product["totalPrice"]).replace(",", ""))
                if str(product["totalPrice"])
                .replace(",", "")
                .replace(".", "", 1)
                .isdigit()
                else round(product["quantity"] * product["pricePerUnit"], 2)
            )
    if not json_data.get("totalPrice"):
        json_data["totalPrice"] = sum(
            p.get("totalPrice", 0) for p in json_data.get("products", [])
        )
    else:
        json_data["totalPrice"] = (
            float(str(json_data["totalPrice"]).replace(",", ""))
            if str(json_data["totalPrice"])
            .replace(",", "")
            .replace(".", "", 1)
            .isdigit()
            else sum(p.get("totalPrice", 0) for p in json_data.get("products", []))
        )
    if not json_data.get("totalVat"):
        json_data["totalVat"] = round(json_data["totalPrice"] * 0.07, 2)
    else:
        json_data["totalVat"] = (
            float(str(json_data["totalVat"]).replace(",", ""))
            if str(json_data["totalVat"]).replace(",", "").replace(".", "", 1).isdigit()
            else round(json_data["totalPrice"] * 0.07, 2)
        )
    if not json_data.get("totalPriceIncludeVat"):
        json_data["totalPriceIncludeVat"] = round(
            json_data["totalPrice"] + json_data["totalVat"], 2
        )
    else:
        json_data["totalPriceIncludeVat"] = (
            float(str(json_data["totalPriceIncludeVat"]).replace(",", ""))
            if str(json_data["totalPriceIncludeVat"])
            .replace(",", "")
            .replace(".", "", 1)
            .isdigit()
            else round(json_data["totalPrice"] + json_data["totalVat"], 2)
        )
    if "priceGuaranteeDay" not in json_data:
        json_data["priceGuaranteeDay"] = 0
    if "deliveryTime" not in json_data:
        json_data["deliveryTime"] = ""
    if "paymentTerms" not in json_data:
        json_data["paymentTerms"] = ""
    if "otherNotes" not in json_data:
        json_data["otherNotes"] = ""
    return json_data

def enhance_with_gemini(json_data):
    if not json_data:
        return None
    validation_formatted = validation_prompt.format(
        extracted_json=json.dumps(json_data, ensure_ascii=False)
    )
    model = genai.GenerativeModel(model_name="gemini-2.5-flash")
    response = model.generate_content(validation_formatted)
    enhanced_text = response.text.strip()
    if "```" in enhanced_text:
        blocks = re.findall(r"```(?:json)?(.*?)```", enhanced_text, re.DOTALL)
        if blocks:
            enhanced_text = blocks[0].strip()
    enhanced_data = json.loads(enhanced_text) if enhanced_text else json_data
    if not isinstance(enhanced_data, dict):
        enhanced_data = extract_json_from_text(enhanced_text)
        return enhanced_data if enhanced_data else json_data
    return enhanced_data

def match_products_with_gemini(target_products, reference_products):
    if not target_products or not reference_products:
        return {"matchedItems": [], "uniqueItems": target_products or []}
    match_prompt_formatted = matching_prompt.format(
        target_products=json.dumps(target_products, ensure_ascii=False),
        reference_products=json.dumps(reference_products, ensure_ascii=False),
    )
    model = genai.GenerativeModel(
        model_name="gemini-2.5-pro",
        generation_config={"temperature": 0.1, "top_p": 0.95},
    )
    response = model.generate_content(match_prompt_formatted)
    match_text = response.text.strip()
    if "```" in match_text:
        blocks = re.findall(r"```(?:json)?(.*?)```", match_text, re.DOTALL)
        if blocks:
            match_text = blocks[0].strip()
    match_data = None
    try:
        match_data = json.loads(match_text)
    except json.JSONDecodeError:
        match_data = extract_json_from_text(match_text)
    if not match_data:
        return {"matchedItems": [], "uniqueItems": target_products}
    if "matchedItems" not in match_data:
        match_data["matchedItems"] = []
    if "uniqueItems" not in match_data:
        match_data["uniqueItems"] = target_products
    return match_data

def authenticate_and_open_sheet(sheet_id):
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(sheet_id).get_worksheet(0)

def ensure_first_three_rows_exist(ws):
    p = []
    for i in range(1, 4):
        p.append({"range": f"A{i}:B{i}", "values": [["", ""]]})
    ws.batch_update(p, value_input_option="USER_ENTERED")

def find_next_available_column(ws):
    vals = ws.get_all_values()
    m = ITEM_MASTER_LIST_COL
    for row in vals[:HEADER_ROW]:
        for i, c in enumerate(row):
            if c.strip():
                m = max(m, i + 1)
    return m + 1

def update_google_sheet_for_single_file(ws, data):
    ensure_first_three_rows_exist(ws)
    start_row = HEADER_ROW + 1
    sheet_values = ws.get_all_values()
    existing_products = []
    SUMMARY_LABELS = [
        "รวมเป็นเงิน",
        "ภาษีมูลค่าเพิ่ม 7%",
        "ยอดรวมทั้งสิ้น",
        "กำหนดยืนราคา (วัน)",
        "ระยะเวลาส่งมอบสินค้าหลังจากได้รับ PO",
        "การชำระเงิน",
        "อื่น ๆ",
    ]
    summary_row_map = {}
    first_summary_row = -1

    for row_idx, row in enumerate(sheet_values[HEADER_ROW:], start=start_row):
        if len(row) >= ITEM_MASTER_LIST_COL and row[ITEM_MASTER_LIST_COL - 1].strip():
            cell_value = row[ITEM_MASTER_LIST_COL - 1].strip()
            if cell_value in SUMMARY_LABELS:
                if first_summary_row == -1:
                    first_summary_row = row_idx
                summary_row_map[cell_value] = row_idx
            else:
                product_name = clean_product_name(cell_value)
                existing_products.append({"name": product_name, "row": row_idx})

    existing_suppliers = {}
    for col_idx in range(
        ITEM_MASTER_LIST_COL,
        len(sheet_values[0]) if sheet_values else 0,
        COLUMNS_PER_SUPPLIER,
    ):
        if COMPANY_NAME_ROW - 1 < len(sheet_values) and col_idx < len(
            sheet_values[COMPANY_NAME_ROW - 1]
        ):
            supplier_name = sheet_values[COMPANY_NAME_ROW - 1][col_idx].strip()
            if supplier_name:
                existing_suppliers[supplier_name] = col_idx

    next_avail_col = find_next_available_column(ws)

    products = data.get("products", [])
    if not products:
        return 0

    for product in products:
        if product.get("name"):
            product["name"] = clean_product_name(product["name"])

    company_name = data.get("company", "Unknown Company")
    col_idx = existing_suppliers.get(company_name, next_avail_col)
    if col_idx == next_avail_col:
        next_avail_col += COLUMNS_PER_SUPPLIER
    batch_requests = [
        {
            "range": f"{get_column_letter(col_idx)}{COMPANY_NAME_ROW}",
            "values": [[company_name]],
        },
        {
            "range": f"{get_column_letter(col_idx)}{CONTACT_INFO_ROW}",
            "values": [[f"{data.get('contact','')}".strip()]],
        },
        {
            "range": f"{get_column_letter(col_idx)}{HEADER_ROW}:{get_column_letter(col_idx+COLUMNS_PER_SUPPLIER-1)}{HEADER_ROW}",
            "values": [["ปริมาณ", "หน่วย", "ราคาต่อหน่วย", "รวมเป็นเงิน"]],
        },
    ]

    reference_data = [{"name": item["name"]} for item in existing_products]
    match_results = match_products_with_gemini(products, reference_data)
    matched_items = match_results["matchedItems"]
    unique_items = match_results["uniqueItems"]

    populated_rows = set()
    for item in matched_items:
        item_name = item["name"]
        for existing in existing_products:
            if existing["name"] == item_name and existing["row"] not in populated_rows:
                batch_requests.append(
                    {
                        "range": f"{get_column_letter(col_idx)}{existing['row']}:{get_column_letter(col_idx+COLUMNS_PER_SUPPLIER-1)}{existing['row']}",
                        "values": [
                            [
                                item.get("quantity", 1),
                                item.get("unit", "ชิ้น"),
                                item.get("pricePerUnit", 0),
                                item.get("totalPrice", 0),
                            ]
                        ],
                    }
                )
                populated_rows.add(existing["row"])
                break

    new_products = []
    for item in unique_items:
        if isinstance(item, dict) and "name" in item and "quantity" in item:
            item["name"] = clean_product_name(item["name"])

            if not any(
                existing["name"] == item["name"] for existing in existing_products
            ):
                new_products.append(item)

    insertion_row = (
        first_summary_row
        if first_summary_row > 0
        else (start_row + len(existing_products))
    )

    if new_products:
        new_rows = []
        for _ in range(len(new_products)):
            new_rows.append([""] * ws.col_count)

        if new_rows:
            ws.insert_rows(new_rows, insertion_row)

        row_shift = len(new_products)
        if first_summary_row > 0:
            for label in summary_row_map:
                summary_row_map[label] += row_shift

        for i, product in enumerate(new_products):
            row = insertion_row + i
            batch_requests.append(
                {
                    "range": f"{get_column_letter(ITEM_MASTER_LIST_COL)}{row}",
                    "values": [[product.get("name", "Unknown Product")]],
                }
            )

            batch_requests.append(
                {
                    "range": f"{get_column_letter(col_idx)}{row}:{get_column_letter(col_idx+COLUMNS_PER_SUPPLIER-1)}{row}",
                    "values": [
                        [
                            product.get("quantity", 1),
                            product.get("unit", "ชิ้น"),
                            product.get("pricePerUnit", 0),
                            product.get("totalPrice", 0),
                        ]
                    ],
                }
            )

    price_col = col_idx + COLUMNS_PER_SUPPLIER - 1
    summary_items = [
        ("รวมเป็นเงิน", data.get("totalPrice", 0)),
        ("ภาษีมูลค่าเพิ่ม 7%", data.get("totalVat", 0)),
        ("ยอดรวมทั้งสิ้น", data.get("totalPriceIncludeVat", 0)),
        ("กำหนดยืนราคา (วัน)", data.get("priceGuaranteeDay", "")),
        ("ระยะเวลาส่งมอบสินค้าหลังจากได้รับ PO", data.get("deliveryTime", "")),
        ("การชำระเงิน", data.get("paymentTerms", "")),
        ("อื่น ๆ", data.get("otherNotes", "")),
    ]

    if summary_row_map:
        for label, value in summary_items:
            if label in summary_row_map:
                batch_requests.append(
                    {
                        "range": f"{get_column_letter(price_col)}{summary_row_map[label]}",
                        "values": [[value]],
                    }
                )
    else:
        summary_row = insertion_row + len(new_products) + 2
        for i, (label, value) in enumerate(summary_items):
            row = summary_row + i
            batch_requests.append(
                {
                    "range": f"{get_column_letter(ITEM_MASTER_LIST_COL)}{row}",
                    "values": [[label]],
                }
            )
            batch_requests.append(
                {"range": f"{get_column_letter(price_col)}{row}", "values": [[value]]}
            )

    if batch_requests:
        ws.batch_update(batch_requests, value_input_option="USER_ENTERED")

    return 1

def update_google_sheet_with_multiple_files(ws, all_json_data):
    if len(all_json_data) == 1:
        return update_google_sheet_for_single_file(ws, all_json_data[0])

    ensure_first_three_rows_exist(ws)
    start_row = HEADER_ROW + 1
    sheet_values = ws.get_all_values()
    existing_products = []
    matched_product_rows = set()

    for row_idx, row in enumerate(sheet_values[HEADER_ROW:], start=start_row):
        if len(row) >= ITEM_MASTER_LIST_COL and row[ITEM_MASTER_LIST_COL - 1].strip():
            product_name = clean_product_name(row[ITEM_MASTER_LIST_COL - 1].strip())
            existing_products.append({"name": product_name, "row": row_idx})

    existing_suppliers = {}
    for col_idx in range(
        ITEM_MASTER_LIST_COL,
        len(sheet_values[0]) if sheet_values else 0,
        COLUMNS_PER_SUPPLIER,
    ):
        if COMPANY_NAME_ROW - 1 < len(sheet_values) and col_idx < len(
            sheet_values[COMPANY_NAME_ROW - 1]
        ):
            supplier_name = sheet_values[COMPANY_NAME_ROW - 1][col_idx].strip()
            if supplier_name:
                existing_suppliers[supplier_name] = col_idx

    next_avail_col = find_next_available_column(ws)

    for data in all_json_data:
        products = data.get("products", [])
        if not products:
            continue

        for product in products:
            if product.get("name"):
                product["name"] = clean_product_name(product["name"])

        company_name = data.get("company", "Unknown Company")

        col_idx = existing_suppliers.get(company_name, next_avail_col)
        if col_idx == next_avail_col:
            next_avail_col += COLUMNS_PER_SUPPLIER

        batch_requests = [
            {
                "range": f"{get_column_letter(col_idx)}{COMPANY_NAME_ROW}",
                "values": [[company_name]],
            },
            {
                "range": f"{get_column_letter(col_idx)}{CONTACT_INFO_ROW}",
                "values": [[f"{data.get('contact','')}".strip()]],
            },
            {
                "range": f"{get_column_letter(col_idx)}{HEADER_ROW}:{get_column_letter(col_idx+COLUMNS_PER_SUPPLIER-1)}{HEADER_ROW}",
                "values": [["ปริมาณ", "หน่วย", "ราคาต่อหน่วย", "รวมเป็นเงิน"]],
            },
        ]

        reference_data = [{"name": item["name"]} for item in existing_products]
        match_results = match_products_with_gemini(products, reference_data)
        matched_items = match_results["matchedItems"]
        unique_items = match_results["uniqueItems"]

        populated_rows = set()

        for item in matched_items:
            item_name = item["name"]
            for existing in existing_products:
                if (
                    existing["name"] == item_name
                    and existing["row"] not in populated_rows
                ):
                    batch_requests.append(
                        {
                            "range": f"{get_column_letter(col_idx)}{existing['row']}:{get_column_letter(col_idx+COLUMNS_PER_SUPPLIER-1)}{existing['row']}",
                            "values": [
                                [
                                    item.get("quantity", 1),
                                    item.get("unit", "ชิ้น"),
                                    item.get("pricePerUnit", 0),
                                    item.get("totalPrice", 0),
                                ]
                            ],
                        }
                    )
                    populated_rows.add(existing["row"])
                    break

        new_products = []
        for item in unique_items:
            if isinstance(item, dict) and "name" in item and "quantity" in item:
                item["name"] = clean_product_name(item["name"])

                if not any(
                    existing["name"] == item["name"] for existing in existing_products
                ):
                    new_products.append(item)

        next_row = start_row + len(existing_products)

        if new_products:
            new_rows = []
            for _ in range(len(new_products)):
                new_rows.append([""] * ws.col_count)

            if new_rows:
                ws.insert_rows(new_rows, next_row)

            for i, product in enumerate(new_products):
                row = next_row + i
                batch_requests.append(
                    {
                        "range": f"{get_column_letter(ITEM_MASTER_LIST_COL)}{row}",
                        "values": [[product.get("name", "Unknown Product")]],
                    }
                )

                batch_requests.append(
                    {
                        "range": f"{get_column_letter(col_idx)}{row}:{get_column_letter(col_idx+COLUMNS_PER_SUPPLIER-1)}{row}",
                        "values": [
                            [
                                product.get("quantity", 1),
                                product.get("unit", "ชิ้น"),
                                product.get("pricePerUnit", 0),
                                product.get("totalPrice", 0),
                            ]
                        ],
                    }
                )

                existing_products.append(
                    {"name": product.get("name", "Unknown Product"), "row": row}
                )

        summary_row = start_row + len(existing_products) + 2
        summary_items = [
            ("รวมเป็นเงิน", data.get("totalPrice", 0)),
            ("ภาษีมูลค่าเพิ่ม 7%", data.get("totalVat", 0)),
            ("ยอดรวมทั้งสิ้น", data.get("totalPriceIncludeVat", 0)),
            ("กำหนดยืนราคา (วัน)", data.get("priceGuaranteeDay", "")),
            ("ระยะเวลาส่งมอบสินค้าหลังจากได้รับ PO", data.get("deliveryTime", "")),
            ("การชำระเงิน", data.get("paymentTerms", "")),
            ("อื่น ๆ", data.get("otherNotes", "")),
        ]

        for i, (label, value) in enumerate(summary_items):
            row = summary_row + i
            batch_requests.append(
                {
                    "range": f"{get_column_letter(ITEM_MASTER_LIST_COL)}{row}",
                    "values": [[label]],
                }
            )
            batch_requests.append(
                {
                    "range": f"{get_column_letter(col_idx+COLUMNS_PER_SUPPLIER-1)}{row}",
                    "values": [[value]],
                }
            )

        if batch_requests:
            ws.batch_update(batch_requests, value_input_option="USER_ENTERED")

        if col_idx not in existing_suppliers.values():
            existing_suppliers[company_name] = col_idx

    return len(all_json_data)

def get_file_type(file_path):
    mime_type, _ = mimetypes.guess_type(file_path)

    if mime_type:
        if mime_type.startswith("image/"):
            return "image"
        elif mime_type == "application/pdf":
            return "pdf"
        elif mime_type in [
            "application/msword",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ]:
            return "word"

    ext = os.path.splitext(file_path)[1].lower()
    if ext in [".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".gif"]:
        return "image"
    elif ext == ".pdf":
        return "pdf"
    elif ext in [".doc", ".docx"]:
        return "word"

    return "unknown"

def process_file(file_path):
    file_type = get_file_type(file_path)
    tmp = tempfile.NamedTemporaryFile(
        delete=False, suffix=os.path.splitext(file_path)[1]
    )
    with open(file_path, "rb") as src:
        tmp.write(src.read())
    tmp.close()
    f = genai.upload_file(path=tmp.name, display_name=os.path.basename(file_path))
    if file_type == "image":
        model = genai.GenerativeModel(
            model_name="gemini-2.5-pro",
            generation_config={"temperature": 0.1, "top_p": 0.95},
        )
        resp = model.generate_content([image_prompt, f])
    else:
        model = genai.GenerativeModel(
            model_name="gemini-2.5-pro",
            generation_config={"temperature": 0.1, "top_p": 0.95},
        )
        resp = model.generate_content([prompt, f])
    d = extract_json_from_text(resp.text)
    if not d or not d.get("products"):
        model_pro = genai.GenerativeModel(
            model_name="gemini-2.5-flash",
            generation_config={"temperature": 0.1, "top_p": 0.95},
        )
        resp_pro = model_pro.generate_content(
            [prompt if file_type != "image" else image_prompt, f]
        )
        d = extract_json_from_text(resp_pro.text)
    d = validate_json_data(d) if d else None
    d = enhance_with_gemini(d) if d else None
    os.unlink(tmp.name)
    return d

def process_files(file_paths, sheet_id=DEFAULT_SHEET_ID):
    data_list = []
    for p in file_paths:
        d = process_file(p)
        if d:
            data_list.append(d)
    if data_list:
        ws = authenticate_and_open_sheet(sheet_id)
        update_google_sheet_with_multiple_files(ws, data_list)
    return data_list

def process_pdfs(pdf_paths, sheet_id=DEFAULT_SHEET_ID):
    return process_files(pdf_paths, sheet_id)

prompt = """# System Message for Product List Extraction (PDF/Text Table Processing)
## CRITICAL: ANTI-HALLUCINATION WARNING
You MUST ONLY extract information that is EXPLICITLY visible in the document. 
- DO NOT add data from anywhere other than the one in the uploaded document. Adding data that is not from the document is a serious mistake.
- DO NOT create or invent any products, prices, or specifications
- DO NOT add data from the attached Example in the prompt.
- DO NOT add products that are not clearly listed as distinct line items
- DO NOT attempt to break down a single product into multiple products
- DO NOT interpret descriptive text as separate products
- If uncertain about any information, LEAVE IT OUT rather than guessing

## CRITICAL: CONTACT INFORMATION EXTRACTION
Pay special attention to contact information in the document header or footer:
- Extract any email addresses (example@domain.com)
- Extract any phone numbers (formats like 02-3384825, 081-1234567, 095-525-2623)
- Put email and phone details in the "contact" field
- Format: "Email: email@example.com, Phone: 081-234-5678"
- Use only the phone number and email address on the letterhead. Do not add "บริษัท ลูก้า แอสเซท จำกัด, 081-781-7283" contact information

## CRITICAL: COMPLETE EXTRACTION REQUIREMENT
You MUST extract ALL products visible in the document:
- Extract EVERY product line item visible in the document
- Preserve hierarchical structure (groups/categories) of products if present
- Ensure NO products are missed or skipped
- Each row with a distinct price is one product

## Input Format
Provide PDF files (or images/tables with text extraction) containing product information (receipts, invoices, product lists, etc.).

## Task
Extract ALL product information EXPLICITLY visible in the document:
- Extract products with quantities, units, and prices
- Preserve parent-child relationships (main categories and sub-items)
- Maintain hierarchical product groupings (numbered sections, categories)
- Also extract additional quotation details like price validity, delivery time, payment terms, etc.

## Hierarchical Structure Handling
Many quotations organize products hierarchically. When you see this:
- Include category names in product descriptions (e.g., "งานบันไดกระจก งานพื้นตก - กระจกเทมเปอร์ใส หนา 10 มม. ขนาด 4.672×0.97 ม.")
- If product descriptions begin with numbers (1, 2, 3...), REMOVE those numbers
- Include all parent category information in each product's name without the leading numbers

## Output Format (JSON only)
You must return ONLY this JSON structure:
{
  "company": "company name or first name + last name (NEVER null)",
  "vat": true,
  "name": "customer name or null",
  "contact": "phone number or email or null",
  "priceGuaranteeDay": 30
  "deliveryTime": "",
  "paymentTerms": "",
  "otherNotes": "",
  "products": [
    {
      "name": "full product description including ALL parent category info, specifications AND dimensions WITHOUT leading numbers",
      "quantity": 1,
      "unit": "match the unit shown in the document (e.g., แผ่น, ตร.ม., ชิ้น, ตัว, เมตร, ชุด)",
      "pricePerUnit": 0,
      "totalPrice": 0
    }
  ],
  "totalPrice": 0,
  "totalVat": 0,
  "totalPriceIncludeVat": 0
}

## Example 1 (Format with Item/ART.No./Description/Qty/Unit/Price columns):
| Item | ART.No. | Description | Qty | Unit | Standard Price | Discount Price | Amount |
|------|---------|-------------|-----|------|----------------|---------------|--------|
| 1    | CPW-xxxx| SPC ลายไม้ 4.5 มิล (ก้างปลา) | 1.00 | ตร.ม. |  | 520.00 | 520.000 |
| 2    |         | ค่าแรงติดตั้ง | 1.00 | ตร.ม. |  | 150.00 | 150.000 |

## Example 2 (Format with ลำดับ/รหัสสินค้า/รายละเอียดสินค้า columns):
| ลำดับ | รหัสสินค้า | รายละเอียดสินค้า | หน่วย | จำนวน | ราคา/หน่วย(บาท) | จำนวนเงิน(บาท) |
|------|---------|----------------|------|------|--------------|------------|
| 1    |         | พื้นไม้ไวนิลลายไม้ปลา 4.5 มม. LKT 4.5 mm x 0.3 mm สีฟ้าเซอร์คูลี (1 กล่อง บรรจุ 18 แผ่น หรือ 1.3 ตร.ม) | ตร.ม. | 1.30 | 680.00 | 884.00 |

## Field Extraction Guidelines

### name (Product Description)
* CRITICAL: Include ALL hierarchical information in each product name:
  - Category names/headings (e.g., "งานบันไดกระจก งานพื้นตก")
  - Sub-category information (e.g., "เหล็กตัวซีชุบสังกะสี")
  - Glass type, thickness (e.g., "กระจกเทมเปอร์ใส หนา 10 มม.")
  - Exact dimensions (e.g., "ขนาด 4.672×0.97 ม.")
* REMOVE any leading numbers (1., 2., 3.) from the product descriptions
* Format hierarchical products as: "[Category Name] - [Material] - [Type] - [Dimensions]"
* Include: ALL distinguishing characteristics that make each product unique
* Example: "งานบันไดกระจก งานพื้นตก - เหล็กตัวซีชุบสังกะสี ไม่รวมปูน - กระจกเทมเปอร์ใส หนา 10 มม. ขนาด 4.672×0.97 ม."

### unit and quantity (DIRECT EXTRACTION RULE)
* Extract unit and quantity DIRECTLY from each line item as shown
* Use the exact unit shown in the document (ชุด, แผ่น, ตร.ม., ชิ้น, ตัว, เมตร, etc.)
* Extract the exact quantity shown for each product (never assume or calculate)
* NEVER create quantities or units that aren't explicitly shown in the document
* Pay special attention to decimal quantities - extract the full decimal precision

### pricePerUnit and totalPrice
* Extract ONLY prices clearly visible in the document
* Use numeric values only (no currency symbols)
* Extract cleanly from pricing fields as shown in each line item
* NEVER calculate or estimate prices that aren't explicitly shown
* NEVER combine different products' prices
* Pay special attention to decimal prices - extract the EXACT decimal values shown

### Additional Quotation Details
* Extract these additional fields if present:
  - "กำหนดยืนราคา (วัน)", "กำหนดยืนราคา", "การยืนราคา" - Price validity period in days (priceGuaranteeDay)
  - "ระยะเวลาส่งมอบสินค้าหลังจากได้รับ PO" - Delivery time after PO (deliveryTime)
  - "การชำระเงิน" - Payment terms (paymentTerms)
  - "อื่น ๆ" - Other notes (otherNotes)
* Extract as text exactly as written, preserving numbers and Thai language

### CRITICAL: Pricing summaries and summary values
* Extract the exact values for these three summary items:
  - "รวมเป็นเงิน" - the initial subtotal (totalPrice)
  - "ภาษีมูลค่าเพิ่ม 7%" - the VAT amount (totalVat)
  - "ยอดรวมทั้งสิ้น" - the final total (totalPriceIncludeVat)
* Alternative labels to match:
  - For totalPrice: "รวม", "รวมเป็นเงิน", "ราคารวม", "Total", "TOTAL AMOUNT", "รวมราคา"
  - For totalVat: "ภาษีมูลค่าเพิ่ม 7%", "VAT 7%"
  - For totalPriceIncludeVat: "ยอดรวมทั้งสิ้น", "รวมทั้งหมด", "รวมเงินทั้งสิน", "ราคารวมสุทธิ", "รวมราคางานทั้งหมดตามสัญญา"
* Extract the exact values as shown (remove commas, currency symbols)
* CRITICAL: Preserve full decimal precision in all monetary values

## FINAL VERIFICATION
Review the extracted products one last time and verify:
1. Count the number of products you've extracted
2. Verify this matches EXACTLY with the number of product rows visible in the document
3. Check that ALL products have proper hierarchical information included WITHOUT leading numbers
4. Ensure NO products are missing - every line item with a price must be extracted
5. Confirm all dimensions and specifications are preserved correctly
6. Verify all decimal values (quantities and prices) maintain their full precision
"""

image_prompt = """# System Message for Product List Extraction from Images
## CRITICAL: ANTI-HALLUCINATION WARNING
You MUST ONLY extract information that is EXPLICITLY visible in the image. 
- DO NOT create or invent any products, prices, or specifications
- DO NOT add products that are not clearly listed as distinct line items
- DO NOT interpret descriptive text as separate products
- If text is unclear or unreadable, mark it as uncertain rather than guessing

## CRITICAL: COMPLETE EXTRACTION REQUIREMENT
You MUST extract ALL products visible in the image:
- Extract EVERY product line item visible in the image
- Preserve hierarchical structure (groups/categories) of products if present
- Ensure NO products are missed or skipped
- Each row with a distinct price is one product

## Input Format
I'm providing an image of a document containing product information.

## Task
Extract ALL product information EXPLICITLY visible in the image:
- Extract products with quantities, units, and prices
- Preserve parent-child relationships (main categories and sub-items)
- Maintain hierarchical product groupings (numbered sections, categories)
- Also extract additional quotation details like price validity, delivery time, payment terms, etc.

## Hierarchical Structure Handling
Many quotations organize products hierarchically. When you see this:
- Include category names in product descriptions (e.g., "งานบันไดกระจก งานพื้นตก - กระจกเทมเปอร์ใส หนา 10 มม. ขนาด 4.672×0.97 ม.")
- If product descriptions begin with numbers (1, 2, 3...), REMOVE those numbers
- Include all parent category information in each product's name without the leading numbers

## Output Format (JSON only)
You must return ONLY this JSON structure:
{
  "company": "company name or first name + last name (NEVER null)",
  "vat": true,
  "contact": "phone number or email or null",
  "priceGuaranteeDay": 30
  "deliveryTime": "",
  "paymentTerms": "",
  "otherNotes": "",
  "products": [
    {
      "name": "full product description including ALL parent category info, specifications AND dimensions WITHOUT leading numbers",
      "quantity": 1,
      "unit": "match the unit shown in the document (e.g., แผ่น, ตร.ม., ชิ้น, ตัว, เมตร, ชุด)",
      "pricePerUnit": 0,
      "totalPrice": 0
    }
  ],
  "totalPrice": 0,
  "totalVat": 0,
  "totalPriceIncludeVat": 0
}


## Example 1 (Format with Item/ART.No./Description/Qty/Unit/Price columns):
| Item | ART.No. | Description | Qty | Unit | Standard Price | Discount Price | Amount |
|------|---------|-------------|-----|------|----------------|---------------|--------|
| 1    | CPW-xxxx| SPC ลายไม้ 4.5 มิล (ก้างปลา) | 1.00 | ตร.ม. |  | 520.00 | 520.000 |
| 2    |         | ค่าแรงติดตั้ง | 1.00 | ตร.ม. |  | 150.00 | 150.000 |

## Example 2 (Format with ลำดับ/รหัสสินค้า/รายละเอียดสินค้า columns):
| ลำดับ | รหัสสินค้า | รายละเอียดสินค้า | หน่วย | จำนวน | ราคา/หน่วย(บาท) | จำนวนเงิน(บาท) |
|------|---------|----------------|------|------|--------------|------------|
| 1    |         | พื้นไม้ไวนิลลายไม้ปลา 4.5 มม. LKT 4.5 mm x 0.3 mm สีฟ้าเซอร์คูลี (1 กล่อง บรรจุ 18 แผ่น หรือ 1.3 ตร.ม) | ตร.ม. | 1.30 | 680.00 | 884.00 |


## Field Extraction Guidelines

### name (Product Description)
* CRITICAL: Include ALL hierarchical information in each product name:
  - Category names/headings (e.g., "งานบันไดกระจก งานพื้นตก")
  - Sub-category information (e.g., "เหล็กตัวซีชุบสังกะสี")
  - Glass type, thickness (e.g., "กระจกเทมเปอร์ใส หนา 10 มม.")
  - Exact dimensions (e.g., "ขนาด 4.672×0.97 ม.")
* REMOVE any leading numbers (1., 2., 3.) from the product descriptions
* Format hierarchical products as: "[Category Name] - [Material] - [Type] - [Dimensions]"
* Include: ALL distinguishing characteristics that make each product unique
* Example: "งานบันไดกระจก งานพื้นตก - เหล็กตัวซีชุบสังกะสี ไม่รวมปูน - กระจกเทมเปอร์ใส หนา 10 มม. ขนาด 4.672×0.97 ม."

### unit and quantity (DIRECT EXTRACTION RULE)
* Extract unit and quantity DIRECTLY from each line item as shown
* Use the exact unit shown in the document (ชุด, แผ่น, ตร.ม., ชิ้น, ตัว, เมตร, จำนวนต่อชุด etc.)
* Extract the exact quantity shown for each product (never assume or calculate)
* NEVER create quantities or units that aren't explicitly shown in the document
* Pay special attention to decimal quantities - extract the full decimal precision

### pricePerUnit and totalPrice
* Extract ONLY prices clearly visible in the document
* Use numeric values only (no currency symbols)
* Extract cleanly from pricing fields as shown in each line item
* NEVER calculate or estimate prices that aren't explicitly shown
* NEVER combine different products' prices
* Pay special attention to decimal prices - extract the EXACT decimal values shown

### Additional Quotation Details
* Extract these additional fields if present:
  - "กำหนดยืนราคา (วัน)", "กำหนดยืนราคา", "การยืนราคา" - Price validity period in days (priceGuaranteeDay)
  - "ระยะเวลาส่งมอบสินค้าหลังจากได้รับ PO" - Delivery time after PO (deliveryTime)
  - "การชำระเงิน" - Payment terms (paymentTerms)
  - "อื่น ๆ" - Other notes (otherNotes)
* Extract as text exactly as written, preserving numbers and Thai language

### CRITICAL: Pricing summaries and summary values
* Extract the exact values for these three summary items:
  - "รวมเป็นเงิน" - the initial subtotal (totalPrice)
  - "ภาษีมูลค่าเพิ่ม 7%" - the VAT amount (totalVat)
  - "ยอดรวมทั้งสิ้น" - the final total (totalPriceIncludeVat)
* Alternative labels to match:
  - For totalPrice: "รวม", "รวมเป็นเงิน", "ราคารวม", "Total", "TOTAL AMOUNT", "รวมราคา"
  - For totalVat: "ภาษีมูลค่าเพิ่ม 7%", "VAT 7%"
  - For totalPriceIncludeVat: "ยอดรวมทั้งสิ้น", "รวมทั้งหมด", "รวมเงินทั้งสิน", "ราคารวมสุทธิ", "รวมราคางานทั้งหมดตามสัญญา"
* Extract the exact values as shown (remove commas, currency symbols)
* CRITICAL: Preserve full decimal precision in all monetary values

## FINAL VERIFICATION
Review the extracted products one last time and verify:
1. Count the number of products you've extracted
2. Verify this matches EXACTLY with the number of product rows visible in the image
3. Check that ALL products have proper hierarchical information included WITHOUT leading numbers
4. Ensure NO products are missing - every line item with a price must be extracted
5. Confirm all dimensions and specifications are preserved correctly
6. Verify all decimal values (quantities and prices) maintain their full precision
"""

validation_prompt = """
You are a data validation expert specializing in Thai construction quotations.
I've extracted product data from a document, but there may be missing products or hierarchical relationships.

## CRITICAL: COMPLETE DATA CHECK
Your primary task is to ensure ALL products are correctly extracted with their hierarchical structure:
1. Check that all products visible in the document have been extracted
2. Ensure parent-child relationships and category groupings are preserved
3. Verify that all products have complete descriptions including their category name
4. Make sure no products are missing dimensions or specifications

## CRITICAL: PRESERVE PRODUCT HIERARCHY
Thai construction quotations often organize products hierarchically by categories:
- Category names with descriptive details
- Materials and specifications
- Dimensions

Each product must include its complete hierarchy:
"[Category Name] - [Material] - [Type] - [Dimensions]"

Examples (DO NOT add data from the attached Example in the prompt): 
- "งานบันไดกระจก งานพื้นตก - เหล็กตัวซีชุบสังกะสี ไม่รวมปูน - กระจกเทมเปอร์ใส หนา 10 มม. ขนาด 4.672×0.97 ม."
- "งานพื้นตก (ชั้นลอย) - เหล็กตัวซีชุบสังกะสี ไม่รวมปูน - เทมเปอร์ใส หนา 10 มม. ขนาด 3.565×0.97 ม."

## CRITICAL: DECIMAL NUMBER ACCURACY
Pay special attention to:
1. Quantities with decimals (extract full precision)
2. Dimensions with decimals (preserve exact measurements)
3. Prices with decimals (maintain exact values)

## CRITICAL: CLEAN PRODUCT DESCRIPTIONS
1. REMOVE any leading numbers (1., 2., 3., etc.) from product descriptions
2. Ensure NO product descriptions begin with numbering
3. Maintain all other hierarchical information and details

Review the data carefully and FIX these issues:
1. ADD any missing products that should be extracted from the source document
2. FIX product names to include complete hierarchical information WITHOUT leading numbers
3. ENSURE all dimensions and specifications are preserved with full decimal precision
4. VERIFY every product has the correct quantity, unit, price and total with full decimal precision

Original extraction:
{extracted_json}

Return ONLY a valid JSON object with no explanations.
"""

matching_prompt = """
You are a product matching expert specializing in construction materials in Thailand.

## TASK
Match products from a target list (new quotation) to a reference list (master product list) based on their attributes, materials, dimensions, and specifications.

## CRITICAL MATCHING RULES
1. Focus on the meaning and specifications, not just text similarity
2. Consider materials, dimensions, thickness, and product type as key matching factors
3. Each reference product can only be matched ONCE (never match the same reference item to multiple target items)
4. If a product cannot be matched with high confidence (>70%), leave it unmatched

## UNIT CONVERSION AWARENESS
Pay special attention to dimensions and units:
- Convert between mm, cm, and m when comparing dimensions (1m = 100cm = 1000mm)
- Match items with similar dimensions even if units differ (e.g., "4672x970 mm" and "4.672x0.97 m" are the same)
- Consider products like "ราวกันตกฝังปูน" with similar specifications as potential matches even if dimensions vary slightly

## INPUT
- Target Products: New products from a quotation that need to be matched
- Reference Products: Existing master list of products to match against

## MATCHING CRITERIA (in priority order)
1. Material type match (e.g., glass with glass, steel with steel)
2. Dimensions match (within 5% tolerance, after unit conversion)
3. Thickness match (within 5% tolerance, after unit conversion)
4. Product type/category match

## Material Type Examples
- Glass: กระจก, glass, tempered, เทมเปอร์
- Steel: เหล็ก, steel, galvanized, ชุบสังกะสี
- Aluminum: อลูมิเนียม, aluminum, aluminium
- Wood: ไม้, wood, timber, plywood

## OUTPUT FORMAT
Return a JSON object with these properties:
{{
  "matchedItems": [
    {{
      "name": "reference product name",
      "quantity": target quantity,
      "unit": target unit,
      "pricePerUnit": target price per unit,
      "totalPrice": target total price
    }}
  ],
  "uniqueItems": [
    {{
      "name": "target product name",
      "quantity": target quantity,
      "unit": target unit,
      "pricePerUnit": target price per unit,
      "totalPrice": target total price
    }}
  ]
}}

Where:
- matchedItems: Array of products that found a match in the reference list
  - Use the reference product name, but the target's quantity, unit, and prices
- uniqueItems: Array of target products that couldn't be matched plus any unused reference products

## Target Products:
{target_products}

## Reference Products:
{reference_products}
"""

def main():
    st.set_page_config(page_title="ระบบประมวลผลใบเสนอราคา", layout="centered")
    st.sidebar.title("เพิ่ม API Key เพื่อเริ่มต้นใช้งาน")
    google_api_key = st.sidebar.text_input(
        "Enter your GOOGLE_API_KEY", 
        value=st.session_state.get("google_api_key", ""), 
        type="password",
        key="google_api_key_input"
    )

    if st.sidebar.button("ยืนยัน", key="confirm_api_key", use_container_width=True):
        if google_api_key:
            try:
                genai.configure(api_key=google_api_key)
                st.session_state.google_api_key = google_api_key
                st.session_state.api_key_confirmed = True
                st.sidebar.success("API Key ถูกบันทึกแล้ว")
            except Exception as e:
                st.session_state.api_key_confirmed = False
                st.sidebar.error(f"เกิดข้อผิดพลาด: {e}")
        else:
            st.session_state.api_key_confirmed = False
            st.sidebar.error("กรุณาใส่ API Key ก่อนยืนยัน")

    st.markdown("<h1 style='text-align: center;'>ระบบประมวลผลใบเสนอราคา</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center;'>อัปโหลดไฟล์ใบเสนอราคา (PDF หรือรูปภาพ) เพื่อประมวลผลและส่งข้อมูลเข้า Google Sheet</p>", unsafe_allow_html=True)

    progress_bar = st.empty()
    progress_bar.markdown(
        """
        <div style="display:flex; justify-content:space-around; align-items:center; margin-bottom: 2rem;">
            <div><div style="background-color:#4285F4;color:white;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">1</div><p style="text-align:center;margin-top:5px;">เลือกไฟล์</p></div>
            <div><div style="background-color:#E8E8E8;color:#666;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">2</div><p style="text-align:center;margin-top:5px;color:#666;">ประมวลผล</p></div>
            <div><div style="background-color:#E8E8E8;color:#666;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">3</div><p style="text-align:center;margin-top:5px;color:#666;">ผลการประมวลผล</p></div>
        </div>
        """, unsafe_allow_html=True
    )

    st.subheader("อัปโหลดข้อมูล")
    sheet_url = st.text_input(
        "Google Sheet URL or ID:",
        value=DEFAULT_SHEET_ID,
        placeholder="ใส่ URL หรือ ID ของ Google Sheet"
    )
    
    uploaded_files = st.file_uploader(
        "เลือกไฟล์ PDF หรือรูปภาพ (สามารถเลือกได้หลายไฟล์)",
        type=['pdf', 'jpg', 'jpeg', 'png'],
        accept_multiple_files=True
    )
    
    if uploaded_files:
        if st.button("🚀 เริ่มประมวลผล", use_container_width=True):
            if not st.session_state.get("api_key_confirmed"):
                st.error("❌ กรุณาใส่และยืนยัน Google API Key ในแถบด้านข้างก่อนเริ่มประมวลผล")
                return

            progress_bar.markdown(
                """
                <div style="display:flex; justify-content:space-around; align-items:center; margin-bottom: 2rem;">
                    <div><div style="background-color:#4285F4;color:white;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">✓</div><p style="text-align:center;margin-top:5px;">เลือกไฟล์</p></div>
                    <div><div style="background-color:#4285F4;color:white;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">2</div><p style="text-align:center;margin-top:5px;">ประมวลผล</p></div>
                    <div><div style="background-color:#E8E8E8;color:#666;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">3</div><p style="text-align:center;margin-top:5px;color:#666;">ผลการประมวลผล</p></div>
                </div>
                """, unsafe_allow_html=True
            )

            with st.spinner('กำลังประมวลผลไฟล์...'):
                file_paths = []
                for uploaded_file in uploaded_files:
                    with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
                        tmp_file.write(uploaded_file.getvalue())
                        file_paths.append(tmp_file.name)
                
                sheet_id = extract_sheet_id_from_url(sheet_url) if sheet_url else DEFAULT_SHEET_ID
                results = process_files(file_paths, sheet_id)
                
                for path in file_paths:
                    os.unlink(path)

            progress_bar.markdown(
                """
                <div style="display:flex; justify-content:space-around; align-items:center; margin-bottom: 2rem;">
                    <div><div style="background-color:#4285F4;color:white;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">✓</div><p style="text-align:center;margin-top:5px;">เลือกไฟล์</p></div>
                    <div><div style="background-color:#4285F4;color:white;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">✓</div><p style="text-align:center;margin-top:5px;">ประมวลผล</p></div>
                    <div><div style="background-color:#4285F4;color:white;border-radius:50%;width:40px;height:40px;text-align:center;line-height:40px;margin:0 auto;">3</div><p style="text-align:center;margin-top:5px;">ผลการประมวลผล</p></div>
                </div>
                """, unsafe_allow_html=True
            )
            
            st.subheader("ผลการประมวลผล")
            if results:
                st.success(f"✅ ประมวลผลสำเร็จ! ประมวลผลได้ {len(results)} ไฟล์ และบันทึกข้อมูลลง Google Sheet เรียบร้อยแล้ว")
                sheet_url_display = f"https://docs.google.com/spreadsheets/d/{sheet_id}"
                st.markdown(f"### [เปิด Google Sheet]({sheet_url_display})")
                
                for i, result in enumerate(results):
                    with st.expander(f"📄 ไฟล์ที่ {i+1}: {result.get('company', 'Unknown Company')}"):
                        st.json(result)
            else:
                st.error("❌ ไม่สามารถประมวลผลไฟล์ได้ กรุณาตรวจสอบไฟล์และลองใหม่อีกครั้ง")

if __name__ == "__main__":
    main()