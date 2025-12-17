from fastapi.responses import FileResponse
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse
import os
from brochure import generate_brochure
from createpowerpoint import extract_event, check_if_processed,send_final_email ,extract_form_data , extract_email ,update_form_with_db ,resolve_project_name ,extract_project_type ,generate_main_ppt ,mark_item_as_processed
from space_images import download_space_images
from summary.webinerbrief import process_webinerbrief
from summary.studentjourney import process_student_journey

app = FastAPI()

FILES_DIR = "files"
os.makedirs(FILES_DIR, exist_ok=True)
BASE_URL = os.getenv("BASE_URL")


@app.post("/monday-webhook")
async def monday_webhook(request: Request):
    body = await request.json()

    # 1️⃣ Extract Monday event
    event, is_challenge = extract_event(body)
    if is_challenge:
        return JSONResponse(content=event)
    if not event:
        return {"status": "ok", "message": "Webhook received but no event data"}

    item_id = event.get("pulseId")
    print(f"🚀 Processing item {item_id}")

    # 2️⃣ Prevent reprocessing of the same item
    already_done, response = check_if_processed(event)
    if already_done:
        return response

    # 3️⃣ Extract form data & selected styles
    form_data, col_vals, selected_styles = extract_form_data(event)

    # 4️⃣ Extract email
    email = extract_email(col_vals, form_data)

    # 5️⃣ Enhance form_data with DB lookup
    form_data, db_details = update_form_with_db(form_data, email)

    # 6️⃣ Resolve final project name
    form_data = resolve_project_name(form_data, col_vals, event)

    # 7️⃣ Extract project type (HOTEL / RESIDENTIAL / F&B)
    project_type = extract_project_type(event).lower()
    print(f"🏷 Project Type: {project_type}")

    results = {}

    # 8️⃣ Generate MAIN PPT
    OUTPUT_PATH = generate_main_ppt(
        event, item_id, selected_styles, form_data, project_type)
    results["full"] = {
        "output_file": OUTPUT_PATH,
        "ppt_type": "full",
        "email": email,
        "project_name": form_data.get("Q. Project Name"),
        "project_type": project_type
    }

    # 9️⃣ Generate Brochure PPT (project-type independent)
    all_pictures = download_space_images(event, item_id)
    brochure_path = generate_brochure(
        item_id, selected_styles, form_data, all_pictures,event )

    results["brochure"] = {
        "output_file": brochure_path,
        "ppt_type": "brochure",
        "email": email,
        "project_name": form_data.get("Q. Project Name"),
        "project_type": project_type
    }

    #  🔟 Mark item as processed
    mark_item_as_processed(item_id)
    
    print("send email")
    send_final_email(item_id,form_data,email)
    print(f"🎉 Successfully processed item {item_id}")

    return {
        "status": "processed",
        "results": results,
        "item_id": item_id,
        "project_type": project_type
    }



@app.get("/download-ppt")
async def download_ppt(item_id: int, ppt_type: str = "output"):
    """
    Download a generated PPT by item_id.
    ppt_type can be 'output' or 'brochure'.
    """
    if ppt_type == "output":
        file_name = f"{item_id}_output.pptx"
    elif ppt_type == "brochure":
        file_name = f"{item_id}_boutput.pptx"
    else:
        return {"error": "Invalid ppt_type. Use 'output' or 'brochure'."}

    file_path = os.path.join(FILES_DIR, file_name)

    if os.path.exists(file_path):
        return FileResponse(
            file_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=file_name
        )
    return {"error": f"File not found for item_id {item_id} and type {ppt_type}"}



@app.post("/webinerbrief")
async def monday_webinerbrief(request: Request):
    data = await request.json()
    return await process_webinerbrief(data)


@app.post("/monday-webhook")
async def monday_student_journey(request: Request):
    data = await request.json()
    return await process_student_journey(data)

@app.get("/hello")
async def hello():
    print("hello for the power point creation")
    return "working"