from fastapi import Request
import json
import requests
import google.generativeai as genai
from google.generativeai import GenerativeModel
import os

# Load keys from environment instead of hardcoding
genai.configure(api_key=os.getenv("GEMINI_API_KEY"))
MONDAY_API_KEY = os.getenv("MONDAY_API_KEY")

MONDAY_API_URL = "https://api.monday.com/v2"
WORKSPACE_ID = 1549158

model = genai.GenerativeModel("gemini-2.5-flash")


# ============================================================
# Generate Student Journey
# ============================================================
def generate_student_journey(project, comments):
    prompt = f"""
You are an AI mentor for interior design students.

Based on the real project update comments below, generate a clear, short, and professional student journey summary.

Write the journey as:
- Concept & inspiration stage  
- Moodboard & style direction  
- Space planning / layout decisions  
- 3D modeling & material selection  
- Feedback revisions  
- Final outcome  

Project Name: {project}

Project Comments:
{comments}

Return ONLY the student journey summary.
"""

    try:
        response = model.generate_content(prompt)

        if hasattr(response, "text") and response.text:
            return response.text

        return "⚠️ Unexpected response format"

    except Exception as e:
        print("Gemini Error:", e)
        return "Error generating journey"


# ============================================================
# Monday API Helpers
# ============================================================
def monday_api(query, variables=None):
    headers = {
        "Authorization": MONDAY_API_KEY,
        "Content-Type": "application/json"
    }
    
    response = requests.post(
        MONDAY_API_URL,
        json={"query": query, "variables": variables or {}},
        headers=headers
    )

    return response.json()


def get_item_name(board_id, item_id):
    query = """
    query ($board_id: ID!, $item_id: [ID!]!) {
      boards(ids: [$board_id]) {
        items_page(limit: 1, query_params: {ids: $item_id}) {
          items { id name }
        }
      }
    }
    """
    variables = {"board_id": str(board_id), "item_id": [str(item_id)]}
    data = monday_api(query, variables)

    try:
        return data["data"]["boards"][0]["items_page"]["items"][0]["name"]
    except:
        return None


def find_board_by_item_name(workspace_id, item_name):
    query = """
    query ($workspace_id: [ID!]) {
      boards (workspace_ids: $workspace_id, limit: 200) {
        id name
      }
    }
    """
    variables = {"workspace_id": [str(workspace_id)]}
    data = monday_api(query, variables)

    for board in data["data"]["boards"]:
        if item_name.lower().strip() == board["name"].lower().strip():
            return board

    return None


def get_board_details(board_id):
    query = """
    query ($board_id: [ID!]) {
      boards(ids: $board_id) {
        id name
        items_page(limit: 200) {
          items {
            id name
            
            updates(page: 1, limit: 100) {
              id text_body created_at
              creator { name email }
              replies { id text_body created_at creator { name email } }
            }

            subitems {
              id name
              updates(page: 1, limit: 100) {
                id text_body created_at
                creator { name email }
                replies { id text_body created_at creator { name email } }
              }
            }
          }
        }
      }
    }
    """
    variables = {"board_id": [str(board_id)]}
    data = monday_api(query, variables)
    board = data["data"]["boards"][0]

    all_comments = {}

    for item in board["items_page"]["items"]:
        item_name = item["name"]
        all_comments[item_name] = []

        # Item updates
        for update in item.get("updates", []):
            all_comments[item_name].append(update.get("text_body"))

            for reply in update.get("replies", []):
                all_comments[item_name].append(reply.get("text_body"))

        # Subitem updates
        for sub in item.get("subitems", []):
            for u in sub.get("updates", []):
                all_comments[item_name].append(u.get("text_body"))

                for r in u.get("replies", []):
                    all_comments[item_name].append(r.get("text_body"))

    return all_comments


def add_comment(item_id, text):
    mutation = """
    mutation ($item_id: ID!, $text: String!) {
      create_update (item_id: $item_id, body: $text) { id body }
    }
    """
    monday_api(mutation, {"item_id": str(item_id), "text": text})


def format_for_monday(text: str):
    return text.replace("\n", "<br>")[:14000]


# ============================================================
# MAIN PROCESS FUNCTION (Called from main.py)
# ============================================================
async def process_student_journey(data):
    print("📩 Incoming Monday Webhook:", json.dumps(data, indent=2))

    event = data.get("event", {})

    board_id = event.get("boardId")
    item_id = event.get("pulseId")

    if not board_id or not item_id:
        return {"status": "error", "msg": "Missing board or item ID"}

    # 1️⃣ Get item name
    item_name = get_item_name(board_id, item_id)
    if not item_name:
        return {"status": "error", "msg": "Item name not found"}

    # 2️⃣ Find matching student board
    matching_board = find_board_by_item_name(WORKSPACE_ID, item_name)
    if not matching_board:
        return {"status": "success", "msg": "No matching student board"}

    student_board_id = matching_board["id"]

    # 3️⃣ Fetch all comments
    comments = get_board_details(student_board_id)

    full_journey_log = ""

    # 4️⃣ Generate journeys for each project
    for project, project_comments in comments.items():
        journey = generate_student_journey(project, project_comments)
        full_journey_log += f"\n\n<b>{project}</b><br>{journey}"

    formatted = format_for_monday(full_journey_log)

    # 5️⃣ Add ONE combined summary back to Monday
    add_comment(
        item_id,
        f"🤖✨ <b>AI Journey Summary</b><br><br>{formatted}"
    )

    return {"status": "success"}
