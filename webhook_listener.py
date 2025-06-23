from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from create_vectordb import main as full_refresh
import threading, time

app = FastAPI()

lock = threading.Lock()
is_cooldown = False
cooldown_timer = None

def reset_cooldown():
    global is_cooldown
    print("🔁 Cooldown period ended.")
    is_cooldown = False

def run_main_with_lock():
    global is_cooldown, cooldown_timer

    if lock.locked():
        print("⏳ Already processing. Skipping...")
        return

    with lock:
        print("🚀 Starting full_refresh()")
        start = time.time()
        full_refresh()
        end = time.time()
        duration = end - start

        print(f"✅ main() completed in {duration:.2f} seconds")

        # Set cooldown for same duration (or a safe buffer like +5 sec)
        is_cooldown = True
        cooldown_timer = threading.Timer(duration + 5, reset_cooldown)
        cooldown_timer.start()

@app.api_route("/webhook", methods=["GET", "POST"])
async def webhook(request: Request):
    validation_token = request.query_params.get("validationToken")
    if validation_token:
        print(f"🔐 Validation request received: {validation_token}")
        return PlainTextResponse(content=validation_token, status_code=200)

    try:
        data = await request.json()
        print("📩 Webhook notification received:", data)

        if not is_cooldown:
            print("⚡ Debounce clear — running full_refresh()")
            threading.Thread(target=run_main_with_lock).start()
        else:
            print("🛑 Cooldown active — skipping execution")

    except Exception as e:
        print("❌ Failed to handle webhook:", e)

    return {"status": "received"}
