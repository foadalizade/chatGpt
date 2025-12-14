import json, os

def load_settings():
    if not os.path.exists("settings.json"):
        with open("settings.json", "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=4)
        return {}

    with open("settings.json", "r", encoding="utf-8") as f:
        return json.load(f)
