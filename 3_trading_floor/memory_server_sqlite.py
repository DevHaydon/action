import os
import sys
import json
import sqlite3

MEMORY_NAME = os.environ.get("MEMORY_NAME", "default")
DB_DIR = os.path.join(os.path.dirname(__file__), "memory")
os.makedirs(DB_DIR, exist_ok=True)
DB_PATH = os.path.join(DB_DIR, f"{MEMORY_NAME}.db")

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()
cursor.execute("CREATE TABLE IF NOT EXISTS memory (key TEXT PRIMARY KEY, value TEXT)")
conn.commit()

print(f"[memory_server_sqlite] Starting SQLite memory server: {DB_PATH}")

try:
    while True:
        line = sys.stdin.readline()
        if not line:
            break
        try:
            cmd = json.loads(line)
            action = cmd.get("action")
            key = cmd.get("key")
            value = cmd.get("value")
            if action == "set":
                cursor.execute("REPLACE INTO memory(key, value) VALUES (?, ?)", (key, json.dumps(value)))
                conn.commit()
                print(json.dumps({"status": "ok"}))
            elif action == "get":
                cursor.execute("SELECT value FROM memory WHERE key = ?", (key,))
                row = cursor.fetchone()
                print(json.dumps({"value": json.loads(row[0]) if row else None}))
            elif action == "clear":
                cursor.execute("DELETE FROM memory")
                conn.commit()
                print(json.dumps({"status": "cleared"}))
            else:
                print(json.dumps({"error": "unknown action"}))
        except Exception as e:
            print(json.dumps({"error": str(e)}))
        sys.stdout.flush()
except KeyboardInterrupt:
    pass
finally:
    conn.close()
    print("[memory_server_sqlite] Shutting down.")
