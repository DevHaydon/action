import os
import sys
import json

# Minimal in-memory key-value store for demonstration
memory = {}

MEMORY_NAME = os.environ.get("MEMORY_NAME", "default")

print(f"[memory_server_inmemory] Starting in-memory memory server: {MEMORY_NAME}")

# Simple REPL for demonstration (replace with your protocol as needed)
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
                memory[key] = value
                print(json.dumps({"status": "ok"}))
            elif action == "get":
                print(json.dumps({"value": memory.get(key)}))
            elif action == "clear":
                memory.clear()
                print(json.dumps({"status": "cleared"}))
            else:
                print(json.dumps({"error": "unknown action"}))
        except Exception as e:
            print(json.dumps({"error": str(e)}))
        sys.stdout.flush()
except KeyboardInterrupt:
    print("[memory_server_inmemory] Shutting down.")
