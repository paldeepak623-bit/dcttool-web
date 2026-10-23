# -*- coding: utf-8 -*-
import subprocess
import time
import re
import sys
import os
import webbrowser
from pathlib import Path

# Safe encoding for Windows console
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE_DIR = Path(__file__).resolve().parent

print("================================================================")
print("  DCTtool V2.0 Enterprise - Online Cloud Portal Launcher")
print("  Developed by Deepak Pal")
print("================================================================\n")
print("[1/3] Starting local engine server on port 5000...")

# Start server.py in background
server_proc = subprocess.Popen(
    [sys.executable, str(BASE_DIR / "server.py")],
    cwd=str(BASE_DIR),
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    encoding="utf-8",
    errors="ignore"
)

# Wait briefly for server to bind
time.sleep(2)

cloudflared_exe = BASE_DIR / "cloudflared.exe"
if not cloudflared_exe.exists():
    print("[ERROR] cloudflared.exe not found in directory!")
    server_proc.terminate()
    sys.exit(1)

print("[2/3] Generating secure, free Online Cloud URL...")
tunnel_proc = subprocess.Popen(
    [str(cloudflared_exe), "tunnel", "--url", "http://127.0.0.1:5000"],
    cwd=str(BASE_DIR),
    stdout=subprocess.PIPE,
    stderr=subprocess.STDOUT,
    text=True,
    encoding="utf-8",
    errors="ignore"
)

public_url = None
start_time = time.time()
while time.time() - start_time < 30:
    line = tunnel_proc.stdout.readline()
    if not line:
        continue
    match = re.search(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com", line)
    if match:
        public_url = match.group(0)
        break

if not public_url:
    print("[WARN] Could not generate cloud link, falling back to local IP.")
    public_url = "http://localhost:5000"

print("\n" + "=" * 65)
print("  DCTTOOL V2.0 - ONLINE CLOUD PORTAL IS NOW LIVE!")
print("=" * 65)
print(f"\n  >> ONLINE PUBLIC LINK FOR OFFICE / ANYWHERE:")
print(f"     {public_url}\n")
print(f"  >> LOCAL WI-FI LINK:")
print(f"     http://localhost:5000\n")
print("=" * 65)
print("  * Ye link aap office ke kisi bhi sadasya ke sath share kar sakte hain.")
print("  * Ve bina kisi software ke sidhe Chrome me Excel file process karenge.")
print("  * Jab bhi aap naya rule update karenge, sabhi ko Reload karte hi mil jayega.")
print("=" * 65 + "\n")

# Copy to clipboard
try:
    import win32clipboard
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(public_url)
    win32clipboard.CloseClipboard()
    print("[OK] Link has been COPIED to your clipboard automatically!")
except Exception:
    try:
        import tkinter as tk
        r = tk.Tk()
        r.withdraw()
        r.clipboard_clear()
        r.clipboard_append(public_url)
        r.update()
        r.destroy()
        print("[OK] Link has been COPIED to your clipboard automatically!")
    except Exception:
        pass

# Save active URL to a text file for easy reference
with open(BASE_DIR / "ACTIVE_ONLINE_URL.txt", "w", encoding="utf-8") as f:
    f.write(f"DCTtool Online Cloud Portal Link:\n{public_url}\n\nGenerated: {time.ctime()}\n")

# Also copy active URL to Desktop for 1-click access
try:
    desktop_url_file = Path.home() / "Desktop" / "DCT_ONLINE_LINK.txt"
    with open(desktop_url_file, "w", encoding="utf-8") as f:
        f.write(f"DCTtool Online Cloud Portal Link:\n{public_url}\n\nGenerated: {time.ctime()}\nShare this link with everyone in office!\n")
except Exception:
    pass

print("[3/3] Opening Online Portal in Google Chrome...")
webbrowser.open(public_url)

print("\nServer is running actively. You can minimize this window.")
print("[Press CTRL+C anytime to stop]\n")

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    print("\nStopping portal server...")
    tunnel_proc.terminate()
    server_proc.terminate()
    print("Closed successfully.")
