"""
Universal One-Click Launcher for Bank Statement Analyzer.
Works for humans and AI agents on ANY Laptop (Windows, Mac, Linux).

Provides 3 access modes simultaneously:
1. Local Laptop Access:       http://localhost:5000
2. Same Wi-Fi Network Access: http://<LOCAL-IP>:5000
3. Different Wi-Fi / Cellular: https://<PUBLIC-URL>.lhr.life (Worldwide)
"""
import os
import sys
import socket
import subprocess
import threading
import time
import re

# Force unbuffered output
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BACKEND_DIR = os.path.join(BASE_DIR, "backend")
sys.path.insert(0, BACKEND_DIR)

from app import app

def get_local_ip():
    """Find local IPv4 address of this machine."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('10.255.255.255', 1))
        ip = s.getsockname()[0]
    except Exception:
        ip = '127.0.0.1'
    finally:
        s.close()
    return ip

def wait_for_local_port(port=5000, timeout=15):
    """Wait until local port 5000 is active and accepting connections."""
    start = time.time()
    while time.time() - start < timeout:
        try:
            with socket.create_connection(('127.0.0.1', port), timeout=1):
                return True
        except OSError:
            time.sleep(0.5)
    return False

def start_public_tunnel():
    """Start official Cloudflare Tunnel to expose port 5000 to the public internet."""
    wait_for_local_port(5000)
    time.sleep(0.5)
    
    cf_path = os.path.join(BASE_DIR, "cloudflared_win.exe")
    if not os.path.exists(cf_path):
        cf_path = os.path.join(BASE_DIR, "cloudflared.exe")

    cmd = [cf_path, "tunnel", "--url", "http://127.0.0.1:5000"]
    pattern = r'https://(?!api\.)[a-zA-Z0-9-]+\.trycloudflare\.com'

    while True:
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1
            )
            for line in proc.stdout:
                if "trycloudflare.com" in line:
                    match = re.search(pattern, line)
                    if match:
                        url = match.group(0)
                        print("\n" + "=" * 65, flush=True)
                        print(" WORLDWIDE PUBLIC LINK (Works on ANY Phone & Laptop, Mobile Data / 5G / Wi-Fi):", flush=True)
                        print(f" >>> {url} <<<", flush=True)
                        print("=" * 65 + "\n", flush=True)
            proc.wait()
        except Exception as e:
            print(f"\n[Tunnel Info] Tunnel status: {e}", flush=True)
        time.sleep(2)

if __name__ == "__main__":
    local_ip = get_local_ip()
    hostname = socket.gethostname()
    
    print("\n" + "=" * 65, flush=True)
    print("  Bank Statement Analyzer — Universal Server", flush=True)
    print("=" * 65, flush=True)
    print(f"  1. Local Laptop:           http://localhost:5000", flush=True)
    print(f"  2. Same Wi-Fi (Network):   http://{local_ip}:5000", flush=True)
    print(f"  3. Hostname:               http://{hostname}:5000", flush=True)
    print("=" * 65, flush=True)
    print("  Starting public internet link for DIFFERENT Wi-Fi networks...", flush=True)
    
    # Launch public tunnel in background thread
    tunnel_thread = threading.Thread(target=start_public_tunnel, daemon=True)
    tunnel_thread.start()
    
    # Run Flask server
    app.run(host="0.0.0.0", port=5000, debug=False)
