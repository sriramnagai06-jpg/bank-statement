"""
Bank Statement Analyzer - Local Development & Production Launcher
Serves the Flask backend and web dashboard at http://localhost:5000
"""
import os
import sys
import socket

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

if __name__ == "__main__":
    local_ip = get_local_ip()
    port = int(os.environ.get("PORT", 5000))
    
    print("\n" + "=" * 60)
    print("  Bank Statement Analyzer & Reconciliation Engine")
    print("=" * 60)
    print(f"  Local Access:       http://localhost:{port}")
    print(f"  Network (Wi-Fi):    http://{local_ip}:{port}")
    print("=" * 60 + "\n")
    
    app.run(host="0.0.0.0", port=port, debug=False)
