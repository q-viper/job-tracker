#!/usr/bin/env bash

if [ -z "${BASH_VERSION:-}" ]; then
  exec bash "$0" "$@"
fi

if [[ "${BASH_SOURCE[0]}" != "$0" ]]; then
  echo "Please run this script directly: bash tunnel.sh"
  return 1 2>/dev/null || exit 1
fi

set -Eeuo pipefail

on_error() {
  local exit_code="$1"
  local line_no="$2"
  local failed_cmd="$3"
  echo
  echo "[error] Command failed at line ${line_no}: ${failed_cmd}"
  echo "[error] Exit code: ${exit_code}"
}
trap 'on_error "$?" "$LINENO" "$BASH_COMMAND"' ERR

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

PYTHON_BIN=""
if command -v python >/dev/null 2>&1; then
  PYTHON_BIN="python"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_BIN="python3"
else
  echo "[error] Python not found. Install Python 3.11+ and retry."
  exit 1
fi

ensure_cloudflared() {
  if command -v cloudflared >/dev/null 2>&1; then
    return 0
  fi

  echo "[setup] cloudflared not found; installing..."
  if command -v apt-get >/dev/null 2>&1; then
    if ! command -v curl >/dev/null 2>&1; then
      sudo apt-get update
      sudo apt-get install -y curl
    fi
    sudo apt-get update
    if ! sudo apt-get install -y cloudflared; then
      echo "[setup] cloudflared package not in default apt repos; adding Cloudflare repo..."
      sudo mkdir -p --mode=0755 /usr/share/keyrings
      curl -fsSL https://pkg.cloudflare.com/cloudflare-main.gpg | sudo tee /usr/share/keyrings/cloudflare-main.gpg >/dev/null
      echo "deb [signed-by=/usr/share/keyrings/cloudflare-main.gpg] https://pkg.cloudflare.com/cloudflared any main" | sudo tee /etc/apt/sources.list.d/cloudflared.list >/dev/null
      sudo apt-get update
      sudo apt-get install -y cloudflared
    fi
  elif command -v dnf >/dev/null 2>&1; then
    sudo dnf install -y cloudflared
  elif command -v yum >/dev/null 2>&1; then
    sudo yum install -y cloudflared
  else
    echo "[error] Unsupported package manager. Install cloudflared manually:"
    echo "https://developers.cloudflare.com/cloudflare-one/connections/connect-networks/downloads/"
    exit 1
  fi

  if ! command -v cloudflared >/dev/null 2>&1; then
    echo "[error] cloudflared install failed."
    exit 1
  fi
}

PORT="$("$PYTHON_BIN" - <<'PY'
import socket

def is_bindable(host: str, port: int) -> bool:
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind((host, port))
        return True
    except OSError:
        return False
    finally:
        s.close()

for p in range(8001, 65536):
    if is_bindable("0.0.0.0", p):
        print(p)
        break
else:
    raise RuntimeError("No available TCP port found in range 8001-65535.")
PY
)"

echo "[setup] Using Python: $PYTHON_BIN"
echo "[setup] App will run on local port: $PORT"

ensure_cloudflared

cleanup() {
  if [ -n "${APP_PID:-}" ] && kill -0 "$APP_PID" 2>/dev/null; then
    kill "$APP_PID" 2>/dev/null || true
  fi
  if [ -n "${TUNNEL_PID:-}" ] && kill -0 "$TUNNEL_PID" 2>/dev/null; then
    kill "$TUNNEL_PID" 2>/dev/null || true
  fi
}
trap cleanup EXIT INT TERM

echo "[run] Starting Django server..."
"$PYTHON_BIN" run.py serve --host "0.0.0.0" --port ${PORT} > /tmp/job_tracker__app.log 2>&1 &
APP_PID=$!

sleep 2
if ! kill -0 "$APP_PID" 2>/dev/null; then
  echo "[error] Django server failed to start. Logs:"
  cat /tmp/job_tracker__app.log || true
  exit 1
fi

echo "[run] Starting Cloudflare tunnel..."
cloudflared tunnel --url "http://127.0.0.1:${PORT}" > /tmp/job_tracker__tunnel.log 2>&1 &
TUNNEL_PID=$!

echo "[info] Waiting for public URL..."
PUBLIC_URL=""
for _ in $(seq 1 30); do
  if grep -Eo 'https://[-a-zA-Z0-9]+\.trycloudflare\.com' /tmp/job_tracker__tunnel.log >/dev/null 2>&1; then
    PUBLIC_URL="$(grep -Eo 'https://[-a-zA-Z0-9]+\.trycloudflare\.com' /tmp/job_tracker__tunnel.log | head -n1)"
    break
  fi
  sleep 1
done

if [ -z "$PUBLIC_URL" ]; then
  echo "[warn] Could not detect public URL yet. Check logs: /tmp/job_tracker__tunnel.log"
else
  echo
  echo "[public] ${PUBLIC_URL}"
  echo ${PUBLIC_URL} > 'url.txt'
  echo
fi

echo "[info] Press Ctrl+C to stop app and tunnel."
wait "$TUNNEL_PID"
