#!/usr/bin/env bash
# Connect the deployed wan-ocr Worker to this computer for a pilot batch.
set -euo pipefail
cd "$(dirname "$0")"
python=.venv/bin/python
if [[ ! -x "$python" ]]; then
  echo 'Set up the project .venv and install requirements.txt first.' >&2
  exit 1
fi
umask 077
work=$(mktemp -d)
backend_pid=''
tunnel_pid=''
cleanup() {
  if [[ -n "$tunnel_pid" ]]; then kill "$tunnel_pid" 2>/dev/null || true; fi
  if [[ -n "$backend_pid" ]]; then kill "$backend_pid" 2>/dev/null || true; fi
  rm -rf "$work"
}
trap cleanup EXIT
trap 'exit 130' INT TERM
if ! command -v npx >/dev/null; then
  echo 'Install Node.js with npm, then rerun this script.' >&2
  exit 1
fi
if ! command -v cloudflared >/dev/null; then
  case "$(uname -s)-$(uname -m)" in
    Darwin-x86_64) arch=amd64 ;;
    Darwin-arm64) arch=arm64 ;;
    *) echo 'Install cloudflared for this platform, then rerun.' >&2; exit 1 ;;
  esac
  echo "Downloading Cloudflare Tunnel for macOS ($arch)…"
  curl -fLsS --retry 2 -o "$work/cloudflared.tgz" \
    "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-darwin-$arch.tgz"
  tar -xzf "$work/cloudflared.tgz" -C "$work"
  "$work/cloudflared" --version
  export PATH="$work:$PATH"
fi
secret=$("$python" -c 'import secrets; print(secrets.token_hex(32))')
# Let the OS allocate a free port, even when the local dashboard is running.
WAN_PROXY_SECRET="$secret" "$python" -u - >"$work/backend.log" 2>&1 <<'PY' &
from factory import Handler, ThreadingHTTPServer
with ThreadingHTTPServer(("127.0.0.1", 0), Handler) as server:
    print(server.server_port, flush=True)
    server.serve_forever()
PY
backend_pid=$!
for ((i=0; i<60; i++)); do
  kill -0 "$backend_pid" 2>/dev/null || { cat "$work/backend.log" >&2; exit 1; }
  port=$(head -1 "$work/backend.log")
  [[ "$port" =~ ^[0-9]+$ ]] && break
  sleep 1
done
[[ "$port" =~ ^[0-9]+$ ]] || { echo 'Python did not start.' >&2; exit 1; }
cloudflared tunnel --url "http://127.0.0.1:$port" >"$work/tunnel.log" 2>&1 &
tunnel_pid=$!
for ((i=0; i<60; i++)); do
  kill -0 "$tunnel_pid" 2>/dev/null || { cat "$work/tunnel.log" >&2; exit 1; }
  origin=$(grep -Eo 'https://[a-z0-9-]+\.trycloudflare\.com' "$work/tunnel.log" | head -1 || true)
  if [[ -n "$origin" ]]; then
    status=$(curl -s --max-time 4 -o /dev/null -w '%{http_code}' "$origin/session" || true)
    [[ "$status" == 403 ]] && break
  fi
  sleep 1
done
[[ "${status:-}" == 403 ]] || { echo 'The protected tunnel did not become ready.' >&2; exit 1; }
# Authenticate interactively before passing secret values on standard input.
if [[ -z "${CLOUDFLARE_API_TOKEN:-}" ]]; then npx --yes wrangler@4 login; fi
printf '{"WAN_PROXY_SECRET":"%s","PYTHON_ORIGIN":"%s"}\n' "$secret" "$origin" |
  npx --yes wrangler@4 secret bulk --name wan-ocr
echo 'Worker settings updated. Open the hosted Analysis dashboard. Keep this Mac awake and this terminal open during the batch.'
wait "$tunnel_pid"
