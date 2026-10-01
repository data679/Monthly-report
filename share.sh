#!/usr/bin/env bash
# Share a view-only, password-protected copy of the log through a temporary Cloudflare link.
#   ./share.sh        shares for 60 minutes
#   ./share.sh 90     shares for 90 minutes
# Ctrl+C stops sharing at once. The link stops working when this script ends, and a new run
# gives a new link. Your own app (port 8765) is not shared and keeps working as usual.
cd "$(dirname "$0")"
MINUTES="${1:-60}"
VIEW_PORT=8766

CLOUDFLARED="$(command -v cloudflared || echo "$HOME/.local/bin/cloudflared")"
if [ ! -x "$CLOUDFLARED" ]; then
  echo "cloudflared isn't installed (it makes the temporary link)."
  exit 1
fi

read -r -s -p "Choose a password for viewers (12+ characters): " PW; echo
if [ "${#PW}" -lt 12 ]; then echo "Too short: use at least 12 characters."; exit 1; fi
read -r -s -p "Type it again: " PW2; echo
if [ "$PW" != "$PW2" ]; then echo "The passwords don't match."; exit 1; fi

LOG="$(mktemp)"
cleanup() {
  kill "$VIEWER_PID" "$TUNNEL_PID" $SLEEP_PID 2>/dev/null
  wait 2>/dev/null
  rm -f "$LOG"
  echo; echo "Stopped sharing. The link no longer works."
}
trap cleanup EXIT
trap 'exit 0' INT TERM

PORT=$VIEW_PORT VIEWER_PASSWORD="$PW" python3 app.py >/dev/null 2>&1 &
VIEWER_PID=$!
unset PW PW2
"$CLOUDFLARED" tunnel --no-autoupdate --url "http://127.0.0.1:$VIEW_PORT" >"$LOG" 2>&1 &
TUNNEL_PID=$!

for _ in $(seq 60); do
  URL="$(grep -o 'https://[a-z0-9-]*\.trycloudflare\.com' "$LOG" | head -1)"
  [ -n "$URL" ] && break
  if ! kill -0 "$TUNNEL_PID" 2>/dev/null; then echo "cloudflared stopped:"; cat "$LOG"; exit 1; fi
  sleep 1
done
if [ -z "$URL" ]; then echo "No link after 60 seconds:"; tail -20 "$LOG"; exit 1; fi

echo
echo "  View-only link:  $URL"
echo "  Viewers enter any user name and the password you chose."
echo "  Sharing stops automatically in $MINUTES minutes, or press Ctrl+C to stop now."
echo
sleep "$((MINUTES * 60))" &
SLEEP_PID=$!
wait $SLEEP_PID
