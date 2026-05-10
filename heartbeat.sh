#!/bin/sh
while true; do
  curl -s -X POST http://192.168.2.111:18800/api/heartbeat \
    -H "Authorization: Bearer jzkj2026" \
    -H "Content-Type: application/json" \
    -d '{"nodeId":"N2"}' > /dev/null 2>&1
  sleep 60
done
