#!/bin/bash
sleep 3
curl -s -X POST http://192.168.2.111:18800/api/register \
  -H "Authorization: Bearer jzkj2026" \
  -H "Content-Type: application/json" \
  -d '{"nodeId":"N2","ip":"192.168.2.109","port":3000}' \
  >> /tmp/ai001-register.log 2>&1
