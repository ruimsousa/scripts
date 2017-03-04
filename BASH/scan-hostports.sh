#!/bin/bash
host=$1
port_first=1
port_last=65535
for ((port=$port_first; port<=$port_last; port++))
do
  echo "scan port: $port"
  (echo >/dev/tcp/$host/$port) >/dev/null 2>&1 && echo "$port open"
done
