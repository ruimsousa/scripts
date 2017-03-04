#! /bin/bash
MY_NAME=${0##*/}
HOST="$1"
[ -z "$HOST" ] && { echo "Usage: $MY_NAME HOST [PORT]" ; exit 1 ; }
PORT="$2"
[ -z "$PORT" ] && PORT=9100
if ( ( echo -n '' >/dev/tcp/$HOST/$PORT ) & PID=$! ; sleep 2 ; kill $PID ; wait $PID )
then if ( ( echo -en '\r' >/dev/tcp/$HOST/$PORT ) & PID=$! ; sleep 2 ; kill $PID ; wait $PID )
     then echo "Port $PORT on host $HOST accepts data"
          exit 0
     fi
     echo "Connection possible to port $PORT on host $HOST but does not accept data"
     exit 1
fi
echo "No connection possible to port $PORT on host $HOST"
exit 1

