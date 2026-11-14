#!/bin/sh
case "$1" in
  Username*) echo "x-access-token" ;;
  *) cat "$HOME/.workbuddy/mdmq_token.txt" | tr -d '\r\n' ;;
esac
