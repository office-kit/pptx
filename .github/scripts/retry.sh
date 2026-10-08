#!/usr/bin/env bash
# Usage: retry.sh <seconds-per-attempt> <command...>
#
# apt mirrors and Playwright's dependency install occasionally stall on the
# hosted runners without failing, which left jobs hanging until the 30-minute
# job timeout. Each attempt is killed after <seconds> and retried, so a stall
# costs minutes and usually recovers on its own.
set -u
ATTEMPTS=3
LOCK_WAIT_SECONDS=120
limit="$1"
shift

# `timeout` only signals the command it started. An apt-get launched through
# sudo runs in its own session (Ubuntu's sudoers sets use_pty), so it survives
# the kill and keeps holding the dpkg lock, and every retry then fails at once
# with "Could not get lock /var/lib/dpkg/lock-frontend". Stop any leftover
# apt/dpkg, wait for the lock, and repair a half-configured dpkg before retrying.
release_apt() {
  command -v apt-get >/dev/null 2>&1 || return 0
  sudo pkill -TERM -x apt-get 2>/dev/null
  sudo pkill -TERM -x dpkg 2>/dev/null
  local waited=0
  while sudo fuser /var/lib/dpkg/lock-frontend /var/lib/dpkg/lock >/dev/null 2>&1; do
    if [ "$waited" -ge "$LOCK_WAIT_SECONDS" ]; then
      sudo pkill -KILL -x apt-get 2>/dev/null
      sudo pkill -KILL -x dpkg 2>/dev/null
      break
    fi
    sleep 2
    waited=$((waited + 2))
  done
  sudo dpkg --configure -a >/dev/null 2>&1
  return 0
}

for attempt in $(seq 1 "$ATTEMPTS"); do
  timeout --kill-after=30 "$limit" "$@" && exit 0
  status=$?
  echo "::warning::Attempt $attempt/$ATTEMPTS of '$*' failed (exit $status)"
  [ "$attempt" -lt "$ATTEMPTS" ] && release_apt
done
exit 1
