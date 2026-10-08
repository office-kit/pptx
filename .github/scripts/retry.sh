#!/usr/bin/env bash
# Usage: retry.sh <seconds-per-attempt> <command...>
#
# apt mirrors and Playwright's dependency install occasionally stall on the
# hosted runners without failing, which left jobs hanging until the 30-minute
# job timeout. Each attempt is killed after <seconds> and retried, so a stall
# costs minutes and usually recovers on its own.
set -u
ATTEMPTS=3
limit="$1"
shift
for attempt in $(seq 1 "$ATTEMPTS"); do
  timeout --kill-after=30 "$limit" "$@" && exit 0
  status=$?
  echo "::warning::Attempt $attempt/$ATTEMPTS of '$*' failed (exit $status)"
done
exit 1
