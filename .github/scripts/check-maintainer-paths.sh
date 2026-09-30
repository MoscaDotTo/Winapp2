#!/usr/bin/env bash
#
# Maintainer-paths guard: does a pull request from outside the repository change a file
# only the maintainer changes?
#
# Contributions to winapp2.ini are entry data under Assembler/ and Winapp3/. Everything
# below is either shipped to users or run with write access:
#
#   winapp2ool/**                  the tool's source and the committed exes that deployed
#                                  clients download and run as their self-update. Old clients
#                                  verify nothing, so whatever lands at those paths is what
#                                  they run. *.md is exempt.
#   Assembler/winapp2ool.exe       run by build-on-merge.yml with a write token
#   Assembler/*.ps1                the build script, run the same way
#   .github/**                     the workflows and the scripts they run
#
# A change to any of these in a PR from a fork or a non-member fails. Unlike the
# generated-artifact guard there is deliberately no override token: the PR author controls
# the PR title and commit messages, so a token would let exactly the change this exists to
# catch bypass it.
#
# Usage:
#   check-maintainer-paths.sh <base-ref> <head-ref>   # diff two refs
#   check-maintainer-paths.sh --files <path>          # read a newline-separated list
# Exit 0 = no protected path changed, 1 = one did, 2 = bad invocation.

set -uo pipefail

if [ "${1:-}" = "--files" ]; then
  [ -n "${2:-}" ] && [ -f "$2" ] || { echo "usage: $0 --files <path>" >&2; exit 2; }
  changed_list=$(cat -- "$2")
elif [ -n "${1:-}" ] && [ -n "${2:-}" ]; then
  changed_list=$(git diff --name-only "$1...$2" --) || {
    echo "error: git diff $1...$2 failed" >&2; exit 2; }
else
  echo "usage: $0 <base-ref> <head-ref> | $0 --files <path>" >&2
  exit 2
fi

protected=""

while IFS= read -r f; do
  [ -n "$f" ] || continue
  case "$f" in
    winapp2ool/*.md) ;;
    winapp2ool/*|Assembler/winapp2ool.exe|.github/*) protected="$protected
$f" ;;
    Assembler/*.ps1)
      # Only the top level: scripts in flavor folders aren't run by the build
      [ "${f#Assembler/*/}" = "$f" ] && protected="$protected
$f" ;;
  esac
done <<EOF
$changed_list
EOF

if [ -z "$protected" ]; then
  echo "OK: no maintainer-only paths changed."
  exit 0
fi

while IFS= read -r f; do
  [ -n "$f" ] || continue
  if [ -n "${GITHUB_ACTIONS:-}" ]; then
    echo "::error file=$f::only the maintainer changes this file"
  else
    echo "error: $f: only the maintainer changes this file"
  fi
done <<EOF
$protected
EOF

if [ -n "${GITHUB_STEP_SUMMARY:-}" ]; then
  {
    echo "### Maintainer-paths guard: fail"
    echo
    echo "This PR changes files that only the maintainer changes:"
    echo
    printf '%s\n' "$protected" | sed '/^$/d; s/^/- `/; s/$/`/'
    echo
    echo "Entry contributions belong under \`Assembler/\` (the EntryBuilder, BrowserBuilder and UWP sources)"
    echo "or \`Winapp3/\`. If you think one of these files needs a change, open an issue instead."
  } >> "$GITHUB_STEP_SUMMARY"
fi

echo "Maintainer-paths guard FAILED."
exit 1
