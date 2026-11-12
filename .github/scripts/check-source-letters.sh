#!/usr/bin/env bash
#
# Source-letter check: is every entry in a per-letter source file filed under its own letter?
#
# EntryBuilder and UWPBuilder use each source section header verbatim as the entry name, and
# -split buckets their output by that name's first character (LetterSplitOutput.LetterBucketFor:
# A-Z in any case, anything else is #). Nothing ties an entry to the file it was written in, so
# [Zoom *] in Sources/EntryBuilder/A.ini builds fine and lands in Generated/EntryBuilder/Z.ini,
# whose header then names Z.ini as its source. That's easy to miss when alphabetizing by hand,
# and it trips the generated-artifact guard's rule 2 later, when an edit to it changes Z.ini
# with only A.ini touched.
#
# We scan the whole tree rather than the changed files, so a misfiled entry already on a
# branch can't hide. It only reads section headers, so it takes well under a second.
#
# Usage:
#   check-source-letters.sh [<repo-root>]   # defaults to the current directory
# Exit 0 = every entry is in its letter's file, 1 = at least one isn't, 2 = bad invocation.

set -uo pipefail

root=${1:-.}
dirs="Assembler/Sources/EntryBuilder Assembler/Sources/UWPBuilder/AppInfo"

files=()
for d in $dirs; do
  [ -d "$root/$d" ] || { echo "error: $root/$d not found" >&2; exit 2; }
  for f in "$root/$d"/*.ini; do
    [ -e "$f" ] || continue
    # Only the 27 letter files. Anything else isn't split by letter.
    case "${f##*/}" in
      '#.ini'|[A-Z].ini) files+=("$f") ;;
    esac
  done
done

if [ "${#files[@]}" -eq 0 ]; then
  echo "error: no per-letter source files under $root" >&2
  exit 2
fi

# LC_ALL=C so awk compares bytes: the first byte of a non-ASCII name isn't A-Z, which goes
# to # just as LetterBucketFor sends it there.
misfiled=$(LC_ALL=C awk -v root="$root/" '
  FNR == 1 {
    letter = FILENAME; sub(/.*\//, "", letter); sub(/\.ini$/, "", letter)
    shown = FILENAME; if (index(shown, root) == 1) shown = substr(shown, length(root) + 1)
    sub(/^\357\273\277/, "")
  }
  {
    sub(/\r$/, "")
    if ($0 !~ /^[ \t]*\[/) next
    header = $0; sub(/^[ \t]*/, "", header)
    first = toupper(substr(header, 2, 1))
    bucket = (first ~ /^[A-Z]$/) ? first : "#"
    if (bucket != letter) printf "%s\t%d\t%s\t%s\n", shown, FNR, header, bucket
  }
' "${files[@]}")

if [ -z "$misfiled" ]; then
  echo "OK: every entry in the per-letter EntryBuilder and UWPBuilder sources is in its letter's file."
  exit 0
fi

count=0
summary=""
while IFS=$'\t' read -r file line header bucket; do
  count=$((count + 1))
  right="${file%/*}/$bucket.ini"
  msg="$header belongs in $right: the build files it under $bucket because of its first character"
  if [ -n "${GITHUB_ACTIONS:-}" ]; then
    echo "::error file=$file,line=$line::$msg"
  else
    echo "error: $file:$line: $msg"
  fi
  summary="$summary- \`$header\` is in \`$file\` (line $line), but belongs in \`$right\`
"
done <<EOF
$misfiled
EOF

if [ -n "${GITHUB_STEP_SUMMARY:-}" ]; then
  {
    echo "### Source-letter check: fail"
    echo
    printf '%s' "$summary"
    echo
    echo "The build files each entry under the first character of its name (A to Z in any case,"
    echo "anything else under \`#\`), whichever source file it's written in. Move each entry above"
    echo "into the file it belongs in."
  } >> "$GITHUB_STEP_SUMMARY"
fi

echo "Source-letter check FAILED: $count entry(s) filed under the wrong letter."
exit 1
