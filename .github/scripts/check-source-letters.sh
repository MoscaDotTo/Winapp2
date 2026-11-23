#!/usr/bin/env bash
#
# Source-letter check: is every entry in a per-letter source file filed under its own letter?
#
# EntryBuilder uses each source section header verbatim as the entry name, and
# -split buckets its output by that name's first character (LetterSplitOutput.LetterBucketFor:
# A-Z in any case, anything else is #). Nothing ties an entry to the file it was written in, so
# [Zoom *] in Sources/EntryBuilder/A.ini builds fine and lands in Generated/EntryBuilder/Z.ini,
# whose header then names Z.ini as its source. That's easy to miss when alphabetizing by hand,
# and it trips the generated-artifact guard's rule 2 later, when an edit to it changes Z.ini
# with only A.ini touched.
#
# Sources/Winapp3 is split by the same rule. Its run writes one unsplit winapp3.ini, so a
# misfiled entry there still builds, but people look for an entry in its letter's file.
#
# We scan the whole tree rather than the changed files, so a misfiled entry already on a
# branch can't hide. It only reads section headers, so it takes well under a second.
#
# It also fails on any file under the three UWPBuilder folders. UWPBuilder was folded into
# EntryBuilder, so nothing builds or writes them. A stale branch that brings back
# Sources/UWPBuilder/ would drop its entries while the build stays green, and the full build's
# Combine would read a revived Generated/UWPBuilder/. The generated-artifact guard skips
# [regen] commits, so only this whole-tree check sees the output folders every time.
#
# Usage:
#   check-source-letters.sh [<repo-root>]   # defaults to the current directory
# Exit 0 = every entry is in its letter's file, 1 = at least one isn't, 2 = bad invocation.

set -uo pipefail

root=${1:-.}
dir="Assembler/Sources/EntryBuilder"
letter_dirs="$dir Assembler/Sources/Winapp3"
retired_dirs="Assembler/Sources/UWPBuilder Assembler/Generated/UWPBuilder Assembler/Archived/UWPBuilder"

retired=$(for d in $retired_dirs; do find "$root/$d" -type f 2>/dev/null; done)
if [ -n "$retired" ]; then
  while IFS= read -r f; do
    shown=${f#"$root/"}
    msg="UWPBuilder was folded into EntryBuilder and nothing builds or writes this file. A Store app is an EntryBuilder entry with Package= in $dir/<L>.ini: move any entry there and delete the file"
    if [ -n "${GITHUB_ACTIONS:-}" ]; then
      echo "::error file=$shown::$msg"
    else
      echo "error: $shown: $msg"
    fi
  done <<EOF
$retired
EOF
  if [ -n "${GITHUB_STEP_SUMMARY:-}" ]; then
    {
      echo "### Source-letter check: fail"
      echo
      echo "The UWPBuilder folders under \`Assembler/\` were retired when UWPBuilder was folded into"
      echo "EntryBuilder. Move each Store-app entry into \`$dir/<L>.ini\` as an EntryBuilder entry"
      echo "with \`Package=\`, and delete the folders."
    } >> "$GITHUB_STEP_SUMMARY"
  fi
  echo "Source-letter check FAILED: a UWPBuilder folder exists, but UWPBuilder was folded into EntryBuilder."
  exit 1
fi

files=()
for d in $letter_dirs; do
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
  echo "OK: every entry in the per-letter EntryBuilder and Winapp3 sources is in its letter's file."
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
