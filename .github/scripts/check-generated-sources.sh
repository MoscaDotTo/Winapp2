#!/usr/bin/env bash
#
# Generated-artifact guard: did this change edit a build output without editing the source
# that produces it?
#
# `Assembler/build winapp2.ps1` generates every .ini under Assembler/Generated/ and
# Assembler/Archived/, the published files in Releases/ (each flavor, archived.ini and their
# changelogs, all named in .github/scripts/flavors.csv), and their deprecated copies (the
# old Winapp2.ini/diff.txt pairs at the root and under Non-CCleaner/, and Winapp3/Archived
# entries.ini). Before the 2026-11
# layout move the generated files lived under Assembler/Entries/, and we still treat those
# paths as generated because a branch made before the move lists them in its diff. Hand-editing any of these looks like it worked,
# since the file on disk says what you wanted. Then the next daily build regenerates it from
# the untouched source and the edit is gone. Nothing fails, the edit just disappears, and you
# find out from tomorrow's changelog instead of from CI today. This script makes it a red X
# right away.
#
# It checks provenance, not correctness. It never runs winapp2ool or compares content; it only
# asks whether the set of changed paths makes sense. artifact-drift.yml and exe-parity.yml are
# the ones that run the generators.
#
# Two rules, both hard failures:
#
#   1. Aggregate. If any generated file changed and no build source changed at all, the
#      change is a hand-edit of output. This is the usual oopsie: editing
#      Assembler/Generated/EntryBuilder/#.ini when you meant
#      Assembler/Sources/EntryBuilder/#.ini.
#
#   2. Per-artifact. Each generated artifact has a source that feeds it:
#        Assembler/Generated/EntryBuilder/<L>.ini <- Assembler/Sources/EntryBuilder/<L>.ini (1:1 by letter)
#        Assembler/Generated/BrowserBuilder/*     <- Assembler/Sources/BrowserBuilder/**
#        Assembler/Archived/**                    <- either builder's sources
#      (The archive takes entries marked Archive= from both builders, so a source edit to
#      either can move it.)
#      A changed artifact whose own source is untouched gets flagged even when some
#      unrelated source changed, which catches a correct source edit plus a stray
#      hand-edit somewhere else in Generated/.
#      We skip rule 2 when a global input changed (the "global inputs" cases below),
#      because those can move artifacts whose per-letter source is untouched. The shared
#      scaffold catalogs are one: package.ini alone feeds every Store-app entry.
#
#      The letter rule assumes every entry is in its own letter's source file. EntryBuilder
#      buckets its output by entry name, not by source file, so an entry filed under the
#      wrong letter (say [Zoom *] in Sources/EntryBuilder/A.ini) lands in Z.ini, and editing
#      it changes Z.ini alongside A.ini, which rule 2 would flag. check-source-letters.sh
#      fails any such entry in the same workflow.
#
# UWPBuilder was folded into EntryBuilder in 2026-11, and a Store app is now an EntryBuilder
# entry with Package=. Its three folders (Sources/UWPBuilder/, Generated/UWPBuilder/ and
# Archived/UWPBuilder/) and their pre-move spellings are retired: deleting a file there passes,
# and adding or changing one fails, since no build reads it and its entries would silently
# drop out. This is the one place a path's status matters, which is why we list changes with
# --name-status.
#
# The aggregate published files (Releases/ and the deprecated Winapp2.ini, diff.txt,
# Non-CCleaner/** and Winapp3/Archived entries.ini copies) only get rule 1. Any source
# anywhere can move them, so there's no narrower source to ask for.
#
# Known limitation: rule 1 can't tell a hand-edit from a deliberate regeneration. "Output
# moved, source didn't" has two innocent causes besides the oopsie: seeding a new artifact
# for the first time, and committing a regeneration to repair drift that some earlier
# change introduced (an exe rebuild, say). Both are output-only commits and both trip
# rule 1. Telling them apart means running the build, which takes minutes on a Windows
# runner, and this check takes seconds because it only looks at path names. So the
# innocent cases declare themselves: put [regen] (or [skip-generated-check]) in the commit
# message or PR title.
#
# For calibration: in the 120 commits before 2026-07-28, five human commits touched
# generated files. One was clean, two were drift regenerations and two seeded a new
# artifact, so expect to need the token a few times a year, nearly always on a commit whose
# message already says "regenerate". Rule 2 flagged none of the five.
#
# Usage:
#   check-generated-sources.sh <base-ref> <head-ref>   # diff two refs
#   check-generated-sources.sh --files <path>          # read a newline-separated list
#                                                     # (for local testing; a line may be a
#                                                     # bare path, read as modified, or
#                                                     # --name-status's "<status><TAB><path>")
# Exit 0 = coherent, 1 = incoherent, 2 = bad invocation.

set -uo pipefail

# ---------------------------------------------------------------- collect changed paths

changed_list=""

if [ "${1:-}" = "--files" ]; then
  [ -n "${2:-}" ] && [ -f "$2" ] || { echo "usage: $0 --files <path>" >&2; exit 2; }
  changed_list=$(cat -- "$2")
elif [ -n "${1:-}" ] && [ -n "${2:-}" ]; then
  # Three-dot: what the head added relative to the merge base, so unrelated commits
  # that landed on the base branch meanwhile are not attributed to this change.
  # --no-renames lists a rename as a deletion and an addition, so both paths get classified.
  changed_list=$(git diff --name-status --no-renames "$1...$2" --) || {
    echo "error: git diff $1...$2 failed" >&2; exit 2; }
else
  echo "usage: $0 <base-ref> <head-ref> | $0 --files <path>" >&2
  exit 2
fi

if [ -z "$changed_list" ]; then
  echo "No changed files; nothing to check."
  exit 0
fi

# Emitted as GitHub annotations when running in Actions, plain text otherwise.
fail=0
summary=""

note() { summary="${summary}$1
"; }

annotate() {
  # annotate <file> <message>
  if [ -n "${GITHUB_ACTIONS:-}" ]; then
    echo "::error file=$1::$2"
  else
    echo "error: $1: $2"
  fi
}

# ---------------------------------------------------------------------------- classify

# Generated artifacts touched, split by kind.
gen_letters=""      # bare letters, e.g. "A M #"
gen_browsers=0
gen_other=""        # other Generated/ files + the published aggregates
gen_archived=""     # Assembler/Archived/ files
gen_archive_copy=0  # Winapp3/Archived entries.ini, kept out of gen_other because of its space
gen_count=0

# Build inputs touched.
src_count=0
src_entrybuilder=""     # bare letters whose EntryBuilder source changed
src_browserbuilder=0
global=0                # a global input changed -> suppress rule 2
global_why=""

retired=""              # files added or changed under a retired UWPBuilder folder

# Each reason is listed once, however many of its files changed.
add_why() {
  case ", $global_why, " in
    *", $1, "*) ;;
    *) global_why="${global_why}${global_why:+, }$1" ;;
  esac
}

tab=$(printf '\t')
cr=$(printf '\r')
nl=$(printf '\nx'); nl=${nl%x}

# Every Asset and Changelog name in the flavor manifest, each between newlines.
manifest="$(dirname -- "$0")/flavors.csv"
[ -f "$manifest" ] || { echo "error: flavor manifest missing: $manifest" >&2; exit 2; }
release_names="$nl"
header=1
# The || test keeps a last row that has no trailing newline.
while IFS=, read -r _key _label asset changelog _old || [ -n "$_key" ]; do
  if [ "$header" -eq 1 ]; then header=0; continue; fi
  [ -n "${asset%"$cr"}" ] || continue
  release_names="$release_names${asset%"$cr"}$nl${changelog%"$cr"}$nl"
done < "$manifest"

while IFS= read -r f; do
  [ -n "$f" ] || continue
  status=M
  case "$f" in
    ?"$tab"*) status=${f%%"$tab"*}; f=${f#*"$tab"} ;;
  esac

  case "$f" in
    # Retired by the UWPBuilder fold, before *.md so a readme can't revive a folder.
    # Assembler/UWP/ and Assembler/Entries/UWP/ are the pre-move spellings.
    'Assembler/Sources/UWPBuilder/'*|'Assembler/Generated/UWPBuilder/'*|'Assembler/Archived/UWPBuilder/'*|'Assembler/UWP/'*|'Assembler/Entries/UWP/'*)
      [ "$status" = D ] || retired="$retired $f"
      continue ;;

    # Docs never affect the build.
    *.md) continue ;;

    # ---- global inputs: these can change any artifact's content on their own.
    # Sources/Scaffolds comes before any other Sources/ case. The bare Assembler/Scaffolds/
    # is its pre-move path (now Assembler/Sources/Scaffolds/), kept for branches made earlier.
    'Assembler/Sources/Scaffolds/'*|'Assembler/Scaffolds/'*)
      src_count=$((src_count + 1)); global=1
      add_why "shared scaffold catalogs"
      continue ;;
    'Assembler/winapp2ool.exe')
      src_count=$((src_count + 1)); global=1
      add_why "the committed winapp2ool.exe"
      continue ;;
    'Assembler/build winapp2.ps1')
      src_count=$((src_count + 1)); global=1
      add_why "the build script"
      continue ;;
    '.github/scripts/flavors.csv')
      src_count=$((src_count + 1)); global=1
      add_why "the flavor manifest"
      continue ;;
    '.github/scripts/release-assets.ps1')
      src_count=$((src_count + 1)); global=1
      add_why "the release asset list"
      continue ;;
    'winapp2ool/'*)
      src_count=$((src_count + 1)); global=1
      add_why "winapp2ool source"
      continue ;;

    # ---- generated artifacts.
    # The old Assembler/Entries/** patterns are the pre-move paths (Entries/Browsers/ is now
    # Generated/BrowserBuilder/, Entries/<L>.ini is Generated/EntryBuilder/<L>.ini), kept for
    # branches made before the move.
    'Assembler/Generated/BrowserBuilder/'*|'Assembler/Entries/Browsers/'*)
      gen_browsers=1; gen_count=$((gen_count + 1)) ;;
    'Assembler/Generated/EntryBuilder/'*.ini|'Assembler/Entries/'*.ini)
      base=${f##*/}; letter=${base%.ini}
      gen_count=$((gen_count + 1))
      case "$letter" in
        # The 27 buckets EntryBuilder -split writes. Anything else isn't a per-letter
        # artifact and goes to gen_other.
        '#'|[A-Z]) gen_letters="$gen_letters $letter" ;;
        *)         gen_other="$gen_other $f" ;;
      esac ;;
    'Assembler/Generated/'*|'Assembler/Entries/'*)
      gen_count=$((gen_count + 1)); gen_other="$gen_other $f" ;;
    # The archive tree: each builder's archived.ini. Before the Assembler/ catch-all below.
    'Assembler/Archived/'*)
      gen_count=$((gen_count + 1)); gen_archived="$gen_archived $f" ;;
    # The release data files and their changelogs, matched by exact name so nothing else in
    # Releases/ counts.
    'Releases/'*)
      case "$release_names" in
        *"$nl${f#Releases/}$nl"*) gen_count=$((gen_count + 1)); gen_other="$gen_other $f" ;;
      esac ;;
    # Their deprecated copies, published until the deprecation window closes.
    'Winapp2.ini'|'diff.txt'|'Non-CCleaner/'*'Winapp2.ini'|'Non-CCleaner/'*'Winapp2.rules'|'Non-CCleaner/'*'diff.txt')
      gen_count=$((gen_count + 1)); gen_other="$gen_other $f" ;;
    'Winapp3/Archived entries.ini')
      gen_count=$((gen_count + 1)); gen_archive_copy=1 ;;

    # ---- per-artifact sources.
    # The bare Assembler/EntryBuilder/ and BrowserBuilder/ patterns are the pre-move paths
    # (now under Assembler/Sources/), kept so a branch made before the move that commits a
    # source with its regeneration still matches them up.
    'Assembler/Sources/EntryBuilder/'*.ini|'Assembler/EntryBuilder/'*.ini)
      src_count=$((src_count + 1))
      base=${f##*/}
      src_entrybuilder="$src_entrybuilder ${base%.ini}" ;;
    'Assembler/Sources/BrowserBuilder/'*|'Assembler/BrowserBuilder/'*)
      src_count=$((src_count + 1)); src_browserbuilder=1 ;;

    # ---- anything else under Assembler/ is a build input: Sources/Flavors Transmute
    # directories, EntryBuilder files that aren't .ini, future generator sources.
    # Counting anything unlisted as a source means a new flavor or generator folder
    # doesn't need an edit here to avoid a false failure.
    'Assembler/'*)
      src_count=$((src_count + 1)) ;;
  esac
done <<EOF
$changed_list
EOF

# ------------------------------------------------------------------------------ verdict

if [ -n "$retired" ]; then
  fail=1
  note "**Files were added or changed in a folder retired by the UWPBuilder fold.**"
  note ""
  note "UWPBuilder was folded into EntryBuilder, and nothing reads these folders anymore."
  note "A Store app is an EntryBuilder entry with \`Package=\`: put it in"
  note "\`Assembler/Sources/EntryBuilder/<L>.ini\` under its name's first letter."
  for f in $retired; do
    annotate "$f" \
      "UWPBuilder was folded into EntryBuilder; Store-app entries belong in Assembler/Sources/EntryBuilder/<L>.ini"
    note "- \`$f\`"
  done
  note ""
fi

if [ "$gen_count" -eq 0 ]; then
  if [ "$fail" -eq 0 ]; then
    echo "No generated artifacts changed; nothing to check."
    exit 0
  fi

elif [ "$src_count" -eq 0 ]; then
  # Rule 1.
  fail=1
  note "**$gen_count generated file(s) changed, and no build source changed.**"
  note ""
  note "If this was meant to be a content change: these files are build output. The next"
  note "build regenerates them from their sources and will silently discard the edit."
  note "Make it in the source instead (each file's source is listed below)."
  note ""
  note "If you committed regenerated output on purpose (repairing drift, or seeding a new"
  note "artifact), that's fine, but this check can't tell the difference. Add \`[regen]\`"
  note "to the commit message or PR title."
  for letter in $gen_letters; do
    annotate "Assembler/Generated/EntryBuilder/$letter.ini" \
      "generated by EntryBuilder; edit Assembler/Sources/EntryBuilder/$letter.ini instead"
    note "- \`Assembler/Generated/EntryBuilder/$letter.ini\` <- \`Assembler/Sources/EntryBuilder/$letter.ini\`"
  done
  if [ "$gen_browsers" -eq 1 ]; then
    annotate "Assembler/Generated/BrowserBuilder/browsers.ini" \
      "generated by BrowserBuilder; edit the sources in Assembler/Sources/BrowserBuilder/ instead"
    note "- \`Assembler/Generated/BrowserBuilder/browsers.ini\` <- \`Assembler/Sources/BrowserBuilder/\`"
  fi
  for f in $gen_archived; do
    annotate "$f" \
      "generated from the entries marked Archive= in the builders' sources; archive or restore an entry there instead"
    note "- \`$f\` <- the entries marked \`Archive=\` in \`Assembler/Sources/\` (EntryBuilder, BrowserBuilder)"
  done
  for f in $gen_other; do
    annotate "$f" "build output; edit the sources under Assembler/ instead"
    note "- \`$f\` <- the build (\`Assembler/build winapp2.ps1\`)"
  done
  if [ "$gen_archive_copy" -eq 1 ]; then
    annotate "Winapp3/Archived entries.ini" \
      "build output; archive an entry by adding Archive=<reason> to it in its source under Assembler/Sources/ instead"
    note "- \`Winapp3/Archived entries.ini\` <- the build (\`Assembler/build winapp2.ps1\`)"
  fi

elif [ "$global" -eq 1 ]; then
  # No rule 2: a global input can move artifacts whose own per-letter source is untouched.
  echo "Build sources changed, including a global input ($global_why)."
  echo "Skipping the per-artifact source check, since a global input can change any artifact."

else
  # Rule 2.
  orphans=""
  for letter in $gen_letters; do
    matched=0
    for s in $src_entrybuilder; do
      [ "$s" = "$letter" ] && { matched=1; break; }
    done
    if [ "$matched" -eq 0 ]; then
      fail=1
      orphans=1
      annotate "Assembler/Generated/EntryBuilder/$letter.ini" \
        "changed, but its source Assembler/Sources/EntryBuilder/$letter.ini did not, so the next build will overwrite this edit"
      note "- \`Assembler/Generated/EntryBuilder/$letter.ini\` changed, but \`Assembler/Sources/EntryBuilder/$letter.ini\` did not"
    fi
  done
  if [ "$gen_browsers" -eq 1 ] && [ "$src_browserbuilder" -eq 0 ]; then
    fail=1; orphans=1
    annotate "Assembler/Generated/BrowserBuilder/browsers.ini" \
      "changed, but no source under Assembler/Sources/BrowserBuilder/ did, so the next build will overwrite this edit"
    note "- \`Assembler/Generated/BrowserBuilder/browsers.ini\` changed, but nothing under \`Assembler/Sources/BrowserBuilder/\` did"
  fi
  if [ -n "$gen_archived" ] && [ -z "$src_entrybuilder" ] && [ "$src_browserbuilder" -eq 0 ]; then
    fail=1; orphans=1
    for f in $gen_archived; do
      annotate "$f" \
        "changed, but no EntryBuilder or BrowserBuilder source did, so the next build will overwrite this edit"
      note "- \`$f\` changed, but no source under \`Assembler/Sources/EntryBuilder/\` or \`BrowserBuilder/\` did"
    done
  fi

  if [ -n "$orphans" ]; then
    note ""
    note "Each artifact above is regenerated from the source next to it. A generated"
    note "file that moves without its source is a hand-edit, and the next build reverts it."
  fi
fi

# ------------------------------------------------------------------------------- report

if [ "$fail" -eq 0 ]; then
  echo "OK: $gen_count generated file(s) changed, and their sources changed too."
  if [ -n "${GITHUB_STEP_SUMMARY:-}" ]; then
    {
      echo "### Generated-artifact guard: pass"
      echo
      echo "$gen_count generated file(s) changed alongside their build sources."
    } >> "$GITHUB_STEP_SUMMARY"
  fi
  exit 0
fi

if [ -n "${GITHUB_STEP_SUMMARY:-}" ]; then
  {
    echo "### Generated-artifact guard: fail"
    echo
    printf '%s' "$summary"
    echo
    echo "Everything under \`Assembler/Generated/\` and \`Assembler/Archived/\`, the published files"
    echo "in \`Releases/\`, and their old copies (\`Winapp2.ini\`, \`diff.txt\`, \`Winapp3/Archived entries.ini\`)"
    echo "are **build output**, produced by \`Assembler/build winapp2.ps1\`. Edit the sources and let"
    echo "the build write the output. To retire an entry, add \`Archive=<reason>\` to it in its source."
    echo
    echo "If you committed regenerated output on purpose, put \`[regen]\` or"
    echo "\`[skip-generated-check]\` in the commit message or PR title to override."
  } >> "$GITHUB_STEP_SUMMARY"
fi

echo
printf '%s' "$summary" | sed 's/^/  /'
echo
echo "Generated-artifact guard FAILED. Edit sources, not build output."
exit 1
