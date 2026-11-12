#!/usr/bin/env bash
#
# Generated-artifact guard: did this change edit a build output without editing the source
# that produces it?
#
# `Assembler/build winapp2.ps1` generates every .ini under Assembler/Generated/, the 14
# published files in Releases/ (seven flavor files and their changelogs), and their
# deprecated copies (the seven old Winapp2.ini/diff.txt pairs at the root and under
# Non-CCleaner/). Before the 2026-11 layout move the generated files lived under
# Assembler/Entries/, and we still treat those paths as generated because a branch made
# before the move lists them in its diff. Hand-editing any of these looks like it worked,
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
#        Assembler/Generated/UWPBuilder/<L>.ini   <- Assembler/Sources/UWPBuilder/AppInfo/<L>.ini (1:1 by letter)
#        Assembler/Generated/BrowserBuilder/*     <- Assembler/Sources/BrowserBuilder/**
#        any other Assembler/Generated/UWPBuilder/* <- Assembler/Sources/UWPBuilder/**
#      A changed artifact whose own source is untouched gets flagged even when some
#      unrelated source changed, which catches a correct source edit plus a stray
#      hand-edit somewhere else in Generated/.
#      We skip rule 2 when a global input changed (the "global inputs" cases below),
#      because those can move artifacts whose per-letter source is untouched. A change to
#      any Sources/UWPBuilder/ file outside AppInfo/<L>.ini (UWP.ini, the scaffold every
#      UWP entry is built from) can move every UWP letter, so it skips only the UWP letter
#      rule. The EntryBuilder letter rule still applies.
#
#      Both letter rules assume every entry is in its own letter's source file. The
#      generators bucket their output by entry name, not by source file, so an entry filed
#      under the wrong letter (say [Zoom *] in AppInfo/A.ini) lands in Z.ini, and editing it
#      changes Z.ini alongside A.ini, which rule 2 would flag. check-source-letters.sh fails
#      any such entry in the same workflow.
#
# The aggregate published files (Releases/ and the deprecated Winapp2.ini, diff.txt and
# Non-CCleaner/** copies) only get rule 1. Any source anywhere can move them, so there's no
# narrower source to ask for.
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
#                                                     # (for local testing)
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
  changed_list=$(git diff --name-only "$1...$2" --) || {
    echo "error: git diff $1...$2 failed" >&2; exit 2; }
else
  echo "usage: $0 <base-ref> <head-ref> | $0 --files <path>" >&2
  exit 2
fi

if [ -z "$changed_list" ]; then
  echo "No changed files; nothing to check."
  exit 0
fi

# ---------------------------------------------------------------------------- classify

# Generated artifacts touched, split by kind.
gen_letters=""      # bare letters, e.g. "A M #"
gen_uwp_letters=""  # bare letters of Generated/UWPBuilder/<L>.ini
gen_browsers=0
gen_uwp=""          # other Generated/UWPBuilder files (and the pre-move Entries/UWP/)
gen_other=""        # other Generated/ files + the published aggregates
gen_count=0

# Build inputs touched.
src_count=0
src_entrybuilder=""     # bare letters whose EntryBuilder source changed
src_browserbuilder=0
src_uwp=0               # any UWPBuilder source changed
src_uwp_letters=""      # bare letters whose AppInfo source changed
src_uwp_global=0        # a UWPBuilder source outside AppInfo/<L>.ini changed
global=0                # a global input changed -> suppress rule 2
global_why=""

while IFS= read -r f; do
  [ -n "$f" ] || continue

  case "$f" in
    # Docs never affect the build.
    *.md) continue ;;

    # ---- global inputs: these can change any artifact's content on their own.
    # Sources/Scaffolds comes before any other Sources/ case. The bare Assembler/Scaffolds/
    # is its pre-move path (now Assembler/Sources/Scaffolds/), kept for branches made earlier.
    'Assembler/Sources/Scaffolds/'*|'Assembler/Scaffolds/'*)
      src_count=$((src_count + 1)); global=1
      global_why="${global_why}${global_why:+, }shared scaffold catalogs"
      continue ;;
    'Assembler/winapp2ool.exe')
      src_count=$((src_count + 1)); global=1
      global_why="${global_why}${global_why:+, }the committed winapp2ool.exe"
      continue ;;
    'Assembler/build winapp2.ps1')
      src_count=$((src_count + 1)); global=1
      global_why="${global_why}${global_why:+, }the build script"
      continue ;;
    'winapp2ool/'*)
      src_count=$((src_count + 1)); global=1
      global_why="${global_why}${global_why:+, }winapp2ool source"
      continue ;;

    # ---- generated artifacts.
    # The old Assembler/Entries/** patterns are the pre-move paths (Entries/Browsers/ is now
    # Generated/BrowserBuilder/, Entries/UWP/ is Generated/UWPBuilder/, Entries/<L>.ini is
    # Generated/EntryBuilder/<L>.ini), kept for branches made before the move.
    'Assembler/Generated/BrowserBuilder/'*|'Assembler/Entries/Browsers/'*)
      gen_browsers=1; gen_count=$((gen_count + 1)) ;;
    'Assembler/Generated/UWPBuilder/'*.ini)
      base=${f##*/}; letter=${base%.ini}
      gen_count=$((gen_count + 1))
      case "$letter" in
        # The 27 buckets UWPBuilder -split writes. Anything else (the single-file
        # uwp.ini UWPBuilder wrote before -split) takes the whole-folder rule.
        '#'|[A-Z]) gen_uwp_letters="$gen_uwp_letters $letter" ;;
        *)         gen_uwp="$gen_uwp $f" ;;
      esac ;;
    'Assembler/Generated/UWPBuilder/'*|'Assembler/Entries/UWP/'*)
      gen_uwp="$gen_uwp $f"; gen_count=$((gen_count + 1)) ;;
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
    # The 14 release data files, named one by one so nothing else in Releases/ counts.
    'Releases/winapp2.ini'|'Releases/winapp2-ccleaner.ini'|'Releases/winapp2-ccleaner7.ini'|'Releases/winapp2-bleachbit.ini'|'Releases/winapp2-fluentcleaner.ini'|'Releases/winapp2-tron.ini'|'Releases/winapp2-systemninja.rules')
      gen_count=$((gen_count + 1)); gen_other="$gen_other $f" ;;
    'Releases/changelog-base.txt'|'Releases/changelog-ccleaner.txt'|'Releases/changelog-ccleaner7.txt'|'Releases/changelog-bleachbit.txt'|'Releases/changelog-fluentcleaner.txt'|'Releases/changelog-tron.txt'|'Releases/changelog-systemninja.txt')
      gen_count=$((gen_count + 1)); gen_other="$gen_other $f" ;;
    # Their deprecated copies, published until the deprecation window closes.
    'Winapp2.ini'|'diff.txt'|'Non-CCleaner/'*'Winapp2.ini'|'Non-CCleaner/'*'Winapp2.rules'|'Non-CCleaner/'*'diff.txt')
      gen_count=$((gen_count + 1)); gen_other="$gen_other $f" ;;

    # ---- per-artifact sources.
    # The bare Assembler/EntryBuilder/, BrowserBuilder/ and UWP/ patterns are the pre-move
    # paths (now under Assembler/Sources/, UWP/ as UWPBuilder/), kept so a branch made before
    # the move that commits a source with its regeneration still matches them up.
    'Assembler/Sources/EntryBuilder/'*.ini|'Assembler/EntryBuilder/'*.ini)
      src_count=$((src_count + 1))
      base=${f##*/}
      src_entrybuilder="$src_entrybuilder ${base%.ini}" ;;
    'Assembler/Sources/BrowserBuilder/'*|'Assembler/BrowserBuilder/'*)
      src_count=$((src_count + 1)); src_browserbuilder=1 ;;
    'Assembler/Sources/UWPBuilder/AppInfo/'*.ini|'Assembler/UWP/AppInfo/'*.ini)
      src_count=$((src_count + 1)); src_uwp=1
      base=${f##*/}; letter=${base%.ini}
      case "$f" in
        # Only a top-level AppInfo/<L>.ini feeds one letter.
        */AppInfo/"$base") ;;
        *) letter="" ;;
      esac
      case "$letter" in
        '#'|[A-Z]) src_uwp_letters="$src_uwp_letters $letter" ;;
        *)         src_uwp_global=1 ;;
      esac ;;
    'Assembler/Sources/UWPBuilder/'*|'Assembler/UWP/'*)
      src_count=$((src_count + 1)); src_uwp=1; src_uwp_global=1 ;;

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

if [ "$gen_count" -eq 0 ]; then
  echo "No generated artifacts changed; nothing to check."
  exit 0
fi

# ------------------------------------------------------------------------------ verdict

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

if [ "$src_count" -eq 0 ]; then
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
  for letter in $gen_uwp_letters; do
    annotate "Assembler/Generated/UWPBuilder/$letter.ini" \
      "generated by UWPBuilder; edit Assembler/Sources/UWPBuilder/AppInfo/$letter.ini instead"
    note "- \`Assembler/Generated/UWPBuilder/$letter.ini\` <- \`Assembler/Sources/UWPBuilder/AppInfo/$letter.ini\`"
  done
  for f in $gen_uwp; do
    annotate "$f" \
      "generated by UWPBuilder; edit the sources in Assembler/Sources/UWPBuilder/ instead"
    note "- \`$f\` <- \`Assembler/Sources/UWPBuilder/\`"
  done
  for f in $gen_other; do
    annotate "$f" "build output; edit the sources under Assembler/ instead"
    note "- \`$f\` <- the build (\`Assembler/build winapp2.ps1\`)"
  done

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
  # A UWPBuilder source outside AppInfo/<L>.ini (UWP.ini) can move every UWP letter.
  if [ "$src_uwp_global" -eq 0 ]; then
    for letter in $gen_uwp_letters; do
      matched=0
      for s in $src_uwp_letters; do
        [ "$s" = "$letter" ] && { matched=1; break; }
      done
      if [ "$matched" -eq 0 ]; then
        fail=1
        orphans=1
        annotate "Assembler/Generated/UWPBuilder/$letter.ini" \
          "changed, but its source Assembler/Sources/UWPBuilder/AppInfo/$letter.ini did not, so the next build will overwrite this edit"
        note "- \`Assembler/Generated/UWPBuilder/$letter.ini\` changed, but \`Assembler/Sources/UWPBuilder/AppInfo/$letter.ini\` did not"
      fi
    done
  fi
  if [ -n "$gen_uwp" ] && [ "$src_uwp" -eq 0 ]; then
    fail=1; orphans=1
    for f in $gen_uwp; do
      annotate "$f" \
        "changed, but no source under Assembler/Sources/UWPBuilder/ did, so the next build will overwrite this edit"
      note "- \`$f\` changed, but nothing under \`Assembler/Sources/UWPBuilder/\` did"
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
    echo "Everything under \`Assembler/Generated/\`, the published files in \`Releases/\`, and"
    echo "their old \`Winapp2.ini\` and \`diff.txt\` copies are **build output**, produced by"
    echo "\`Assembler/build winapp2.ps1\`. Edit the sources and let the build write the output."
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
