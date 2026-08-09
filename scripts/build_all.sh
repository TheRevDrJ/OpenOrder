#!/bin/sh
# build_all.sh - THE joint build. Mac and Windows, one event, one number.
#
#   sh scripts/build_all.sh
#
# HIS RULING (2026-08-09): "Both every time." Mac and PC build together on
# every change, never one without the other unless he says so for a specific
# reason. The reason it is one script and not two habits is rules.md sec 20 -
# a step that lives only in "and then I also run the other one" does not
# exist.
#
# ONE NUMBER PER EVENT. build.sh (the mac leg, first) owns the single bump;
# the windows leg scp's that same uncommitted counter to zora and READS it.
# Both artifacts wear the same number, which is what lets us answer "which
# binary is he running" when he reports something.
#
# The order is load-bearing:
#   1. mac      bumps the counter, installs to /Applications
#   2. windows  reads that counter, builds on zora
#   3. fan-out  zip to phoenix and back here, SHA256 on every copy
#   4. lock     commit the counter and push
#
# Step 4 is the counter ONLY. Tagging a release is a judgment (it happens on a
# version change, not every build) and stays a human decision - see rules.md
# sec 4 and sec 20's second counterweight.
set -e

cd "$(dirname "$0")/.."

echo "=================================================="
echo " JOINT BUILD - mac + windows"
echo "=================================================="
echo

# -- gates ------------------------------------------------------------------
echo "[gates] typecheck ..."
( cd frontend && npx tsc -b --noEmit )
echo "        clean."

DIRT=$(git status --porcelain | awk '{print $2}' | grep -v -x -e 'build-number.txt' || true)
if [ -n "$DIRT" ]; then
  echo "ABORT: uncommitted changes - commit and push before a build (build_rules.md 1.4):" >&2
  echo "$DIRT" >&2
  exit 1
fi
echo "        tree clean."
echo

# -- 1. mac (owns the bump) -------------------------------------------------
sh ./build.sh
echo

# -- 2. windows -------------------------------------------------------------
sh scripts/build_win_remote.sh
echo

# -- 3. fan-out -------------------------------------------------------------
sh scripts/fanout_win.sh
echo

# -- 4. lock the number -----------------------------------------------------
N=$(cat build-number.txt)
VER=$(node -p "require('./frontend/package.json').version")
if [ -n "$(git status --porcelain build-number.txt)" ]; then
  echo "[lock] committing the number ..."
  git add build-number.txt
  git commit -q -m "build $N - the number"
  git push -q
  git push -q zora master
  echo "       build $N committed and pushed (origin + zora)."
fi

echo
echo "=================================================="
echo " JOINT BUILD COMPLETE - v$VER build $N"
echo "   mac      /Applications/OpenOrder.app"
echo "   windows  phoenix c:\\claude\\OpenOrder-win\\  + ~/claude-build/openorder-stage/"
echo "=================================================="
