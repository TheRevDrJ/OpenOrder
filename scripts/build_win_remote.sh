#!/bin/sh
# build_win_remote.sh - the WINDOWS leg of a joint build, as ONE command.
# Run on the Mac. Drives Zora end to end: push, narrow clean, pull, counter,
# build, receipt. Nothing about this leg is composed by hand.
#
# Modelled directly on FirstLight's script of the same name, which exists
# because a hand-composed leg lost a receipt to its own `tail -25` pipe. The
# lesson there is the one that governs here (rules.md sec 20): a step that
# lives only in "how Bob usually does it" does not exist, because Bob does not
# carry habits across sessions.
#
# WHAT IT ENCODES:
#   . a NARROW clean on the mirror, never a blanket `git checkout -- .`
#     (a blanket revert would also eat the counter we are about to send)
#   . the counter scp STRICTLY AFTER the pull - the ordering that once put
#     four FirstLight artifacts under two numbers
#   . a read-back of HEAD and the counter on the FAR side before building
#   . a FULL log to disk, never a pipe, so no receipt can be truncated away
#   . a fail-closed receipt check: no RECEIPT line, no success
set -e

ZORA_REPO='d:\claude\OpenOrder'
LOGDIR="$HOME/claude-build/openorder-logs"   # outside Dropbox: regenerable
mkdir -p "$LOGDIR"

fail() { echo "ABORT: $*" >&2; exit 1; }

# -- gate: the build photographs the WORKING TREE, so no uncommitted SOURCE --
#
# NOT a blanket clean-tree check. By the time this leg runs, the mac leg has
# necessarily bumped build-number.txt and that bump is committed only AFTER
# the whole event - so a blanket gate could never pass. Allow exactly that one
# known side-effect and abort on anything else.
DIRT=$(git status --porcelain | awk '{print $2}' | grep -v -x -e 'build-number.txt' || true)
[ -z "$DIRT" ] || fail "uncommitted changes beyond the build's own side-effects - commit and push first (build_rules.md 1.4):
$DIRT"

LOCAL_HEAD=$(git rev-parse HEAD)
LOCAL_N=$(cat build-number.txt)
LOCAL_VER=$(node -p "require('./frontend/package.json').version")
echo "=== LOCAL: v$LOCAL_VER build $LOCAL_N @ $(git rev-parse --short HEAD) ==="

# -- 1. publish to the build box --------------------------------------------
echo "=== PUSH zora ==="
git push zora master

# -- 2. clean NARROWLY, then pull -------------------------------------------
# The mirror dirties build-number.txt every build (we scp a new value onto a
# tracked file). Reverting just that file lets the pull run clean; a blanket
# revert here would throw away the counter we are about to send.
echo "=== ZORA: narrow clean, then pull ==="
ssh zora "cd $ZORA_REPO; git checkout -- build-number.txt; git pull" \
  > "$LOGDIR/zora-pull.log" 2>&1 || { tail -20 "$LOGDIR/zora-pull.log" >&2; fail "the pull on Zora failed - see $LOGDIR/zora-pull.log"; }
tail -2 "$LOGDIR/zora-pull.log"

ZORA_HEAD=$(ssh zora "cd $ZORA_REPO; git rev-parse HEAD" | tr -d '\r\n')
[ "$ZORA_HEAD" = "$LOCAL_HEAD" ] || fail "Zora is on $ZORA_HEAD, the Mac is on $LOCAL_HEAD - the pull did not land the source we are building."

# -- 3. the counter, STRICTLY AFTER the pull --------------------------------
echo "=== COUNTER -> zora (after the pull, never batched with it) ==="
scp -q build-number.txt "zora:d:/claude/OpenOrder/build-number.txt"
ZORA_N=$(ssh zora "cd $ZORA_REPO; cat build-number.txt" | tr -d '\r\n')
[ "$ZORA_N" = "$LOCAL_N" ] || fail "Zora's counter reads '$ZORA_N', the Mac's reads '$LOCAL_N' - the scp did not land."
echo "  verified on the far side: build $ZORA_N"

# -- 4. build, FULL log to disk, receipt checked ----------------------------
# No pipes. The log IS the capture; we read it afterwards. A build whose
# RECEIPT line is missing is a FAILED build even if the exit code was 0,
# because a receipt we cannot read is a build we cannot identify.
LOG="$LOGDIR/win-build.log"
echo "=== ZORA BUILD ==="
ssh zora "cd $ZORA_REPO; powershell -ExecutionPolicy Bypass -File scripts\\build_win.ps1" \
  > "$LOG" 2>&1 || { tail -30 "$LOG" >&2; fail "the windows build failed - full log at $LOG"; }

RECEIPT=$(grep 'RECEIPT:' "$LOG" | tr -d '\r' | tail -1)
[ -n "$RECEIPT" ] || { tail -30 "$LOG" >&2; fail "the windows build printed no RECEIPT line - full log at $LOG"; }
echo "$RECEIPT" | grep -q "build $LOCAL_N" || fail "windows stamped the wrong number: $RECEIPT (expected build $LOCAL_N)"
echo "$RECEIPT" | grep -q "v$LOCAL_VER" || fail "windows stamped the wrong version: $RECEIPT (expected v$LOCAL_VER)"
echo "  $RECEIPT"
echo "  full log: $LOG"

echo
echo "WINDOWS LEG COMPLETE - v$LOCAL_VER build $LOCAL_N, receipt verified."
echo "Next: sh scripts/fanout_win.sh   (phoenix + this Mac, SHA256 on every copy)"
