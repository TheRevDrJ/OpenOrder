#!/bin/sh
# fanout_win.sh - send the Windows artifact from the build machine to its two
# standing destinations, and PROVE each copy arrived intact.
#
# Run ON THE MAC (it brokers between the two Windows boxes):
#   sh scripts/fanout_win.sh
#
# A Windows build has two standing destinations (build_rules.md, PC section),
# and both were being done by hand off a runbook on FirstLight until a build
# report went out claiming a version that was not actually installed. Scripted
# here so it happens whether anyone remembers it or not.
#
# AND IT HASHES. Copying is the leg where corruption actually happens - a
# truncated scp, a half-written file over a flaky link - and a damaged binary
# looks exactly like a code bug when he runs it. Every copy is verified SHA256
# against the source on the build machine, and a mismatch aborts loudly rather
# than handing him something to waste an evening on.
#
# WHY A ZIP AND NOT THE FOLDER: PyInstaller onedir is a directory of hundreds
# of files. Hashing each across the wire is impractical, and a partial
# directory copy is indistinguishable from a code bug at runtime. One archive
# means ONE hash answers "did the whole artifact arrive intact".
#
# Hostnames are roles; they live in rules_project.md:
#   zora    = the Windows BUILD machine
#   phoenix = the Windows TEST rig
#   Defiant = this Mac, his main machine - he wants the artifact in hand here too
set -e

# TWO forms of the same path, deliberately. PowerShell wants backslashes; scp
# parses the path itself and needs FORWARD slashes - a backslash path silently
# becomes "No such file or directory".
ZORA_ZIP='d:\claude\openorder-stage\OpenOrder-win.zip'
ZORA_ZIP_SCP='d:/claude/openorder-stage/OpenOrder-win.zip'
PHX='jonathan@phoenix.local'
PHX_DIR='c:\claude\OpenOrder-win'
MAC_STAGE="$HOME/claude-build/openorder-stage"

VER=$(node -p "require('./frontend/package.json').version")
N=$(cat build-number.txt)

fail() { echo "ABORT: $*" >&2; exit 1; }

mkdir -p "$MAC_STAGE"

# -- the source of truth, read off the build machine before anything moves ---
SRC_HASH=$(ssh zora "(Get-FileHash '$ZORA_ZIP' -Algorithm SHA256).Hash" | tr -d '\r' | tr 'a-z' 'A-Z')
[ -n "$SRC_HASH" ] || fail "no OpenOrder-win.zip on zora at $ZORA_ZIP - was the build run?"
echo "  source $SRC_HASH"

# -- leg 1: the Windows test rig --------------------------------------------
# scp -3 routes host-to-host without staging through the Mac.
echo "-- phoenix --------------------------------------------"
scp -3 -q "zora:$ZORA_ZIP_SCP" "$PHX:c:/claude/OpenOrder-win.zip" \
  || fail "copy to phoenix failed."

PHX_HASH=$(ssh "$PHX" "(Get-FileHash 'c:\\claude\\OpenOrder-win.zip' -Algorithm SHA256).Hash" | tr -d '\r' | tr 'a-z' 'A-Z')
[ "$PHX_HASH" = "$SRC_HASH" ] || fail "the phoenix copy is CORRUPT - got $PHX_HASH, expected $SRC_HASH"
echo "  copy   $PHX_HASH  (matches)"

# Unpack. A RUNNING OpenOrder write-locks its own exe on Windows, so this can
# fail with the previous build open on his screen. Abort and SAY so - never
# kill a process on a machine he might be using (rules.md sec 12).
ssh "$PHX" "if (Test-Path '$PHX_DIR') { Remove-Item -Recurse -Force '$PHX_DIR' }; Expand-Archive -Path 'c:\\claude\\OpenOrder-win.zip' -DestinationPath '$PHX_DIR' -Force" \
  > /dev/null 2>&1 \
  || fail "could not unpack on phoenix. If OpenOrder is RUNNING there it holds its own exe open - ask Jonathan to close it, then run this again."

PHX_EXE=$(ssh "$PHX" "if (Test-Path '$PHX_DIR\\OpenOrder.exe') { 'yes' }" | tr -d '\r\n')
[ "$PHX_EXE" = "yes" ] || fail "unpacked on phoenix but $PHX_DIR\\OpenOrder.exe is not there."
printf '  unpacked: %s\\OpenOrder.exe\n' "$PHX_DIR"

# -- leg 2: his Mac ---------------------------------------------------------
echo "-- Defiant (this Mac) ---------------------------------"
scp -q "zora:$ZORA_ZIP_SCP" "$MAC_STAGE/OpenOrder-win.zip" || fail "copy back to the Mac failed."
MAC_HASH=$(shasum -a 256 "$MAC_STAGE/OpenOrder-win.zip" | awk '{print $1}' | tr 'a-z' 'A-Z')
[ "$MAC_HASH" = "$SRC_HASH" ] || fail "the Mac copy is CORRUPT - got $MAC_HASH, expected $SRC_HASH"
echo "  copy   $MAC_HASH  (matches)"
echo "  at     $MAC_STAGE/OpenOrder-win.zip"

echo
echo "FAN-OUT COMPLETE - v$VER build $N on phoenix and on this Mac, both SHA256-verified."
