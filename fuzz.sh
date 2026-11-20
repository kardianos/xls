#!/usr/bin/env bash
# Run Go fuzzing, or any go test command, under a memory cap.
#
# A fuzzer that finds an allocation bug can use all of the machine's memory.
# The fuzz targets in this repository refuse to fuzz unless they see a cap of
# at most min(16 GiB, half of RAM).
#
# Usage:
#   ./fuzz.sh [-m MEM] [-t TIME] [-p N] PKG TARGET   fuzz TARGET in PKG
#   ./fuzz.sh [-m MEM] -- GO-TEST-ARGS...            run "go test ARGS" capped
#
# Examples:
#   ./fuzz.sh . FuzzOpen
#   ./fuzz.sh -t 10m ./ole2 FuzzOpen
#   ./fuzz.sh -- -run 'FuzzOpen/0123abcd' .          replay a crasher
#
# MEM is the cap for the whole process tree (default 4G). On Linux with a
# systemd user session it is enforced as a cgroup limit (MemoryMax, no swap).
# Otherwise each process gets an address-space limit (ulimit -v) of MEM.
set -euo pipefail

mem=4G time=60s par=4
while getopts m:t:p: opt; do
	case $opt in
	m) mem=$OPTARG ;;
	t) time=$OPTARG ;;
	p) par=$OPTARG ;;
	*) sed -n '2,20s/^# \{0,1\}//p' "$0"; exit 2 ;;
	esac
done
# getopts consumes a "--" that ends the options; note whether it did.
passthrough=0
if ((OPTIND > 1)) && [[ ${*:OPTIND-1:1} == -- ]]; then
	passthrough=1
fi
shift $((OPTIND - 1))

if ((passthrough)); then
	cmd=(go test "$@")
elif [[ $# -eq 2 ]]; then
	cmd=(go test -run '^$' -fuzz "^$2\$" -fuzztime "$time" -parallel "$par" "$1")
else
	sed -n '2,20s/^# \{0,1\}//p' "$0"
	exit 2
fi

if command -v systemd-run >/dev/null && systemctl --user show-environment >/dev/null 2>&1; then
	# OOMPolicy=continue keeps the fuzz coordinator alive when a worker is
	# killed, so it can report the failure and save the input.
	# The inner shell checks that the cap really applies before running.
	exec systemd-run --user --scope --quiet --collect \
		-p MemoryMax="$mem" -p MemorySwapMax=0 -p OOMPolicy=continue -- \
		bash -c 'cg=$(sed -n "s/^0:://p" /proc/self/cgroup)
			max=$(cat "/sys/fs/cgroup$cg/memory.max" 2>/dev/null)
			if [[ -z $max || $max == max ]]; then
				echo "fuzz.sh: memory cap not applied (memory.max=${max:-unknown})" >&2
				exit 1
			fi
			exec "$@"' fuzz.sh "${cmd[@]}"
fi

kb=$(numfmt --from=iec "$mem")
kb=$((kb / 1024))
ulimit -v "$kb"
if [[ $(ulimit -v) != "$kb" ]]; then
	echo "fuzz.sh: memory cap not applied" >&2
	exit 1
fi
exec "${cmd[@]}"
