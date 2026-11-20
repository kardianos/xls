package fuzzcap

import (
	"bufio"
	"os"
	"path"
	"strconv"
	"strings"
	"syscall"
)

// memoryLimit returns the smallest of the address-space limit and the
// memory.max of this process's cgroup v2 and its ancestors.
func memoryLimit() (uint64, bool) {
	var limit uint64
	found := false
	take := func(v uint64) {
		if !found || v < limit {
			limit, found = v, true
		}
	}
	var rl syscall.Rlimit
	if syscall.Getrlimit(syscall.RLIMIT_AS, &rl) == nil && rl.Cur != ^uint64(0) {
		take(rl.Cur)
	}
	b, err := os.ReadFile("/proc/self/cgroup")
	if err != nil {
		return limit, found
	}
	for _, line := range strings.Split(string(b), "\n") {
		p, ok := strings.CutPrefix(line, "0::")
		if !ok {
			continue
		}
		for dir := path.Clean(p); ; dir = path.Dir(dir) {
			if v, err := os.ReadFile(path.Join("/sys/fs/cgroup", dir, "memory.max")); err == nil {
				if n, err := strconv.ParseUint(strings.TrimSpace(string(v)), 10, 64); err == nil {
					take(n)
				}
			}
			if dir == "/" || dir == "." {
				break
			}
		}
	}
	return limit, found
}

func physicalMemory() (uint64, bool) {
	f, err := os.Open("/proc/meminfo")
	if err != nil {
		return 0, false
	}
	defer f.Close()
	s := bufio.NewScanner(f)
	for s.Scan() {
		if rest, ok := strings.CutPrefix(s.Text(), "MemTotal:"); ok {
			kb, err := strconv.ParseUint(strings.TrimSpace(strings.TrimSuffix(strings.TrimSpace(rest), "kB")), 10, 64)
			return kb << 10, err == nil
		}
	}
	return 0, false
}
