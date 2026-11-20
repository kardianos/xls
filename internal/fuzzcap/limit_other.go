//go:build !linux

package fuzzcap

// The limit cannot be detected on this system; set XLS_FUZZ_UNCAPPED=1 after
// arranging a memory cap another way.
func memoryLimit() (uint64, bool)    { return 0, false }
func physicalMemory() (uint64, bool) { return 0, false }
