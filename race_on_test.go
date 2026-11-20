//go:build race

package xls

// slow scales time limits in tests: the race detector slows code about tenfold.
const slow = 10
