//go:build race

package yymmdd

// slow scales time limits in tests: the race detector slows code about tenfold.
const slow = 10
