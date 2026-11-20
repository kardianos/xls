package xls

import (
	"math/bits"
	"sort"
)

// maxRows is the number of rows a BIFF5 or BIFF8 sheet can have.
const maxRows = 1 << 16

// ref is a cell range stored compactly: rows and columns, inclusive.
type ref struct{ r0, r1, c0, c1 uint16 }

func toRef(c CellRange) ref {
	return ref{uint16(c.FirstRow), uint16(c.LastRow), uint16(c.FirstCol), uint16(c.LastCol)}
}

func (r ref) cellRange() CellRange {
	return CellRange{FirstRow: int(r.r0), LastRow: int(r.r1), FirstCol: int(r.c0), LastCol: int(r.c1)}
}

// rangeIndex finds a cell range that contains a cell in O(log n) time. It
// is a segment tree over rows: each range is stored in the at most 32 nodes
// that exactly cover its rows. A node's ranges are kept sorted by first
// column, with a running record of the range that reaches furthest right,
// so a lookup is a binary search in each node on the path from the row's
// leaf to the root. Overlapping ranges are handled; which one is returned
// is then unspecified.
//
// Building the index takes time and memory linear in the number of
// (range, node) pairs.
type rangeIndex struct {
	ranges []ref
	idx    []int32             // range numbers, grouped by node, sorted by first column
	reach  []int32             // reach[k]: of idx[group start..k], the range reaching furthest right
	sparse map[uint32][2]int32 // node -> [start, end) in idx
	dense  [][2]int32          // the same, indexed by node, for many ranges
	levels uint32              // bit k is set if some node at depth k holds a range
	rows   [][2]int32          // sorted, disjoint row spans that some range covers
}

// denseMin is the number of ranges from which a dense node table (1 MiB)
// is used instead of a map.
const denseMin = 1 << 15

func newRangeIndex(ranges []ref) rangeIndex {
	x := rangeIndex{ranges: ranges}
	if len(ranges) == 0 {
		return x
	}
	// Order the ranges by first column with a counting sort.
	var count [maxCols + 1]int32
	for _, r := range ranges {
		count[min(int(r.c0), maxCols-1)+1]++
	}
	for c := 1; c <= maxCols; c++ {
		count[c] += count[c-1]
	}
	order := make([]int32, len(ranges))
	for i, r := range ranges {
		c := min(int(r.c0), maxCols-1)
		order[count[c]] = int32(i)
		count[c]++
	}

	// Pairs of (node, range), in first-column order.
	type pair struct {
		node uint32
		r    int32
	}
	pairs := make([]pair, 0, 2*len(ranges))
	for _, i := range order {
		r := ranges[i]
		lo, hi := uint32(r.r0)+maxRows, uint32(r.r1)+1+maxRows
		for lo < hi {
			if lo&1 == 1 {
				pairs = append(pairs, pair{lo, i})
				lo++
			}
			if hi&1 == 1 {
				hi--
				pairs = append(pairs, pair{hi, i})
			}
			lo >>= 1
			hi >>= 1
		}
	}

	// Group the pairs by node with a stable radix sort (node ids have 17
	// bits), keeping first-column order within each node.
	tmp := make([]pair, len(pairs))
	for _, shift := range []uint{0, 9} {
		var c [513]int32
		for _, p := range pairs {
			c[(p.node>>shift)&511+1]++
		}
		for k := 1; k < len(c); k++ {
			c[k] += c[k-1]
		}
		for _, p := range pairs {
			b := (p.node >> shift) & 511
			tmp[c[b]] = p
			c[b]++
		}
		pairs, tmp = tmp, pairs
	}

	x.idx = make([]int32, len(pairs))
	x.reach = make([]int32, len(pairs))
	if len(ranges) >= denseMin {
		x.dense = make([][2]int32, 2*maxRows)
	} else {
		x.sparse = make(map[uint32][2]int32)
	}
	for k := 0; k < len(pairs); {
		n := pairs[k].node
		start := k
		best := pairs[k].r
		for ; k < len(pairs) && pairs[k].node == n; k++ {
			i := pairs[k].r
			x.idx[k] = i
			if ranges[i].c1 > ranges[best].c1 {
				best = i
			}
			x.reach[k] = best
		}
		g := [2]int32{int32(start), int32(k)}
		if x.dense != nil {
			x.dense[n] = g
		} else {
			x.sparse[n] = g
		}
		x.levels |= 1 << (bits.Len32(n) - 1)
	}

	// Rows that some range covers, as sorted disjoint spans. Many ranges
	// are sorted by first row with a counting sort, a few with sort.Slice,
	// so a sheet with a handful of ranges does not pay for a full table.
	byRow := make([]int32, len(ranges))
	if len(ranges) >= 1<<12 {
		starts := make([]int32, maxRows+1)
		for _, r := range ranges {
			starts[int(r.r0)+1]++
		}
		for k := 1; k <= maxRows; k++ {
			starts[k] += starts[k-1]
		}
		for i, r := range ranges {
			byRow[starts[r.r0]] = int32(i)
			starts[r.r0]++
		}
	} else {
		for i := range byRow {
			byRow[i] = int32(i)
		}
		sort.Slice(byRow, func(a, b int) bool { return ranges[byRow[a]].r0 < ranges[byRow[b]].r0 })
	}
	for _, i := range byRow {
		r0, r1 := int32(ranges[i].r0), int32(ranges[i].r1)
		if n := len(x.rows); n > 0 && r0 <= x.rows[n-1][1]+1 {
			x.rows[n-1][1] = max(x.rows[n-1][1], r1)
			continue
		}
		x.rows = append(x.rows, [2]int32{r0, r1})
	}
	return x
}

// find returns the index of a range that contains the cell, or -1.
func (x *rangeIndex) find(row, col int) int {
	if x.idx == nil || row < 0 || row >= maxRows || col < 0 {
		return -1
	}
	// Rows no range covers, the common case, need no tree walk.
	lo, hi := 0, len(x.rows)
	for lo < hi {
		m := int(uint(lo+hi) >> 1)
		if x.rows[m][1] < int32(row) {
			lo = m + 1
		} else {
			hi = m
		}
	}
	if lo == len(x.rows) || x.rows[lo][0] > int32(row) {
		return -1
	}
	for n := uint32(row) + maxRows; n >= 1; n >>= 1 {
		if x.levels&(1<<(bits.Len32(n)-1)) == 0 {
			continue
		}
		var g [2]int32
		if x.dense != nil {
			g = x.dense[n]
		} else {
			g = x.sparse[n]
		}
		// The last range in the group starting at or before col.
		lo, hi := g[0], g[1]
		for lo < hi {
			m := int32(uint32(lo+hi) >> 1)
			if int(x.ranges[x.idx[m]].c0) <= col {
				lo = m + 1
			} else {
				hi = m
			}
		}
		if lo > g[0] {
			// Every range in the group covers the row; the one reaching
			// furthest right contains col if any of them does.
			if i := x.reach[lo-1]; int(x.ranges[i].c1) >= col {
				return int(i)
			}
		}
	}
	return -1
}
