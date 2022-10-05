package yymmdd

import (
	"strconv"
	"strings"
	"unicode/utf8"
)

const (
	tYEAR   = "T_YEAR_MARK"
	tMONTH  = "T_MONTH_MARK"
	tDAY    = "T_DAY_MARK"
	tHOUR   = "T_HOUR_MARK"
	tMINUTE = "T_MINUTE_MARK"
	tSECOND = "T_SECOND_MARK"
	tRAW    = "T_RAW_MARK"
	tEOF    = "T_EOF"
)

// LexToken holds is a (type, value) array.
type LexToken [3]string

// EOF character
var EOF string = "+++EOF+++"

// lexerState represents the state of the scanner
// as a function that returns the next state.
type lexerState func(*lexer) lexerState

// run lexes the input by executing state functions until
// the state is nil.
func (l *lexer) Run() {
	for state := l.initialState; state != nil; {
		state = state(l)
	}
}

// lexLayout creates a new scanner for the input string.
func lexLayout(input string) (*lexer, []LexToken) {
	l := &lexer{
		input:  input,
		tokens: make([]LexToken, 0),
		lineno: 1,
	}
	l.initialState = initLexerState
	l.Run()
	return l, l.tokens
}

// lexer holds the state of the scanner.
type lexer struct {
	input        string     // the string being scanned.
	start        int        // start position of this item.
	pos          int        // current position in the input.
	width        int        // width of last rune read from input.
	tokens       []LexToken // scanned items.
	initialState lexerState

	lineno int
}

// next returns the next rune in the input.
func (l *lexer) next() string {
	var r rune
	if l.pos >= len(l.input) {
		l.width = 0
		return EOF
	}
	r, l.width = utf8.DecodeRuneInString(l.input[l.pos:])
	l.pos += l.width
	return string(r)
}

// ignore skips over the pending input before this point.
func (l *lexer) ignore() {
	l.start = l.pos
}

// backup steps back one rune.
// Can be called only once per call of next.
func (l *lexer) backup() {
	l.pos -= l.width
}

// acceptRun consumes a run of runes from the valid set.
func (l *lexer) acceptRun(valid string) {
	for strings.Index(valid, l.next()) >= 0 {
	}
	l.backup()
}

// emit passes an item back to the client.
func (l *lexer) emit(t string) {
	l.tokens = append(l.tokens, LexToken{t, l.input[l.start:l.pos], strconv.Itoa(l.lineno)})
	l.start = l.pos
}

// emit passes an item back to the client.
func (l *lexer) emitRaw() {
	if l.pos-l.start > 1 {
		l.tokens = append(l.tokens, LexToken{tRAW, l.input[l.start : l.pos-1], strconv.Itoa(l.lineno)})
		l.start = l.pos - 1
	}
}

// initialState is the starting point for the
// scanner. It scans through each character and decides
// which state to create for the lexer. lexerState == nil
// is exit scanner.
func initLexerState(l *lexer) lexerState {
	for r := l.next(); r != EOF; r = l.next() {
		switch r {
		case "y":
			l.emitRaw()
			l.acceptRun("y")
			l.emit(tYEAR)
		case "Y":
			l.emitRaw()
			l.acceptRun("Y")
			l.emit(tYEAR)
		case "M":
			l.emitRaw()
			l.acceptRun("M")
			l.emit(tMONTH)
		case "d":
			l.emitRaw()
			l.acceptRun("d")
			l.emit(tDAY)
		case "D":
			l.emitRaw()
			l.acceptRun("D")
			l.emit(tDAY)
		case "h":
			l.emitRaw()
			l.acceptRun("h")
			l.emit(tHOUR)
		case "m":
			l.emitRaw()
			l.acceptRun("m")
			l.emit(tMINUTE)
		case "s":
			l.emitRaw()
			l.acceptRun("s")
			l.emit(tSECOND)
		}
	}

	l.emit(tEOF)
	return nil
}
