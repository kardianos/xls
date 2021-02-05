package yymmdd

import "time"

func Parse(value, format string) (time.Time, error) {
	_, tokens := lexLayout(format)
	ds := parse(tokens)
	return ds.Parse(value)
}

// Parse creates a new parser with the recommended
// parameters.
func parse(tokens []LexToken) formatter {
	p := &parser{
		tokens: tokens,
		pos:    -1,
	}
	p.initState = initialParserState
	return p.run()
}

// run starts the statemachine
func (p *parser) run() formatter {
	var f formatter
	for state := p.initState; state != nil; {
		state = state(p, &f)
	}
	return f
}

// parserState represents the state of the scanner
// as a function that returns the next state.
type parserState func(*parser, *formatter) parserState

// nest returns what the next token AND
// advances p.pos.
func (p *parser) next() *LexToken {
	if p.pos >= len(p.tokens)-1 {
		return nil
	}
	p.pos += 1
	return &p.tokens[p.pos]
}

// the parser type
type parser struct {
	tokens []LexToken
	pos    int
	serial int

	initState parserState
}

// the starting state for parsing
func initialParserState(p *parser, f *formatter) parserState {
	var t *LexToken
	for t = p.next(); t[0] != tEOF; t = p.next() {
		var item ItemFormatter
		switch t[0] {
		case tYEAR:
			item = new(YearFormatter)
		case tMONTH:
			item = new(MonthFormatter)
		case tDAY:
			item = new(DayFormatter)
		case tHOUR:
			item = new(MonthFormatter)
		case tMINUTE:
			item = new(MonthFormatter)
		case tSECOND:
			item = new(MonthFormatter)
		case tRAW:
			item = new(basicFormatter)
		}
		item.setOriginal(t[1])
		f.Items = append(f.Items, item)
	}
	if len(t[1]) > 0 {
		r := new(basicFormatter)
		r.origin = t[1]
		f.Items = append(f.Items, r)
	}
	return nil
}
