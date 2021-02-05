package yymmdd

import (
	"errors"
	"fmt"
	"time"
)

func Format(time time.Time, layout string) string {
	_, tokens := lexLayout(layout)
	ds := parse(tokens)
	return ds.Format(time)
}

var errInvalidTimeFormat = errors.New("invalid time format")

type formatter struct {
	Items []ItemFormatter
}

func (f *formatter) Format(t time.Time) string {
	var gf string
	for _, i := range f.Items {
		itemFormatString, _ := i.translateToGolangFormat()
		gf += itemFormatString
	}
	return t.Format(gf)
}

func (f *formatter) Parse(str string) (time.Time, error) {
	var gf string
	for _, i := range f.Items {
		itemFormatString, _ := i.translateToGolangFormat()
		gf += itemFormatString
	}
	return time.Parse(gf, str)
}

type ItemFormatter interface {
	translateToGolangFormat() (string, error)
	setOriginal(string)
}

type basicFormatter struct {
	origin string
}

func (self *basicFormatter) translateToGolangFormat() (string, error) {
	return self.origin, nil
}

func (self *basicFormatter) setOriginal(o string) {
	self.origin = o
}

func (self *basicFormatter) String() string {
	return fmt.Sprintf("basic formatter as (%s)", self.origin)
}

type YearFormatter struct {
	basicFormatter
}

func (y *YearFormatter) translateToGolangFormat() (string, error) {
	switch len(y.origin) {
	case 4:
		return "2006", nil
	case 2:
		return "06", nil
	default:
		return "", errInvalidTimeFormat
	}
}

func (y *YearFormatter) String() string {
	return fmt.Sprintf("year formatter as (%s)", y.origin)
}

type MonthFormatter struct {
	basicFormatter
}

func (self *MonthFormatter) translateToGolangFormat() (string, error) {
	switch len(self.origin) {
	case 3:
		return "Jan", nil
	case 2:
		return "01", nil
	case 1:
		return "1", nil
	default:
		return "", errInvalidTimeFormat
	}
}

func (self *MonthFormatter) String() string {
	return fmt.Sprintf("month formatter as (%s)", self.origin)
}

type DayFormatter struct {
	basicFormatter
}

func (self *DayFormatter) translateToGolangFormat() (string, error) {
	switch len(self.origin) {
	case 2:
		return "02", nil
	case 1:
		return "2", nil
	default:
		return "", errInvalidTimeFormat
	}
}

func (self *DayFormatter) String() string {
	return fmt.Sprintf("day formatter as (%s)", self.origin)
}

type HourFormatter struct {
	basicFormatter
}

func (self *HourFormatter) translateToGolangFormat() (string, error) {
	switch len(self.origin) {
	case 2:
		return "15", nil
	case 1:
		return "15", nil
	default:
		return "", errInvalidTimeFormat
	}
}

func (self *HourFormatter) String() string {
	return fmt.Sprintf("hour formatter as (%s)", self.origin)
}

type MinuteFormatter struct {
	basicFormatter
}

func (self *MinuteFormatter) translateToGolangFormat() (string, error) {
	switch len(self.origin) {
	case 2:
		return "04", nil
	case 1:
		return "4", nil
	default:
		return "", errInvalidTimeFormat
	}
}

func (self *MinuteFormatter) String() string {
	return fmt.Sprintf("minute formatter as (%s)", self.origin)
}

type SecondFormatter struct {
	basicFormatter
}

func (self *SecondFormatter) translateToGolangFormat() (string, error) {
	switch len(self.origin) {
	case 2:
		return "05", nil
	case 1:
		return "5", nil
	default:
		return "", errInvalidTimeFormat
	}
}

func (self *SecondFormatter) String() string {
	return fmt.Sprintf("second formatter as (%s)", self.origin)
}
