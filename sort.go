package xlsx

import (
	"fmt"
	"regexp"
)

const (
	SortDesc = 1
	SortAsc  = 2
)

type Sort struct {
	Column    string
	Direction int
}

func (s Sort) DescendingAsString() string {
	if s.Direction == SortDesc {
		return "1"
	}

	return "0" // or maybe 2, not sure
}

func (s Sort) StateRange(topLeftCell, topRightCell string) string {
	re := regexp.MustCompile(`^([a-zA-Z]+)\d+$`)
	start := re.ReplaceAllString(topLeftCell, "${1}2")
	end := topRightCell
	fmt.Printf("%s:%s", start, end)
	return fmt.Sprintf("%s:%s", start, end)
}

// ConditionRange in column to do sorting e.g. N1:N20
func (s Sort) ConditionRange(topLeftCell, topRightCell string) string {
	re := regexp.MustCompile(`^[a-zA-Z]+(\d+)$`)
	start := re.ReplaceAllString(topLeftCell, s.Column+"$1")
	end := re.ReplaceAllString(topRightCell, s.Column+"$1")
	return fmt.Sprintf("%s:%s", start, end)
}
