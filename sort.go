package xlsx

import (
	"fmt"
	"regexp"
)

const (
	sortDesc = 1
	sortAsc  = 2
)

type Sort struct {
	Column    string
	Direction int
}

func (s Sort) DescendingAsString() string {
	if s.Direction == sortDesc {
		return "1"
	}

	return "0" // or maybe 2, not sure
}

// Range in column to do sorting e.g. N1:N20
func (s Sort) Range(topLeftCell, topRightCell string) string {
	re := regexp.MustCompile(`^[a-zA-Z]+(\d+)$`)
	start := re.ReplaceAllString(topLeftCell, s.Column+"$1")
	end := re.ReplaceAllString(topRightCell, s.Column+"$1")
	return fmt.Sprintf("%s:%s", start, end)
}
