package xlsx

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

const (
	SortDesc       = 1
	SortAsc        = 2
	SortAllStrings = "all-strings"
	SortAllDollars = "all-dollars"
)

// SortStrategy for sheet
type SortStrategy struct {
	Column           string
	ColumnIndex      int
	ColumnValuesType string
	Direction        int
	HasHeader        bool
}

func (s SortStrategy) descendingAsString() string {
	if s.Direction == SortDesc {
		return "1"
	}

	return "0" // or maybe 2, not sure
}

func (s *SortStrategy) stateRange(topLeftCell, topRightCell string) string {
	re := regexp.MustCompile(`^([a-zA-Z]+)\d+$`)
	start := re.ReplaceAllString(topLeftCell, "${1}2")
	end := topRightCell
	return fmt.Sprintf("%s:%s", start, end)
}

// conditionRange in column to do sorting e.g. N1:N20
func (s *SortStrategy) conditionRange(topLeftCell, topRightCell string) string {
	re := regexp.MustCompile(`^[a-zA-Z]+(\d+)$`)
	start := re.ReplaceAllString(topLeftCell, s.Column+"$1")
	end := re.ReplaceAllString(topRightCell, s.Column+"$1")
	return fmt.Sprintf("%s:%s", start, end)
}

func (s *SortStrategy) shouldSwapRows(rowA, rowB *Row) bool {
	switch s.ColumnValuesType {
	case SortAllStrings:
		if s.Direction == SortDesc {
			return strings.ToLower(rowA.Cells[s.ColumnIndex].String()) < strings.ToLower(rowB.Cells[s.ColumnIndex].String())
		}

		return strings.ToLower(rowA.Cells[s.ColumnIndex].String()) > strings.ToLower(rowB.Cells[s.ColumnIndex].String())
	case SortAllDollars:
		if s.Direction == SortDesc {
			return cellWithDollarsToFloat(rowA.Cells[s.ColumnIndex]) > cellWithDollarsToFloat(rowB.Cells[s.ColumnIndex])
		}

		return cellWithDollarsToFloat(rowA.Cells[s.ColumnIndex]) > cellWithDollarsToFloat(rowB.Cells[s.ColumnIndex])
	default:
		return false
	}
}

func cellWithDollarsToFloat(c *Cell) float64 {
	stripped := strings.Trim(strings.Replace(c.String(), "$", "", 1), " ")
	if f, err := strconv.ParseFloat(stripped, 64); err == nil {
		fmt.Println(f)
		return f
	}

	return 0
}
