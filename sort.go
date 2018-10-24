package xlsx

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
)

const (
	SortDesc           = 1
	SortAsc            = 2
	SortAllStrings     = "all-strings"
	SortAllDollars     = "all-dollars"
	SortAllPercentages = "all-percentages"
)

// SortStrategy for sheet
type SortStrategy struct {
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
	repl := "${1}2"

	if !s.HasHeader {
		repl = "${1}1"
	}

	start := re.ReplaceAllString(topLeftCell, repl)
	end := topRightCell
	return fmt.Sprintf("%s:%s", start, end)
}

// conditionRange in column to do sorting e.g. N1:N20
func (s *SortStrategy) conditionRange(topLeftCell, topRightCell string) string {
	re := regexp.MustCompile(`^[a-zA-Z]+(\d+)$`)
	column := resolveColumn(s.ColumnIndex)
	start := re.ReplaceAllString(topLeftCell, column+"$1")
	end := re.ReplaceAllString(topRightCell, column+"$1")
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
	case SortAllPercentages:
		if s.Direction == SortDesc {
			return cellWithPercentageToFloat(rowA.Cells[s.ColumnIndex]) > cellWithPercentageToFloat(rowB.Cells[s.ColumnIndex])
		}

		return cellWithPercentageToFloat(rowA.Cells[s.ColumnIndex]) > cellWithPercentageToFloat(rowB.Cells[s.ColumnIndex])
	default:
		return false
	}
}

func cellWithDollarsToFloat(c *Cell) float64 {
	stripped := strings.Trim(strings.Replace(c.String(), "$", "", 1), " ")
	if f, err := strconv.ParseFloat(stripped, 64); err == nil {
		return f
	}

	return 0
}

func cellWithPercentageToFloat(c *Cell) float64 {
	stripped := strings.Trim(strings.Replace(c.String(), "%", "", 1), " ")
	if f, err := strconv.ParseFloat(stripped, 64); err == nil {
		return f
	}

	return 0
}

func resolveColumn(number int) string {
	output := ""
	remainder := (number) % 26
	prefix := int(math.Floor(float64(number / 26)))
	if prefix > 0 {
		output = output + string('A'+prefix-1)
	}
	return output + string('A'+remainder)
}
