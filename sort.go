package xlsx

import (
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
)

type CellValueType string
type SortDirection int

const (
	SortDesc                    SortDirection = 1
	SortAsc                     SortDirection = 2
	SortAllStrings              CellValueType = "all-strings"
	SortAllStringsCaseSensitive CellValueType = "all-strings-case-sensitive"
	SortAllDollars              CellValueType = "all-dollars"
	SortAllPercentages          CellValueType = "all-percentages"
	SortAllFloats               CellValueType = "all-floats"
)

// SortStrategy for sheet
type SortStrategy struct {
	// Optional
	Column string

	// Required
	ColumnIndex      int
	ColumnValuesType CellValueType
	Direction        SortDirection
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
	column := s.getColumn()
	start := re.ReplaceAllString(topLeftCell, column+"$1")
	end := re.ReplaceAllString(topRightCell, column+"$1")
	return fmt.Sprintf("%s:%s", start, end)
}

func (s *SortStrategy) shouldSwapRows(rowA, rowB *Row) bool {
	switch s.ColumnValuesType {
	case SortAllStrings:
		v1 := strings.ToLower(rowA.Cells[s.ColumnIndex].String())
		v2 := strings.ToLower(rowB.Cells[s.ColumnIndex].String())

		if s.Direction == SortDesc {
			return v1 < v2
		}

		return v1 > v2
	case SortAllStringsCaseSensitive:
		v1 := rowA.Cells[s.ColumnIndex].String()
		v2 := rowB.Cells[s.ColumnIndex].String()

		if s.Direction == SortDesc {
			return v1 < v2
		}

		return v1 > v2
	case SortAllDollars:
		v1 := parseCellValueToFloat(rowA.Cells[s.ColumnIndex], "$")
		v2 := parseCellValueToFloat(rowB.Cells[s.ColumnIndex], "$")

		if s.Direction == SortDesc {
			return v1 < v2
		}

		return v1 > v2
	case SortAllPercentages:
		v1 := parseCellValueToFloat(rowA.Cells[s.ColumnIndex], "%")
		v2 := parseCellValueToFloat(rowB.Cells[s.ColumnIndex], "%")

		if s.Direction == SortDesc {
			return v1 < v2
		}

		return v1 > v2
	case SortAllFloats:
		v1 := parseCellValueToFloat(rowA.Cells[s.ColumnIndex], "")
		v2 := parseCellValueToFloat(rowB.Cells[s.ColumnIndex], "")

		if s.Direction == SortDesc {
			return v1 < v2
		}

		return v1 > v2
	default:
		return false
	}
}

func (s *SortStrategy) getColumn() string {
	if s.Column != "" {
		return s.Column
	}

	return resolveColumn(s.ColumnIndex)
}

func parseCellValueToFloat(c *Cell, cut string) float64 {
	stripped := strings.Trim(strings.Replace(c.String(), cut, "", 1), " ")
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
