package xlsx

import (
	"fmt"
	"testing"
)

func Test_ResolveColumn(t *testing.T) {
	t.Parallel()

	tt := []struct {
		index  int
		column string
	}{
		{0, "A"},
		{1, "B"},
		{7, "H"},
		{10, "K"},
		{22, "W"},
		{25, "Z"},
		{26, "AA"},
		{30, "AE"},
	}

	for _, tc := range tt {
		tc := tc

		t.Run(fmt.Sprintf("Index: %d", tc.index), func(t *testing.T) {
			s := SortStrategy{ColumnIndex: tc.index}
			c := s.getColumn()

			if c != tc.column {
				t.Fatalf("expected column to be %s, got %s", tc.column, c)
			}
		})
	}
}

func Test_ParseCellValueToFloat(t *testing.T) {
	t.Parallel()

	tt := []struct {
		input    string
		expected float64
		cut      string
	}{
		{"$ 0.87", 0.87, "$"},
		{"$  4.99", 4.99, "$"},
		{"$ 55.00 ", 55.00, "$"},
		{"$ 68", 68.00, "$"},
		{"$234", 234.00, "$"},
		{"199.02", 199.02, "$"},
		{"%111.42", 111.42, "%"},
		{"%09.40", 9.40, "%"},
		{"%00.42", 0.42, "%"},
		{"34.98", 34.98, "%"},
	}

	for _, tc := range tt {
		tc := tc

		t.Run(fmt.Sprintf("Input %s and cut %s", tc.input, tc.cut), func(t *testing.T) {
			f := parseCellValueToFloat(&Cell{Value: tc.input}, tc.cut)

			if f != tc.expected {
				t.Fatalf("expected column to be %.2f, got %.2f", tc.expected, f)
			}
		})
	}
}

func Test_SortingByStringValues(t *testing.T) {
	t.Parallel()

	tt := []struct {
		value         string
		order         SortDirection
		expectedIndex int
		hasHeader     bool
	}{
		{"Zap Hike'N Strike Stun 950,000 Volts Gun/Flashlight", SortAsc, 4, false},
		{"Zap Hike'N Strike Stun 950,000 Volts Gun/Flashlight", SortDesc, 0, false},
		{"Zap Hike'N Strike Stun 950,000 Volts Gun/Flashlight", SortAsc, 5, true},
		{"Zap Hike'N Strike Stun 950,000 Volts Gun/Flashlight", SortDesc, 1, true},
		{"911 Air Horn", SortAsc, 1, false},
		{"911 Air Horn", SortDesc, 3, false},
		{"911 Air Horn", SortAsc, 2, true},
		{"911 Air Horn", SortDesc, 4, true},
		{"Aya Concealed Carry Purse (Brown)", SortAsc, 1, false},
		{"Aya Concealed Carry Purse (Brown)", SortDesc, 3, false},
		{"Aya Concealed Carry Purse (Brown)", SortAsc, 2, true},
		{"Aya Concealed Carry Purse (Brown)", SortDesc, 4, true},
		{"Home Safe Safety Beam", SortAsc, 2, false},
		{"Home Safe Safety Beam", SortDesc, 2, false},
		{"Home Safe Safety Beam", SortAsc, 3, true},
		{"Home Safe Safety Beam", SortDesc, 3, true},
	}

	tvs := []string{
		`Women's Radiant Concealed Carry Purse: Wine`,
		`5 Inch IR Dummy Camera Silver`,
		`Can Safe Shaving Cream`,
		`zaaap product`,
	}

	for _, tc := range tt {
		tc := tc

		t.Run(fmt.Sprintf("Value %s and order %v", tc.value, tc.order), func(t *testing.T) {
			file := NewFile()
			sheet, err := file.AddSheet("Sheet1")
			if err != nil {
				t.Fatal(err)
			}

			if tc.hasHeader {
				r := sheet.AddRow()
				r.AddCell()
				cell := r.AddCell()
				cell.Value = `Amazon Title`
				r.AddCell()
			}

			r := sheet.AddRow()
			r.AddCell()
			cell := r.AddCell()
			cell.Value = tc.value
			r.AddCell()

			for _, v := range tvs {
				r := sheet.AddRow()
				r.AddCell()
				cell := r.AddCell()
				cell.Value = v
				r.AddCell()
			}

			sheet.SortByColumn(&SortStrategy{
				ColumnIndex:      1,
				Direction:        tc.order,
				ColumnValuesType: SortAllStrings,
				HasHeader:        tc.hasHeader,
			})

			if sheet.Rows[tc.expectedIndex].Cells[1].String() != tc.value {
				t.Fatalf("expected cell to contain '%s', got '%s'", tc.value, sheet.Rows[tc.expectedIndex].Cells[1].String())
			}
		})
	}
}

func Test_SortingByPercentageValues(t *testing.T) {
	t.Parallel()

	tt := []struct {
		value         string
		order         SortDirection
		expectedIndex int
		hasHeader     bool
	}{
		{"%23.5285141468048096", SortAsc, 4, false},
		{"% 23.5285141468048096", SortDesc, 0, false},
		{"%  23.5285141468048096", SortAsc, 5, true},
		{"23.5285141468048096", SortDesc, 1, true},
		{"%  7.11", SortAsc, 3, false},
		{"% 7.11", SortDesc, 1, false},
		{"%7.11", SortAsc, 4, true},
		{"%7.11", SortDesc, 2, true},
		{"%2.5", SortAsc, 1, false},
		{"%2.5", SortDesc, 3, false},
		{"%2.5", SortAsc, 2, true},
		{"%2.5", SortDesc, 4, true},
		{"% 5.690", SortAsc, 2, false},
		{"% 5.690", SortDesc, 2, false},
		{"% 5.690", SortAsc, 3, true},
		{"% 5.690", SortDesc, 3, true},
	}

	tvs := []string{
		`%5.751677989959717`,
		`%  8.293103218078613`,
		`1.2449438571929932`,
		`% 3.7047061920166016`,
	}

	for _, tc := range tt {
		tc := tc

		t.Run(fmt.Sprintf("Value %s and order %v", tc.value, tc.order), func(t *testing.T) {
			file := NewFile()
			sheet, err := file.AddSheet("Sheet1")
			if err != nil {
				t.Fatal(err)
			}

			if tc.hasHeader {
				r := sheet.AddRow()
				r.AddCell()
				cell := r.AddCell()
				cell.Value = `ROI`
				r.AddCell()
			}

			r := sheet.AddRow()
			r.AddCell()
			cell := r.AddCell()
			cell.Value = tc.value
			r.AddCell()

			for _, v := range tvs {
				r := sheet.AddRow()
				r.AddCell()
				cell := r.AddCell()
				cell.Value = v
				r.AddCell()
			}

			sheet.SortByColumn(&SortStrategy{
				ColumnIndex:      1,
				Direction:        tc.order,
				ColumnValuesType: SortAllPercentages,
				HasHeader:        tc.hasHeader,
			})

			if sheet.Rows[tc.expectedIndex].Cells[1].String() != tc.value {
				t.Fatalf("expected cell to contain '%s', got '%s'", tc.value, sheet.Rows[tc.expectedIndex].Cells[1].String())
			}
		})
	}
}
