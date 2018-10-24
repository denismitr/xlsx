package xlsx

import (
	"fmt"
	"testing"
)

func Test_ResolveColumn(t *testing.T) {
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
			c := resolveColumn(tc.index)

			if c != tc.column {
				t.Fatalf("expected column to be %s, got %s", tc.column, c)
			}
		})
	}
}

func Test_CellWithDollarsToFloat(t *testing.T) {
	tt := []struct {
		input    string
		expected float64
	}{
		{"$ 0.87", 0.87},
		{"$  4.99", 4.99},
		{"$ 55.00 ", 55.00},
		{"$ 68", 68.00},
		{"$234", 234.00},
	}

	for _, tc := range tt {
		tc := tc

		t.Run(fmt.Sprintf("Input: %s", tc.input), func(t *testing.T) {
			f := cellWithDollarsToFloat(&Cell{Value: tc.input})

			if f != tc.expected {
				t.Fatalf("expected column to be %.2f, got %.2f", tc.expected, f)
			}
		})
	}
}
