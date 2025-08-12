package main

import "testing"

func TestSanitizeFileComponent(t *testing.T) {
	cases := map[string]string{
		"":        "sheet",
		" ":       "sheet",
		"a/b":     "a_b",
		"x*y?z":   "x_y_z",
		"<name>":  "_name_",
		"valid":   "valid",
	}
	for in, want := range cases {
		got := sanitizeFileComponent(in)
		if got != want {
			t.Fatalf("sanitizeFileComponent(%q) = %q, want %q", in, got, want)
		}
	}
}
