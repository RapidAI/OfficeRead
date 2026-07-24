package officeread

import "testing"

func TestCleanTextRepairsLegacyWordApostropheMojibake(t *testing.T) {
	if got := cleanText("Commission\ubb69 policy"); got != "Commission's policy" {
		t.Fatalf("cleanText() = %q", got)
	}
}

func TestCyrillicProseWithDatesIsNotEncodingTableNoise(t *testing.T) {
	prose := "Сколько лет этой Системе? Она была учреждена 25 августа 1916 г. и расширена в 2003 г."
	if looksLikeCyrillicEncodingTableNoise(prose) {
		t.Fatalf("Cyrillic prose was misclassified as encoding-table noise: %q", prose)
	}
}
