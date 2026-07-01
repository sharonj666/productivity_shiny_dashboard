from pathlib import Path
import sys
import unittest
from unittest.mock import patch


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import clean_productivity as cleaning  # noqa: E402


def row(source_row: int, year: int, species: str = "") -> dict[str, str]:
    return {
        "DATE": f"{year}-05-01",
        "Species": species,
        "PLOT": "Colony",
        "Nest#": "R1",
        "A or B chick": "A",
        "Eggs": "1",
        "Chicks": "1",
        "Status": "H",
        "PFR": f"B{source_row}",
        "LOCATION": "A1",
        "_source_row": str(source_row),
    }


class MultiSpeciesAnalysisTests(unittest.TestCase):
    def test_species_column_separates_keys_and_summaries(self):
        rows = [row(2, 2024, "ROST"), row(3, 2024, "LETE")]
        with patch.object(cleaning, "read_sheet", return_value=("Data", rows)):
            result = cleaning.analyze_workbooks(Path("unused.xlsx"))

        self.assertEqual({item["species"] for item in result["nests"]}, {"ROST", "LETE"})
        self.assertEqual(len({item["nest_key"] for item in result["nests"]}), 2)
        overall = [item for item in result["summary"] if item["group"] == "Overall"]
        self.assertEqual({item["species"] for item in overall}, {"ROST", "LETE"})

    def test_workbook_species_fallback_is_required_by_caller_and_applied(self):
        rows = [row(2, 2025)]
        with patch.object(cleaning, "read_sheet", return_value=("Data", rows)):
            result = cleaning.analyze_workbooks(
                Path("unused.xlsx"), species_default="COTE"
            )

        self.assertEqual(result["nests"][0]["species"], "COTE")
        self.assertIn("|COTE|", result["nests"][0]["nest_key"])

    def test_resights_match_within_year_and_species(self):
        productivity = [row(2, 2024, "ROST"), row(3, 2024, "LETE")]
        for item in productivity:
            item["PFR"] = "SHARED"
            item["Status"] = "H"
        resight = {
            "Favorite Date": "2024-07-01",
            "Species": "LETE",
            "Combo": "SHARED",
            "Fledged?": "YES!",
            "Age": "Chick",
            "_source_row": "2",
        }
        with patch.object(
            cleaning,
            "read_sheet",
            side_effect=[("Productivity", productivity), ("Resights", [resight])],
        ):
            result = cleaning.analyze_workbooks(
                Path("productivity.xlsx"), resight_path=Path("resights.xlsx")
            )

        fledged = {
            item["species"]: item["verified_fledged"] for item in result["chicks"]
        }
        self.assertEqual(fledged, {"ROST": False, "LETE": True})


if __name__ == "__main__":
    unittest.main()
