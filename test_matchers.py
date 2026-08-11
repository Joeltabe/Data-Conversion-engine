# -*- coding: utf-8 -*-
"""Unit tests for the matchers module. Run: python test_matchers.py"""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from matchers import (  # noqa: E402
    build_candidates,
    build_name_queries,
    classify_match,
    jaro_winkler,
    matricule_fuzzy_key,
    matricule_similarity,
    name_similarity,
    normalize_matricule,
    normalize_name,
    parse_api_payload,
    score_candidate,
)


class TestNormalizeName(unittest.TestCase):
    def test_strips_titles_accents_punctuation(self):
        s, t = normalize_name("Prof. NANA, FOUDA-Jean (Dr.)")
        self.assertEqual(s, "NANA FOUDA JEAN")
        self.assertEqual(t, ["NANA", "FOUDA", "JEAN"])

    def test_case_insensitive(self):
        self.assertEqual(normalize_name("aseh iya makane")[0], "ASEH IYA MAKANE")

    def test_empty(self):
        self.assertEqual(normalize_name(""), ("", []))
        self.assertEqual(normalize_name(None), ("", []))


class TestNormalizeMatricule(unittest.TestCase):
    def test_strips_separators(self):
        self.assertEqual(normalize_matricule("LMU-24NUR001"), "LMU24NUR001")
        self.assertEqual(normalize_matricule("LMU 24NUR018"), "LMU24NUR018")
        self.assertEqual(normalize_matricule("  lmu24nur005 "), "LMU24NUR005")

    def test_fuzzy_key_o0(self):
        self.assertEqual(
            matricule_fuzzy_key("LMU24NURO16"),
            matricule_fuzzy_key("LMU24NUR016"),
        )

    def test_fuzzy_key_il1(self):
        self.assertEqual(
            matricule_fuzzy_key("LMU24NUR001"),
            matricule_fuzzy_key("1MU24NUR001"),
        )

    def test_similarity_levels(self):
        self.assertEqual(matricule_similarity("LMU-24NUR001", "LMU24NUR001"), 1.0)
        self.assertEqual(matricule_similarity("LMU24NURO16", "LMU24NUR016"), 0.9)
        self.assertEqual(matricule_similarity("LMU24NUR001", "LMU24NUR999"), 0.0)
        self.assertEqual(matricule_similarity("", "LMU24NUR001"), 0.0)

    def test_lost_prefix_variant(self):
        self.assertEqual(matricule_similarity("24NURO16", "LMU24NUR016"), 0.9)
        self.assertEqual(matricule_similarity("24NUR016", "LMU-24NUR016"), 0.9)


class TestJaroWinkler(unittest.TestCase):
    def test_identical(self):
        self.assertEqual(jaro_winkler("MAKANE", "MAKANE"), 1.0)

    def test_close(self):
        self.assertGreater(jaro_winkler("MAKANE", "MAKANI"), 0.9)

    def test_far(self):
        self.assertLess(jaro_winkler("MAKANE", "ASONGAFAC"), 0.6)


class TestNameSimilarity(unittest.TestCase):
    def test_identical(self):
        self.assertEqual(name_similarity("ASEH IYA MAKANE", "ASEH IYA MAKANE"), 1.0)

    def test_reordered(self):
        self.assertEqual(
            name_similarity("ASEH IYA MAKANE", "MAKANE ASEH IYA"), 1.0
        )

    def test_missing_middle(self):
        self.assertGreaterEqual(
            name_similarity("ASONGAFAC JERRY NGIMKENG", "ASONGAFAC NGIMKENG"), 0.85
        )

    def test_typo(self):
        self.assertGreaterEqual(
            name_similarity("ASEH IYA MAKANE", "ASEH IYA MAKANI"), 0.9
        )

    def test_accented(self):
        self.assertEqual(
            name_similarity("NANA FOUDA", "NANA FOUDA"), 1.0
        )
        self.assertGreaterEqual(
            name_similarity("RENÉ TCHOUPOU", "RENE TCHOUPOU"), 0.95
        )

    def test_different_students_low(self):
        self.assertLess(name_similarity("NGWA LEM JANET", "FON INNOCENT TABE"), 0.5)

    def test_same_surname_not_verified(self):
        # shares surname only -> should NOT reach the strong threshold
        self.assertLess(name_similarity("NGWA LEM JANET", "NGWA JOEL ATANGA"), NAME_STRONG := 0.92)
        self.assertLess(name_similarity("NGWA LEM JANET", "NGWA JOEL ATANGA"), 0.9)

    def test_empty(self):
        self.assertEqual(name_similarity("", "ASEH"), 0.0)


class TestBuildQueries(unittest.TestCase):
    def test_full_first(self):
        q = build_name_queries("ASEH IYA MAKANE")
        self.assertEqual(q[0], "ASEH IYA MAKANE")

    def test_reverse_and_pair_order(self):
        q = build_name_queries("ASEH IYA MAKANE", limit=10)
        joined = " ".join(q)
        self.assertIn("MAKANE ASEH IYA", joined)
        self.assertIn("MAKANE ASEH", joined)
        self.assertIn("ASEH MAKANE", joined)

    def test_dedup_and_cap(self):
        q = build_name_queries("JEAN NANA", limit=4)
        self.assertEqual(len(set(q)), len(q))
        self.assertLessEqual(len(q), 4)

    def test_single_word(self):
        self.assertEqual(build_name_queries("BLESS"), ["BLESS"])

    def test_empty(self):
        self.assertEqual(build_name_queries(""), [])


class TestScoring(unittest.TestCase):
    def test_matricule_anchor_boost(self):
        conf_exact = score_candidate("NGWA LEM JANET", "NGWA LEM JANET",
                                     "LMU24NUR014", "LMU-24NUR014")
        conf_nomatch = score_candidate("NGWA LEM JANET", "NGWA LEM JANET",
                                       "LMU24NUR014", "LMU24NUR888")
        self.assertEqual(conf_exact, 1.0)
        self.assertEqual(conf_nomatch, 1.0)  # name alone is perfect anyway
        self.assertGreaterEqual(conf_exact, conf_nomatch)

    def test_boost_when_matricule_matches(self):
        # mediocre name, exact matricule -> boosted above raw name score
        conf = score_candidate("NGWA LEM JANET", "NGWA JANET",
                               "LMU24NUR014", "LMU24NUR014")
        self.assertGreater(conf, name_similarity("NGWA LEM JANET", "NGWA JANET"))

    def test_no_boost_when_matricule_differs(self):
        conf = score_candidate("NGWA LEM JANET", "NGWA JANET",
                               "LMU24NUR014", "LMU24NUR999")
        self.assertEqual(conf, name_similarity("NGWA LEM JANET", "NGWA JANET"))


class TestClassification(unittest.TestCase):
    def _cand(self, pdf_name, pdf_mat, api_name, api_mat):
        raw = {api_mat: {"name": api_name, "years": {"2023/2024"}, "query_used": "Q"}}
        return build_candidates(pdf_name, pdf_mat, raw)

    def test_verified_exact(self):
        c = self._cand("ASEH IYA MAKANE", "LMU24NUR001", "ASEH IYA MAKANE", "LMU-24NUR001")
        self.assertEqual(classify_match("ASEH IYA MAKANE", "LMU24NUR001", c), "verified")

    def test_verified_same_matricule_typo_name(self):
        c = self._cand("NGWA LEM JANET", "LMU24NUR014", "NGWA LEM JOAN", "LMU24NUR014")
        self.assertEqual(classify_match("NGWA LEM JANET", "LMU24NUR014", c), "verified")

    def test_mismatch(self):
        c = self._cand("ASEH IYA MAKANE", "LMU24NUR001", "ASEH IYA MAKANE", "LMU-24NUR111")
        self.assertEqual(classify_match("ASEH IYA MAKANE", "LMU24NUR001", c), "mismatch")

    def test_same_surname_not_review(self):
        # shares surname only -> below review threshold (avoids noise)
        c = self._cand("NGWA LEM JANET", "LMU24NUR014", "NGWA JOEL ATANGA", "LMU24NUR888")
        self.assertEqual(classify_match("NGWA LEM JANET", "LMU24NUR014", c), "not_found")

    def test_review_ambiguous_middle(self):
        # strong but imperfect overlap with different matricule -> review
        c = self._cand("NGWA LEM JANET", "LMU24NUR014", "NGWA JANET", "LMU24NUR888")
        self.assertEqual(classify_match("NGWA LEM JANET", "LMU24NUR014", c), "review")

    def test_not_found(self):
        self.assertEqual(classify_match("X Y", "M", []), "not_found")


class TestParseApiPayload(unittest.TestCase):
    def test_hit(self):
        self.assertEqual(
            parse_api_payload({"matricule": " LMU-24NUR001 ", "fname": " MAKANE "}),
            {"matricule": "LMU-24NUR001", "name": "MAKANE"},
        )

    def test_miss(self):
        self.assertIsNone(parse_api_payload({"ans3": "Did not find a Name like that !"}))
        self.assertIsNone(parse_api_payload(None))
        self.assertIsNone(parse_api_payload([]))


if __name__ == "__main__":
    unittest.main(verbosity=2)
