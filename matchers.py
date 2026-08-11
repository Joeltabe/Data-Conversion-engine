# -*- coding: utf-8 -*-
"""
Production-grade name & matricule matching utilities.

Pure, dependency-free functions used by the transcript engine's portal
cross-check. Covers:

  * Unicode-aware name normalisation (accents, titles, punctuation, casing)
  * OCR-tolerant matricule normalisation (separators, O/0, I/1/L ambiguity)
  * Token-set + Jaro-Winkler + sequence name similarity
  * Prioritised, bounded search-query generation for the portal API
  * Candidate scoring and result classification (verified / mismatch /
    review / not_found)
"""

from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher

# Match thresholds
NAME_STRONG = 0.92  # high-confidence name agreement
NAME_REVIEW = 0.70  # below this, a candidate is not a plausible match

TITLES = {
    "MR", "MRS", "MS", "MISS", "MADAM", "MADAME", "MLLE", "MME",
    "DR", "PROF", "PROFESSOR", "ENG", "ENGR", "HON", "REV", "PASTOR",
    "SIR", "MASTER", "CHEF",
}

_ALNUM = re.compile(r"[^A-Z0-9]")
_KEEP_LETTERS = re.compile(r"[^A-Z ]+")


# ────────────────────────────────────────────────────────────────────────────
#  NORMALISATION
# ────────────────────────────────────────────────────────────────────────────
def strip_accents(text):
    """Remove combining diacritics (é -> e, ü -> u, Ñ -> N)."""
    return "".join(
        ch for ch in unicodedata.normalize("NFKD", str(text))
        if not unicodedata.combining(ch)
    )


def normalize_name(name):
    """Return ``(canonical string, token list)``.

    Strips accents, uppercases, drops titles and punctuation so that
    "Prof. NANA, FOUDA-Jean" and "NANA FOUDA JEAN" compare equal.
    """
    if not name:
        return "", []
    s = _KEEP_LETTERS.sub(" ", strip_accents(name).upper())
    tokens = [t for t in s.split() if t and t not in TITLES]
    return " ".join(tokens), tokens


def normalize_matricule(mat):
    """Canonical matricule: uppercased, separators removed.

    ``LMU-24NUR001`` -> ``LMU24NUR001``, ``LMU 24NUR018`` -> ``LMU24NUR018``.
    """
    if not mat:
        return ""
    return _ALNUM.sub("", strip_accents(mat).upper())


def matricule_fuzzy_key(mat):
    """Matricule comparison key tolerant to OCR O/0 and I/1/L ambiguity.

    ``24NURO16`` and ``24NUR016`` produce the same key.
    """
    s = normalize_matricule(mat)
    return s.replace("O", "0").replace("I", "1").replace("L", "1")


def matricule_similarity(pdf_mat, api_mat):
    """Matricule agreement: 1.0 exact, 0.9 OCR-fuzzy, 0.0 otherwise.

    ``0.9`` covers O/0 and I/1/L confusion as well as a lost school prefix
    (``24NUR016`` vs ``LMU24NUR016``) when the shared suffix is long enough
    to be meaningful. Verification still requires strong name agreement, so
    the weaker 0.9 signal can never verify on its own.
    """
    a = normalize_matricule(pdf_mat)
    b = normalize_matricule(api_mat)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    fa = matricule_fuzzy_key(a)
    fb = matricule_fuzzy_key(b)
    if fa == fb:
        return 0.9
    if len(a) >= 6 and len(b) >= 6 and (fa.endswith(fb) or fb.endswith(fa)):
        return 0.9
    return 0.0


# ────────────────────────────────────────────────────────────────────────────
#  STRING / NAME SIMILARITY
# ────────────────────────────────────────────────────────────────────────────
def _jaro(a, b):
    if a == b:
        return 1.0
    la, lb = len(a), len(b)
    if la == 0 or lb == 0:
        return 0.0
    match_dist = max(0, max(la, lb) // 2 - 1)
    a_matches = [False] * la
    b_matches = [False] * lb
    matches = 0
    for i in range(la):
        start = max(0, i - match_dist)
        end = min(i + match_dist + 1, lb)
        for j in range(start, end):
            if b_matches[j] or a[i] != b[j]:
                continue
            a_matches[i] = True
            b_matches[j] = True
            matches += 1
            break
    if matches == 0:
        return 0.0
    transpositions = 0
    k = 0
    for i in range(la):
        if not a_matches[i]:
            continue
        while not b_matches[k]:
            k += 1
        if a[i] != b[k]:
            transpositions += 1
        k += 1
    m = float(matches)
    return (m / la + m / lb + (m - transpositions / 2.0) / m) / 3.0


def jaro_winkler(a, b, prefix_weight=0.1):
    """Jaro-Winkler similarity (0..1), boosts common short prefixes."""
    a, b = str(a).upper(), str(b).upper()
    j = _jaro(a, b)
    if j in (0.0, 1.0):
        return j
    prefix = 0
    for ca, cb in zip(a, b):
        if ca != cb:
            break
        prefix += 1
        if prefix == 4:
            break
    return j + prefix * prefix_weight * (1.0 - j)


def _token_set_ratio(tokens_a, tokens_b):
    """Token-set ratio: order- and extra-token-insensitive, 0..1.

    "ASEH IYA MAKANE" vs "MAKANE ASEH IYA" -> 1.0.
    "NGWA LEM JANET"  vs "NGWA LEM JOAN"  -> ~0.79.
    """
    set_a, set_b = set(tokens_a), set(tokens_b)
    intersection = sorted(set_a & set_b)
    diff_a = sorted(set_a - set_b)
    diff_b = sorted(set_b - set_a)
    c_a = " ".join(intersection + diff_a)
    c_b = " ".join(intersection + diff_b)
    s1 = SequenceMatcher(None, c_a, c_b).ratio()
    s2 = SequenceMatcher(None, c_a, " ".join(tokens_b)).ratio()
    s3 = SequenceMatcher(None, c_b, " ".join(tokens_a)).ratio()
    return max(s1, s2, s3)


def _token_jw_score(tokens_a, tokens_b):
    """Average best Jaro-Winkler match per token, handles typos."""
    if not tokens_a or not tokens_b:
        return 0.0
    if len(tokens_a) <= len(tokens_b):
        smaller, larger = tokens_a, tokens_b
    else:
        smaller, larger = tokens_b, tokens_a
    total = 0.0
    for t in smaller:
        total += max(jaro_winkler(t, u) for u in larger)
    return total / len(smaller)


def name_similarity(a, b):
    """Combined name similarity (0..1).

    Robust to token reordering, dropped/extra middle names, initials and
    spelling typos — but conservative enough to separate same-surname
    strangers. Uses the best of: raw sequence ratio, token-set ratio and
    per-token Jaro-Winkler.
    """
    na, ta = normalize_name(a)
    nb, tb = normalize_name(b)
    if not ta or not tb:
        return 0.0
    if na == nb:
        return 1.0
    seq = SequenceMatcher(None, na, nb).ratio()
    ts = _token_set_ratio(ta, tb)
    jw = _token_jw_score(ta, tb)
    # Token-set ratio is the anchor; Jaro-Winkler only boosts when the token
    # sets already agree (catches typos) and is never allowed to carry a
    # match on its own (avoids same-surname false positives).
    blended = 0.5 * ts + 0.5 * jw
    return round(max(seq, ts, blended), 4)


# ────────────────────────────────────────────────────────────────────────────
#  QUERY GENERATION
# ────────────────────────────────────────────────────────────────────────────
def build_name_queries(name, limit=10):
    """Prioritised, deduplicated search queries for a name.

    Order matters: full name first, then surname-first orderings (common
    for the portal), then initial/abbreviated forms, then single tokens.
    Each query is upper-cased and whitespace-collapsed. Returned list is
    capped at ``limit`` entries.
    """
    _, tokens = normalize_name(name)
    if not tokens:
        return []
    seen, out = set(), []
    n = len(tokens)

    def add(s):
        s = " ".join(s.split()).upper()
        if s and s not in seen:
            seen.add(s)
            out.append(s)

    add(" ".join(tokens))
    if n >= 2:
        add(" ".join(reversed(tokens)))
        add(f"{tokens[-1]} {tokens[0]}")
        add(f"{tokens[0]} {tokens[-1]}")
        for i in range(1, n - 1):
            add(" ".join(tokens[i:] + tokens[:i]))
        if n >= 3:
            add(f"{tokens[0]} {tokens[-1]}")
        add(f"{tokens[-1]} {tokens[0][0]}")
        add(f"{tokens[0][0]} {tokens[-1]}")
    for t in tokens:
        add(t)
    return out[:limit]


# ────────────────────────────────────────────────────────────────────────────
#  SCORING & CLASSIFICATION
# ────────────────────────────────────────────────────────────────────────────
def score_candidate(pdf_name, api_name, pdf_mat, api_mat):
    """Combine name + matricule evidence into one confidence score (0..1).

    An exact matricule match strongly anchors identity, so a moderate name
    match is boosted; when matricules differ, only the name similarity
    remains.
    """
    ns = name_similarity(pdf_name, api_name)
    ms = matricule_similarity(pdf_mat, api_mat)
    if ms >= 1.0:
        return round(max(ns, 0.55 + 0.45 * ns), 4)
    if ms >= 0.9:
        return round(max(ns, 0.50 + 0.45 * ns), 4)
    return round(ns, 4)


def classify_match(pdf_name, pdf_mat, candidates):
    """Classify the best candidate: verified | mismatch | review | not_found.

    ``candidates`` must be pre-scored dicts (see ``build_candidates``).
    """
    if not candidates:
        return "not_found"
    best = candidates[0]
    ns = best.get("name_similarity", 0.0)
    ms = best.get("matricule_similarity", 0.0)
    if ms >= 1.0 and ns >= 0.50:
        return "verified"
    if ms >= 0.9 and ns >= NAME_STRONG:
        return "verified"
    if ns >= NAME_STRONG and ms <= 0.0:
        return "mismatch"
    if best.get("confidence", 0.0) >= NAME_REVIEW:
        return "review"
    return "not_found"


def build_candidates(pdf_name, pdf_mat, raw):
    """Turn raw per-matricule accumulations into scored, sorted candidates.

    ``raw`` maps api-matricule -> {"name", "years": set, "query_used"}.
    """
    out = []
    for mat, rec in raw.items():
        api_name = rec.get("name", "")
        ns = name_similarity(pdf_name, api_name)
        ms = matricule_similarity(pdf_mat, mat)
        out.append({
            "matricule": mat,
            "name": api_name,
            "name_similarity": round(ns, 4),
            "matricule_similarity": ms,
            "confidence": score_candidate(pdf_name, api_name, pdf_mat, mat),
            "years": sorted(rec.get("years", ())),
            "query_used": rec.get("query_used", ""),
        })
    out.sort(key=lambda c: (-c["confidence"], -c["name_similarity"]))
    return out


def parse_api_payload(data):
    """Validate and normalise a portal response.

    A hit looks like ``{"matricule": "LMU-24NUR001", "fname": "..."}`` and a
    miss like ``{"ans3": "Did not find a Name like that !"}``. Returns a
    ``{"matricule", "name"}`` dict or ``None``.
    """
    if not isinstance(data, dict):
        return None
    mat = data.get("matricule")
    if not mat:
        return None
    name = data.get("fname") or data.get("name") or data.get("full_name") or ""
    return {"matricule": str(mat).strip(), "name": str(name).strip()}
