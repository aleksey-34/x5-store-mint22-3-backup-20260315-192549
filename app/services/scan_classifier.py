from __future__ import annotations

import re
from dataclasses import dataclass


@dataclass(frozen=True)
class ScanClassification:
    predicted_doc_type: str
    confidence: float
    source: str
    matched_keywords: list[str]


# Lightweight keyword-based classifier as a practical MVP before heavier ML.
KEYWORDS: dict[str, tuple[str, ...]] = {
    "order": (
        "приказ",
        "order",
        "назнач",
        "допуск",
        "наряд",
        "permit",
        "appointment",
        "распоряж",
        "prikaz",
    ),
    "employee_passport": (
        "паспорт",
        "passport",
        "удостовер",
        "certificate",
        "свидетельств",
        "снилс",
        "инн",
        "id",
    ),
    "hidden_work_act": (
        "awr",
        "акт скрыт",
        "скрыт",
        "hidden work",
        "освидетельств",
        "приемк",
        "бетон",
        "армир",
    ),
}


def _normalize(value: str) -> str:
    lowered = value.lower()
    lowered = lowered.replace("_", " ")
    lowered = re.sub(r"\s+", " ", lowered)
    return lowered.strip()


def _score_text(text: str) -> tuple[dict[str, int], dict[str, list[str]]]:
    normalized = _normalize(text)
    scores = {doc_type: 0 for doc_type in KEYWORDS}
    hits = {doc_type: [] for doc_type in KEYWORDS}

    for doc_type, keywords in KEYWORDS.items():
        for keyword in keywords:
            if keyword in normalized:
                scores[doc_type] += 1
                hits[doc_type].append(keyword)

    return scores, hits


def _confidence(top_score: int, second_score: int) -> float:
    if top_score <= 0:
        return 0.0

    # Keep confidence conservative for noisy scanned text.
    if second_score <= 0:
        return round(min(0.96, 0.55 + 0.08 * top_score), 2)

    ratio = top_score / float(top_score + second_score)
    return round(max(0.5, min(0.94, ratio)), 2)


def _best_label(scores: dict[str, int], hits: dict[str, list[str]], source: str) -> ScanClassification:
    ordered = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    top_label, top_score = ordered[0]
    second_score = ordered[1][1] if len(ordered) > 1 else 0

    if top_score <= 0:
        return ScanClassification(
            predicted_doc_type="unknown",
            confidence=0.0,
            source=source,
            matched_keywords=[],
        )

    return ScanClassification(
        predicted_doc_type=top_label,
        confidence=_confidence(top_score, second_score),
        source=source,
        matched_keywords=hits[top_label][:5],
    )


def classify_scan_filename(filename: str) -> ScanClassification:
    scores, hits = _score_text(filename)
    return _best_label(scores=scores, hits=hits, source="filename")


def classify_scan_candidate(filename: str, ocr_text: str | None = None) -> ScanClassification:
    filename_scores, filename_hits = _score_text(filename)

    if not ocr_text:
        return _best_label(scores=filename_scores, hits=filename_hits, source="filename")

    text_scores, text_hits = _score_text(ocr_text)

    combined_scores: dict[str, int] = {}
    combined_hits: dict[str, list[str]] = {}
    for doc_type in KEYWORDS:
        combined_scores[doc_type] = filename_scores[doc_type] * 2 + text_scores[doc_type] * 3
        combined_hits[doc_type] = [*filename_hits[doc_type], *text_hits[doc_type]]

    return _best_label(scores=combined_scores, hits=combined_hits, source="filename+ocr")
