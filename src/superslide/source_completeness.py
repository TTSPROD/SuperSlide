"""Prevent unfilled source-template placeholder text from leaking into LLM drafts."""

from copy import deepcopy
import re


_DUMMY_TEXT = re.compile(
    r"(?:\blorem\s+ipsum\b|"
    r"\b(?:название|заголовок)(?:\s+(?:блока|карточки|графика|диаграммы|раздела))?\s*[#№]\s*\d+\b|"
    r"\bтекстовый\s+заголовок\s*[#№]\s*\d+\b|"
    r"\bи\s+небольшой\s+подзаголовок\b)",
    re.IGNORECASE,
)


def is_template_placeholder(sample):
    """Recognize demonstrative source text, not ordinary source content.

    Compact catalog samples may be truncated and combine multiple paragraphs.
    Generic words such as «Итог» are not sufficient to identify a placeholder.
    """
    return isinstance(sample, str) and bool(_DUMMY_TEXT.search(sample))


def clear_unfilled_template_text(draft, catalog):
    """Clear omitted samples only when catalog explicitly marks that safe.

    Run only on LLM-generated drafts, not manual specifications. User-provided
    updates, tables, charts, photos and icons are never changed. The source
    catalog controls the scope: only exposed slide text fields are considered.
    An unknown or invalid slide is left for normal schema validation to reject.
    """
    result = deepcopy(draft)
    if not isinstance(result, dict) or not isinstance(result.get("slides"), list):
        return result
    for slide in result["slides"]:
        if not isinstance(slide, dict) or not isinstance(slide.get("updates", []), list):
            continue
        variants = catalog.get(slide.get("layout_id"), {}).get("variants", {})
        fields = variants.get(slide.get("variant_id"), {}).get("fields", {})
        updates = slide.setdefault("updates", [])
        used = {item.get("field_id") for item in updates if isinstance(item, dict)}
        for field_id, field in fields.items():
            if (field_id in used or not field.get("safe_to_clear_if_unused", False)
                    or not field_id.startswith("slide_text_")
                    or field.get("kind") != "text" or not is_template_placeholder(field.get("sample_text"))):
                continue
            roles = field.get("text_budget", {}).get("paragraph_roles", [])
            count = field.get("text_budget", {}).get("max_paragraphs", field.get("paragraph_count", 1))
            if "heading" in roles and "body" in roles and isinstance(count, int) and 1 <= count <= 100:
                updates.append({"field_id": field_id, "paragraphs": [""] * count})
            else:
                updates.append({"field_id": field_id, "text": ""})
    return result


def incomplete_slot_groups(draft, catalog, slide_indexes=None):
    """Report missing content in reliably identified repeated field groups.

    Group members correspond to real visual slots. Leaving one empty or using
    template placeholder text is not repaired by hiding arbitrary shapes:
    callers can ask the model to choose a matching variant instead.
    """
    issues = []
    if not isinstance(draft, dict) or not isinstance(draft.get("slides"), list):
        return issues
    indexes = slide_indexes or range(1, len(draft["slides"]) + 1)
    for index, slide in zip(indexes, draft["slides"]):
        if not isinstance(slide, dict) or not isinstance(slide.get("updates", []), list):
            continue
        variant = catalog.get(slide.get("layout_id"), {}).get("variants", {}).get(slide.get("variant_id"), {})
        updates = {u.get("field_id"): u for u in slide.get("updates", []) if isinstance(u, dict)}
        for group in variant.get("item_slots", {}).get("groups", []):
            if group.get("kind") != "numbered_text_fields" or len(group.get("field_ids", [])) < 2:
                continue
            missing = []
            for field_id in group["field_ids"]:
                update = updates.get(field_id, {})
                text = update.get("text")
                if text is None and isinstance(update.get("paragraphs"), list):
                    text = " ".join(p for p in update["paragraphs"] if isinstance(p, str))
                if not isinstance(text, str) or not text.strip() or is_template_placeholder(text):
                    missing.append(field_id)
            if missing:
                issues.append({"slide_index": index, "reason": "unfilled_group", "field_ids": missing,
                               "limit": len(group["field_ids"]),
                               "actual": len(group["field_ids"]) - len(missing)})
    return issues


def incomplete_template_fields(draft, catalog, slide_indexes=None):
    """Find omitted/unfilled sample text even when group geometry is unknown.

    Excludes known group members, which are reported together by
    incomplete_slot_groups. Applying this check to LLM output only allows
    explicit source-demo and manually authored JSON to remain unchanged.
    """
    issues = []
    if not isinstance(draft, dict) or not isinstance(draft.get("slides"), list):
        return issues
    indexes = slide_indexes or range(1, len(draft["slides"]) + 1)
    for index, slide in zip(indexes, draft["slides"]):
        if not isinstance(slide, dict) or not isinstance(slide.get("updates", []), list):
            continue
        variant = catalog.get(slide.get("layout_id"), {}).get("variants", {}).get(slide.get("variant_id"), {})
        grouped = {field_id for group in variant.get("item_slots", {}).get("groups", [])
                   if group.get("kind") == "numbered_text_fields" and len(group.get("field_ids", [])) >= 2
                   for field_id in group.get("field_ids", [])}
        updates = {u.get("field_id"): u for u in slide.get("updates", []) if isinstance(u, dict)}
        for field_id, field in variant.get("fields", {}).items():
            if (field_id in grouped or not field_id.startswith("slide_text_") or field.get("kind") != "text"
                    or not is_template_placeholder(field.get("sample_text"))):
                continue
            update = updates.get(field_id, {})
            text = update.get("text")
            if text is None and isinstance(update.get("paragraphs"), list):
                text = " ".join(p for p in update["paragraphs"] if isinstance(p, str))
            if not isinstance(text, str) or not text.strip() or is_template_placeholder(text):
                issues.append({"slide_index": index, "field_id": field_id, "reason": "template_placeholder"})
    return issues
