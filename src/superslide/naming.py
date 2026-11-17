"""Readable, portable download filenames derived from presentation titles."""

import unicodedata


_FORBIDDEN = set('/\\:*?"<>|')
_RESERVED = {"CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"} | {
    prefix + number for prefix in ("COM", "LPT") for number in "123456789¹²³"
}
_MAX_STEM_CHARS = 96
_MAX_STEM_BYTES = 180


def presentation_filename(title: str) -> str:
    """Keep a title recognizable while making it safe as one PPTX filename.

    Replacing path separators with spaces retains the words of the title.
    Invisible formatting controls are removed, and the stem is bounded by
    both characters and UTF-8 bytes for Unicode-friendly filesystem limits.
    """
    characters = []
    for character in unicodedata.normalize("NFC", title):
        category = unicodedata.category(character)
        if character in _FORBIDDEN or category == "Cc":
            characters.append(" ")
        elif category not in {"Cf", "Cs"}:
            characters.append(character)
    stem = " ".join("".join(characters).split()).rstrip(" .")
    while stem.casefold().endswith(".pptx"):
        stem = stem[:-5].rstrip(" .")
    stem = stem.strip(" .")
    if not stem:
        stem = "Презентация"
    if stem.partition(".")[0].rstrip().upper() in _RESERVED:
        stem = "_" + stem
    stem = stem[:_MAX_STEM_CHARS]
    while len(stem.encode("utf-8")) > _MAX_STEM_BYTES:
        stem = stem[:-1]
    return stem.rstrip(" .") + ".pptx"
