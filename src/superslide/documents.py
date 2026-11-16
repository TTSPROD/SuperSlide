"""Local, bounded extraction of presentation reference documents."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Sequence
from zipfile import BadZipFile, ZipFile


class DocumentError(ValueError):
    """A reference document cannot be processed without losing information."""


@dataclass(frozen=True)
class SourceChunk:
    id: str
    document: str
    location: str
    text: str

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class DocumentBundle:
    chunks: list[SourceChunk] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "chunks": [chunk.to_dict() for chunk in self.chunks],
            "warnings": list(self.warnings),
        }

    def context_text(self) -> str:
        return "\n\n".join(
            f"[{chunk.id}] {chunk.document} ({chunk.location})\n{chunk.text}"
            for chunk in self.chunks
        )


def _extract_docx(path: Path, document_id: str) -> list[SourceChunk]:
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError as exc:
        raise DocumentError("Для DOCX установите зависимость python-docx.") from exc

    try:
        document = Document(path)
        chunks = []
        paragraph_number = table_number = 0
        # Walk XML order so paragraphs and tables keep their reading order.
        for element in document.element.body.iterchildren():
            if element.tag == qn("w:p"):
                paragraph_number += 1
                text = Paragraph(element, document).text.strip()
                if text:
                    chunks.append(SourceChunk(
                        f"{document_id}:paragraph{paragraph_number}", path.name,
                        f"paragraph {paragraph_number}", text,
                    ))
            elif element.tag == qn("w:tbl"):
                table_number += 1
                for row_number, row in enumerate(Table(element, document).rows, 1):
                    text = " | ".join(cell.text.strip() for cell in row.cells)
                    if text.replace("|", "").strip():
                        chunks.append(SourceChunk(
                            f"{document_id}:table{table_number}:row{row_number}",
                            path.name, f"table {table_number}, row {row_number}", text,
                        ))
        return chunks
    except DocumentError:
        raise
    except Exception as exc:
        raise DocumentError(f"Не удалось прочитать DOCX «{path.name}». Проверьте файл.") from exc


def _extract_pdf(path: Path, document_id: str) -> tuple[list[SourceChunk], list[str]]:
    try:
        from pypdf import PdfReader
    except ImportError as exc:
        raise DocumentError("Для PDF установите зависимость pypdf.") from exc

    try:
        reader = PdfReader(path)
        if reader.is_encrypted:
            raise DocumentError(f"PDF «{path.name}» защищён паролем. Нужен незашифрованный файл.")
        chunks = []
        warnings = []
        for number, page in enumerate(reader.pages, 1):
            text = (page.extract_text() or "").strip()
            if text:
                chunks.append(SourceChunk(f"{document_id}:p{number}", path.name, f"page {number}", text))
            else:
                warnings.append(
                    f"«{path.name}», страница {number}: нет извлекаемого текста; "
                    "возможен скан или пустая страница. OCR не выполнялся."
                )
        if not chunks:
            raise DocumentError(
                f"В PDF «{path.name}» нет извлекаемого текстового слоя. "
                "Для сканов сначала выполните локальное OCR."
            )
        return chunks, warnings
    except DocumentError:
        raise
    except Exception as exc:
        raise DocumentError(f"Не удалось прочитать PDF «{path.name}». Проверьте файл.") from exc


def load_documents(
    paths: Sequence[str | Path], *, max_total_chars: int = 160_000,
    max_file_chars: int = 100_000, max_file_bytes: int = 25 * 1024 * 1024,
    max_files: int = 10,
) -> DocumentBundle:
    """Read only the supplied local files; reject limits instead of truncating."""
    limits = (max_total_chars, max_file_chars, max_file_bytes, max_files)
    if any(isinstance(value, bool) or not isinstance(value, int) or value <= 0 for value in limits):
        raise DocumentError("Лимиты документов должны быть положительными целыми числами.")
    if len(paths) > max_files:
        raise DocumentError(f"Допускается не более {max_files} документов.")
    bundle = DocumentBundle()
    total_chars = 0
    seen: set[Path] = set()
    for number, supplied_path in enumerate(paths, 1):
        path = Path(supplied_path).expanduser()
        try:
            resolved = path.resolve(strict=True)
            if not resolved.is_file():
                raise DocumentError(f"«{path.name}» не является файлом.")
            if resolved in seen:
                raise DocumentError(f"Файл «{path.name}» передан повторно.")
            seen.add(resolved)
            if resolved.stat().st_size > max_file_bytes:
                raise DocumentError(f"«{path.name}» превышает лимит {max_file_bytes} байт на файл.")
            document_id = f"doc{number}"
            suffix = resolved.suffix.lower()
            warnings: list[str] = []
            if suffix in {".txt", ".md"}:
                data = resolved.read_bytes()
                if len(data) > max_file_bytes:
                    raise DocumentError(f"«{path.name}» превышает лимит {max_file_bytes} байт на файл.")
                try:
                    text = data.decode("utf-8-sig").strip()
                except UnicodeDecodeError as exc:
                    raise DocumentError(f"Сохраните «{path.name}» в кодировке UTF-8.") from exc
                chunks = [SourceChunk(f"{document_id}:text", resolved.name, "text", text)] if text else []
            elif suffix == ".docx":
                # Bound expansion before python-docx opens the ZIP container.
                try:
                    with ZipFile(resolved) as archive:
                        if sum(info.file_size for info in archive.infolist()) > max_file_bytes * 8:
                            raise DocumentError(f"Распакованный DOCX «{path.name}» слишком велик.")
                except BadZipFile as exc:
                    raise DocumentError(f"«{path.name}» не является корректным DOCX.") from exc
                chunks = _extract_docx(resolved, document_id)
                warnings.append(
                    f"«{resolved.name}»: извлечены абзацы и таблицы основного текста; "
                    "рисунки, диаграммы, колонтитулы и текстовые поля не анализировались."
                )
            elif suffix == ".pdf":
                chunks, warnings = _extract_pdf(resolved, document_id)
                warnings.append(
                    f"«{resolved.name}»: извлечён текстовый слой PDF; "
                    "изображения, схемы и структура таблиц не анализировались."
                )
            else:
                raise DocumentError(f"Неподдерживаемый формат «{path.name}». Используйте TXT, MD, DOCX или PDF.")
            if not chunks:
                raise DocumentError(f"В документе «{path.name}» нет текста.")
            file_chars = sum(len(chunk.text) for chunk in chunks)
            if file_chars > max_file_chars:
                raise DocumentError(
                    f"«{path.name}» содержит {file_chars} символов при лимите {max_file_chars}. "
                    "Выберите нужный фрагмент или явно увеличьте лимит. Текст не обрезан."
                )
            total_chars += file_chars
            if total_chars > max_total_chars:
                raise DocumentError(
                    f"Общий объём документов {total_chars} символов превышает лимит {max_total_chars}. "
                    "Сократите входные документы или явно увеличьте лимит. Текст не обрезан."
                )
            bundle.chunks.extend(chunks)
            bundle.warnings.extend(warnings)
        except DocumentError:
            raise
        except FileNotFoundError as exc:
            raise DocumentError(f"Документ «{path.name}» не найден.") from exc
        except OSError as exc:
            raise DocumentError(f"Не удалось открыть документ «{path.name}».") from exc
    return bundle
