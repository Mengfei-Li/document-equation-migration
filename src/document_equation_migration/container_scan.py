from __future__ import annotations

import hashlib
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path


DOCX_STORY_TYPES = {
    "word/document.xml": "main",
    "word/footnotes.xml": "footnote",
    "word/endnotes.xml": "endnote",
    "word/comments.xml": "comment",
}


@dataclass(slots=True)
class StoryPartScan:
    part_path: str
    story_type: str
    omml_count: int = 0
    ole_object_count: int = 0
    odf_math_count: int = 0
    field_code_count: int = 0
    graphic_reference_count: int = 0

    def to_dict(self) -> dict[str, int | str]:
        return {
            "part_path": self.part_path,
            "story_type": self.story_type,
            "omml_count": self.omml_count,
            "ole_object_count": self.ole_object_count,
            "odf_math_count": self.odf_math_count,
            "field_code_count": self.field_code_count,
            "graphic_reference_count": self.graphic_reference_count,
        }


@dataclass(slots=True)
class ContainerScanResult:
    input_path: str
    input_sha256: str
    container_format: str
    package_kind: str
    entry_count: int
    entries: list[str] = field(default_factory=list)
    story_parts: list[StoryPartScan] = field(default_factory=list)
    embedding_targets: list[str] = field(default_factory=list)
    media_targets: list[str] = field(default_factory=list)
    relationship_parts: list[str] = field(default_factory=list)
    object_parts: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, object]:
        return {
            "input_path": self.input_path,
            "input_sha256": self.input_sha256,
            "container_format": self.container_format,
            "package_kind": self.package_kind,
            "entry_count": self.entry_count,
            "entries": self.entries,
            "story_parts": [item.to_dict() for item in self.story_parts],
            "embedding_targets": self.embedding_targets,
            "media_targets": self.media_targets,
            "relationship_parts": self.relationship_parts,
            "object_parts": self.object_parts,
            "notes": self.notes,
        }


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _story_type_for_docx_part(part_path: str) -> str:
    if part_path in DOCX_STORY_TYPES:
        return DOCX_STORY_TYPES[part_path]
    if part_path.startswith("word/header"):
        return "header"
    if part_path.startswith("word/footer"):
        return "footer"
    return "other"


def _count(data: bytes, pattern: bytes) -> int:
    return len(re.findall(re.escape(pattern), data))


def _scan_docx(path: Path) -> ContainerScanResult:
    with zipfile.ZipFile(path) as zf:
        entries = sorted(zf.namelist())
        story_candidates = [
            name
            for name in entries
            if name.startswith("word/")
            and name.endswith(".xml")
            and "/_rels/" not in name
            and (
                name == "word/document.xml"
                or name.startswith("word/header")
                or name.startswith("word/footer")
                or name in {"word/footnotes.xml", "word/endnotes.xml", "word/comments.xml"}
            )
        ]
        story_parts: list[StoryPartScan] = []
        for name in story_candidates:
            data = zf.read(name)
            part = StoryPartScan(
                part_path=name,
                story_type=_story_type_for_docx_part(name),
                omml_count=_count(data, b"<m:oMath") + _count(data, b"<m:oMathPara"),
                ole_object_count=_count(data, b"<o:OLEObject"),
                field_code_count=_count(data, b"EMBED Equation"),
                graphic_reference_count=_count(data, b"<v:imagedata") + _count(data, b"<a:blip"),
            )
            if part.omml_count or part.ole_object_count or part.field_code_count or part.graphic_reference_count:
                story_parts.append(part)
        return ContainerScanResult(
            input_path=str(path),
            input_sha256=_sha256_file(path),
            container_format="docx",
            package_kind="ooxml-zip",
            entry_count=len(entries),
            entries=entries,
            story_parts=story_parts,
            embedding_targets=[name for name in entries if name.startswith("word/embeddings/")],
            media_targets=[name for name in entries if name.startswith("word/media/")],
            relationship_parts=[name for name in entries if name.endswith(".rels")],
        )


def _scan_odf_zip(path: Path, container_format: str) -> ContainerScanResult:
    with zipfile.ZipFile(path) as zf:
        entries = sorted(zf.namelist())
        story_parts: list[StoryPartScan] = []
        object_parts: list[str] = []
        for name in entries:
            if not name.endswith(".xml"):
                continue
            if name == "content.xml" or name.endswith("/content.xml"):
                data = zf.read(name)
                part = StoryPartScan(
                    part_path=name,
                    story_type="main" if name == "content.xml" else "object",
                    odf_math_count=_count(data, b"<math:math"),
                    graphic_reference_count=_count(data, b"<draw:object") + _count(data, b"<draw:object-ole"),
                )
                if part.odf_math_count or part.graphic_reference_count:
                    story_parts.append(part)
                if name.endswith("/content.xml") and name != "content.xml":
                    object_parts.append(name)
        return ContainerScanResult(
            input_path=str(path),
            input_sha256=_sha256_file(path),
            container_format=container_format,
            package_kind="odf-zip",
            entry_count=len(entries),
            entries=entries,
            story_parts=story_parts,
            object_parts=object_parts,
            notes=["ODF scan treats content.xml and object subdocuments as primary scan surfaces."],
        )


def _scan_fodt(path: Path) -> ContainerScanResult:
    data = path.read_bytes()
    story_parts = [
        StoryPartScan(
            part_path=path.name,
            story_type="main",
            odf_math_count=_count(data, b"<math:math"),
            graphic_reference_count=_count(data, b"<draw:object") + _count(data, b"<draw:object-ole"),
        )
    ]
    return ContainerScanResult(
        input_path=str(path),
        input_sha256=_sha256_file(path),
        container_format="fodt",
        package_kind="flat-xml",
        entry_count=1,
        entries=[path.name],
        story_parts=story_parts,
        notes=["Flat ODF scan operates on the top-level XML file."],
    )


def scan_container(path: str | Path) -> ContainerScanResult:
    resolved = Path(path).resolve()
    suffix = resolved.suffix.lower()
    if suffix == ".docx":
        return _scan_docx(resolved)
    if suffix == ".odt":
        return _scan_odf_zip(resolved, "odt")
    if suffix == ".fodt":
        return _scan_fodt(resolved)
    raise ValueError(f"Unsupported input format: {resolved.suffix}")
