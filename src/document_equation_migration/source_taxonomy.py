from __future__ import annotations

from enum import StrEnum


class SourceFamily(StrEnum):
    MATHTYPE_OLE = "mathtype-ole"
    OMML_NATIVE = "omml-native"
    EQUATION_EDITOR_3_OLE = "equation-editor-3-ole"
    AXMATH_OLE = "axmath-ole"
    ODF_NATIVE = "odf-native"
    LIBREOFFICE_TRANSFORMED = "libreoffice-transformed"
    GRAPHIC_FALLBACK = "graphic-fallback"
    UNKNOWN_OLE = "unknown-ole"


class SourceRole(StrEnum):
    NATIVE_SOURCE = "native-source"
    TRANSFORMED_SOURCE = "transformed-source"
    GENERATED_OUTPUT = "generated-output"
    PREVIEW_ONLY = "preview-only"


PRIMARY_SOURCE_FAMILIES = frozenset(
    {
        SourceFamily.MATHTYPE_OLE,
        SourceFamily.OMML_NATIVE,
        SourceFamily.EQUATION_EDITOR_3_OLE,
        SourceFamily.ODF_NATIVE,
    }
)

BRIDGE_SOURCE_FAMILIES = frozenset(
    {
        SourceFamily.AXMATH_OLE,
        SourceFamily.LIBREOFFICE_TRANSFORMED,
    }
)

FALLBACK_SOURCE_FAMILIES = frozenset(
    {
        SourceFamily.GRAPHIC_FALLBACK,
        SourceFamily.UNKNOWN_OLE,
    }
)

SOURCE_FAMILY_DESCRIPTIONS: dict[SourceFamily, str] = {
    SourceFamily.MATHTYPE_OLE: "DOCX OLE object with MathType / MTEF payload",
    SourceFamily.OMML_NATIVE: "WordprocessingML native OMML formula",
    SourceFamily.EQUATION_EDITOR_3_OLE: "Legacy Microsoft Equation Editor 3.0 OLE object",
    SourceFamily.AXMATH_OLE: "AxMath OLE object or Word add-in artifact",
    SourceFamily.ODF_NATIVE: "ODF-native formula object",
    SourceFamily.LIBREOFFICE_TRANSFORMED: "LibreOffice bridge or transformed output",
    SourceFamily.GRAPHIC_FALLBACK: "Graphic-only preview without native payload",
    SourceFamily.UNKNOWN_OLE: "Unclassified OLE object retained for review",
}


def normalize_source_family(value: str | SourceFamily) -> SourceFamily:
    if isinstance(value, SourceFamily):
        return value
    return SourceFamily(value)


def normalize_source_role(value: str | SourceRole) -> SourceRole:
    if isinstance(value, SourceRole):
        return value
    return SourceRole(value)
