"""Validated, lossless client data; no floats or integer identifiers."""
from dataclasses import dataclass, asdict
from datetime import date
from decimal import Decimal, localcontext
import json
import os
from pathlib import Path
import re
import tempfile


class ValidationError(ValueError):
    def __init__(self, message, **values):
        self.message = message
        self.values = values
        super().__init__(message.format(**values))


def text_value(value, label, limit=500):
    if not isinstance(value, str) or not value.strip():
        raise ValidationError("{label} must be a non-empty string.", label=label)
    if len(value) > limit or any(ord(c) < 32 or 0xD800 <= ord(c) <= 0xDFFF or ord(c) in (0xFFFE, 0xFFFF) for c in value):
        raise ValidationError("{label} must be one line, at most {limit} characters, with valid text.", label=label, limit=limit)
    return value


def decimal_value(value, label):
    text_value(value, label, 24)
    if not re.fullmatch(r"[0-9]{1,15}(?:\.[0-9]{1,8})?", value):
        raise ValidationError("{label}: use a non-negative decimal (up to 15 integer and 8 fractional digits).", label=label)
    return Decimal(value)


@dataclass(frozen=True)
class LandRecord:
    khasra: str
    area: str
    lagan: str

    def __post_init__(self):
        text_value(self.khasra, "Khasra number", 100)
        decimal_value(self.area, "Area")
        decimal_value(self.lagan, "Lagan")


@dataclass(frozen=True)
class ClientData:
    client_name: str
    father_name: str
    address: str
    village: str
    tehsil: str
    district: str
    document_date: str
    age: str
    caste: str
    bahi: str
    records: tuple
    text_mode: str = "legacy_devlys"
    unicode_font: str = "DevLys 010"

    def __post_init__(self):
        labels = {"client_name": "Client name", "father_name": "Father's name", "address": "Address", "village": "Village", "tehsil": "Tehsil", "district": "District"}
        for key, label in labels.items():
            text_value(getattr(self, key), label)
        text_value(self.bahi, "Bahi number", 100)
        # Empty new fields are allowed while opening an older saved client.
        # Saving or generating requires completing them; never invent an age/caste.
        if not isinstance(self.age, str) or (self.age != "" and (not re.fullmatch(r"[0-9]{1,3}", self.age) or not 1 <= int(self.age) <= 150)):
            raise ValidationError("Age must be a whole number of years between 1 and 150.")
        if self.caste != "":
            text_value(self.caste, "Caste")
        if not isinstance(self.document_date, str):
            raise ValidationError("Document date must be an ISO date string.")
        try:
            parsed = date.fromisoformat(self.document_date)
            if parsed.isoformat() != self.document_date:
                raise ValueError()
        except ValueError:
            raise ValidationError("Document date must be a real date in YYYY-MM-DD format.") from None
        if self.text_mode not in ("legacy_devlys", "unicode"):
            raise ValidationError("Choose legacy_devlys or unicode text mode.")
        if self.text_mode == "unicode" and self.unicode_font != "Mangal":
            raise ValidationError("The Unicode font must be Mangal. DevLys 010 requires pre-encoded text.")
        if not isinstance(self.records, (tuple, list)) or not 1 <= len(self.records) <= 500:
            raise ValidationError("Add between 1 and 500 land records before saving or generating.")
        if not all(isinstance(r, LandRecord) for r in self.records):
            raise ValidationError("Invalid land record.")
        object.__setattr__(self, "records", tuple(self.records))
        if self.text_mode == "legacy_devlys":
            values = [getattr(self, k) for k in ("client_name", "father_name", "address", "village", "tehsil", "district", "caste", "bahi")]
            values.extend(record.khasra for record in self.records)
            # Legacy glyph codes are Latin/Western punctuation, never Devanagari.
            # This does not claim to validate or translate DevLys glyph semantics.
            if any(any(ord(c) > 255 and not 0x2000 <= ord(c) <= 0x206F for c in value) for value in values):
                raise ValidationError("DevLys mode requires already encoded keyboard text, not Unicode Hindi. No automatic conversion is performed.")

    def require_complete(self):
        text_value(self.age, "Age")
        text_value(self.caste, "Caste")

    @property
    def document_font(self):
        return "Mangal" if self.text_mode == "unicode" else "DevLys 010"

    @property
    def totals(self):
        with localcontext() as ctx:
            ctx.prec = 40
            return (sum((Decimal(r.area) for r in self.records), Decimal(0)),
                    sum((Decimal(r.lagan) for r in self.records), Decimal(0)))

    def to_dict(self):
        result = asdict(self)
        result["records"] = [asdict(record) for record in self.records]
        result["schema_version"] = 2
        return result

    @classmethod
    def from_dict(cls, raw):
        if not isinstance(raw, dict) or type(raw.get("schema_version")) is not int or raw["schema_version"] not in (1, 2):
            raise ValidationError("Unsupported or missing JSON schema_version (expected 1 or 2).")
        allowed = set(cls.__dataclass_fields__) | {"schema_version"}
        if set(raw) - allowed:
            raise ValidationError("Unknown JSON fields: {fields}", fields=", ".join(sorted(set(raw) - allowed)))
        data = dict(raw)
        version = data.pop("schema_version")
        try:
            records = data.pop("records")
            if not isinstance(records, list):
                raise TypeError()
            if version == 1:
                if any(key in data for key in ("age", "caste", "bahi")):
                    raise ValidationError("Version 1 JSON cannot contain version 2 client fields.")
                bahis = [r["bahi"] for r in records]
                if not bahis or any(b != bahis[0] for b in bahis):
                    raise ValidationError("Old JSON has different Bahi numbers. A document requires one Bahi; resolve the numbers before importing.")
                data["bahi"] = bahis[0]
                data["age"] = ""
                data["caste"] = ""  # The previous app never collected this value.
                records = [{k: v for k, v in r.items() if k != "bahi"} for r in records]
            data["records"] = tuple(LandRecord(**r) for r in records)
            return cls(**data)
        except (TypeError, KeyError) as exc:
            raise ValidationError("JSON has missing fields or invalid record structure.") from exc


def load_client(path):
    with Path(path).open(encoding="utf-8") as stream:
        return ClientData.from_dict(json.load(stream))


def atomic_write(path, writer, overwrite=False):
    """Write beside destination, then publish; a failure leaves no partial file."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, name = tempfile.mkstemp(prefix=".landdocs-", dir=path.parent)
    os.close(fd)
    try:
        writer(Path(name))
        if overwrite:
            os.replace(name, path)
        else:
            os.link(name, path)  # Atomic create-only: never replace an existing DOCX.
    finally:
        if os.path.exists(name):
            os.unlink(name)


def save_client(client, path):
    client.require_complete()
    if Path(path).suffix.lower() != ".json":
        raise ValidationError("Client data must be saved to a .json file.")
    atomic_write(path, lambda p: p.write_text(json.dumps(client.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8"), overwrite=True)
