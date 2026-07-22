from __future__ import annotations

from dataclasses import dataclass, field
import json
from pathlib import Path


@dataclass(slots=True)
class ConversionReport:
    source: str
    destination: str | None = None
    warnings: list[dict[str, object]] = field(default_factory=list)

    def warning(self, code: str, message: str, **details: object) -> None:
        self.warnings.append({"code": code, "message": message, **details})

    def to_dict(self) -> dict[str, object]:
        return {"source": self.source, "destination": self.destination,
                "warning_count": len(self.warnings), "warnings": self.warnings}


def write_json_file(path: str | Path, value: object) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.name}.tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(target)
