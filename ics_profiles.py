"""Validated per-user importer profiles with atomic JSON persistence."""

from dataclasses import asdict, dataclass, replace
from datetime import date
import json
import os
from pathlib import Path
import tempfile
from uuid import UUID

from ics_http import validate_feed_url


class ProfileError(ValueError):
    """Invalid profile data; never includes source URLs or raw JSON."""


def _text(value):
    if not isinstance(value, str) or not value.strip():
        raise ProfileError("Profile text fields must be nonempty strings")


@dataclass(frozen=True)
class Destination:
    store_id: str
    folder_id: str
    display_name: str

    def __post_init__(self):
        for value in (self.store_id, self.folder_id, self.display_name):
            _text(value)


@dataclass(frozen=True)
class ExportCoverage:
    start: date
    end: date
    confirmed: bool = False

    def __post_init__(self):
        if (type(self.start) is not date or type(self.end) is not date
                or self.end <= self.start or type(self.confirmed) is not bool):
            raise ProfileError("Invalid export coverage")


@dataclass(frozen=True)
class ImportProfile:
    profile_id: UUID
    name: str
    source_url: str
    destination: Destination
    floating_timezone: str
    window_start: date
    window_end: date
    source_coverage: ExportCoverage | None = None

    def __post_init__(self):
        if not isinstance(self.profile_id, UUID) or not isinstance(self.destination, Destination):
            raise ProfileError("Expected a UUID and concrete destination")
        _text(self.name)
        _text(self.floating_timezone)
        validate_feed_url(self.source_url)
        if self.source_coverage is not None and not isinstance(self.source_coverage, ExportCoverage):
            raise ProfileError("Invalid source coverage")
        if (type(self.window_start) is not date or type(self.window_end) is not date
                or self.window_end <= self.window_start):
            raise ProfileError("Profile window must contain increasing dates")


def update_profile(profile, *, source_url=None, destination=None,
                   confirm_source_change=False, confirm_destination_change=False):
    """Explicit mapping changes; confirmation is supplied by the future UI/caller."""
    if source_url is not None and source_url != profile.source_url and confirm_source_change is not True:
        raise ProfileError("Source change requires explicit confirmation")
    if destination is not None and destination != profile.destination and confirm_destination_change is not True:
        raise ProfileError("Destination change requires explicit confirmation")
    return replace(profile, source_url=profile.source_url if source_url is None else source_url,
                   destination=profile.destination if destination is None else destination,
                   source_coverage=None if source_url is not None and source_url != profile.source_url
                   else profile.source_coverage)


def default_profiles_path():
    appdata = os.environ.get("APPDATA")
    if not appdata:
        raise ProfileError("APPDATA is unavailable; provide an explicit profiles path")
    return Path(appdata) / "Exchange Mirror" / "ics-importer" / "profiles.json"


def _unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ProfileError("Duplicate JSON field")
        result[key] = value
    return result


def _iso_date(value):
    if not isinstance(value, str):
        raise ProfileError("Expected an ISO date")
    parsed = date.fromisoformat(value)
    if parsed.isoformat() != value:
        raise ProfileError("Expected a YYYY-MM-DD date")
    return parsed


def _validate_profiles(profiles):
    profiles = tuple(profiles)
    if any(not isinstance(p, ImportProfile) for p in profiles):
        raise ProfileError("Expected importer profiles")
    if len({p.profile_id for p in profiles}) != len(profiles):
        raise ProfileError("Duplicate profile UUID")
    return profiles


def load_profiles(path=None):
    path = Path(path) if path is not None else default_profiles_path()
    try:
        if not path.exists():
            return ()
        with path.open("rb") as stream:
            raw = stream.read(1_000_001)
        if len(raw) > 1_000_000:
            raise ProfileError("Profiles file is too large")
        document = json.loads(raw.decode("utf-8"), object_pairs_hook=_unique_object)
        if (not isinstance(document, dict) or set(document) != {"version", "profiles"}
                or type(document["version"]) is not int or document["version"] != 1
                or not isinstance(document["profiles"], list)):
            raise ProfileError("Unsupported profiles document")
        profiles = []
        for record in document["profiles"]:
            fields = set(ImportProfile.__dataclass_fields__)
            if (not isinstance(record, dict) or not fields - {"source_coverage"} <= set(record)
                    or not set(record) <= fields):
                raise ProfileError("Invalid profile fields")
            values = dict(record)
            values["profile_id"] = UUID(values["profile_id"])
            values["destination"] = Destination(**values["destination"])
            values["window_start"] = _iso_date(values["window_start"])
            values["window_end"] = _iso_date(values["window_end"])
            coverage = values.get("source_coverage")
            if coverage is not None:
                if not isinstance(coverage, dict) or set(coverage) != {"start", "end", "confirmed"}:
                    raise ProfileError("Invalid source coverage fields")
                values["source_coverage"] = ExportCoverage(_iso_date(coverage["start"]),
                                                           _iso_date(coverage["end"]), coverage["confirmed"])
            profiles.append(ImportProfile(**values))
        return _validate_profiles(profiles)
    except ProfileError:
        raise
    except Exception:
        raise ProfileError("Cannot read or validate importer profiles") from None


def save_profiles(profiles, path=None):
    profiles = _validate_profiles(profiles)
    path = Path(path) if path is not None else default_profiles_path()
    records = []
    for profile in profiles:
        record = asdict(profile)
        record.update(profile_id=str(profile.profile_id), window_start=profile.window_start.isoformat(),
                      window_end=profile.window_end.isoformat())
        if profile.source_coverage is not None:
            record["source_coverage"].update(start=profile.source_coverage.start.isoformat(),
                                              end=profile.source_coverage.end.isoformat())
        records.append(record)
    payload = json.dumps({"version": 1, "profiles": records}, ensure_ascii=True, indent=2).encode("utf-8")
    if len(payload) > 1_000_000:
        raise ProfileError("Profiles file is too large")
    temporary = None
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=path.parent, prefix=".profiles-", suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, path)
    except Exception:
        raise ProfileError("Cannot save importer profiles; previous file was not replaced") from None
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
