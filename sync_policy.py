"""Pure synchronization rules; no Outlook or configuration-file access."""

from enum import Enum


class SyncDecision(Enum):
    NORMAL = "normal"
    EXCLUDE = "exclude"
    ANONYMIZE = "anonymize"


def parse_categories(value, separator):
    """Parse complete names using Outlook's regional list separator."""
    if not separator:
        raise ValueError("Category separator must not be empty")
    return [name.strip() for name in (value or "").split(separator) if name.strip()]


def decide_sync(categories, *, series_categories=(),
                do_not_sync_category="DoNotSync", secret_category="Secret",
                anonymize_all=False):
    """Combine already-resolved occurrence and master categories.

    Callers must resolve required series data successfully before calling;
    missing or unreadable master data must not be represented as an empty list.
    """
    names = {name.strip().casefold() for name in (*categories, *series_categories)}
    if do_not_sync_category.strip().casefold() in names:
        return SyncDecision.EXCLUDE
    if anonymize_all or secret_category.strip().casefold() in names:
        return SyncDecision.ANONYMIZE
    return SyncDecision.NORMAL


def build_target_fields(source_fields, *, subject_prefix, mirror_category,
                        separator, decision=SyncDecision.NORMAL,
                        anonymized_subject="Belegt"):
    """Calculate the values used by both comparison and writing.

    Recipient and formatted-body cleanup belongs to the COM adapter;
    these fields alone are not a complete privacy implementation.
    """
    if decision is SyncDecision.EXCLUDE:
        raise ValueError("Excluded appointments have no target fields")
    target = dict(source_fields)
    if decision is SyncDecision.ANONYMIZE:
        target.update(Subject=anonymized_subject, Body="", Location="",
                      Categories="", RequiredAttendees="", OptionalAttendees="",
                      Resources="")
    if "Subject" in target and subject_prefix:
        subject = target["Subject"] or ""
        if not subject.startswith(subject_prefix):
            target["Subject"] = subject_prefix + subject
    if "Categories" in target:
        categories = parse_categories(target["Categories"], separator)
        if mirror_category and mirror_category.casefold() not in {
            name.casefold() for name in categories
        }:
            categories.append(mirror_category)
        target["Categories"] = (separator + " ").join(categories)
    # Apply fixed Outlook flags last, after all copied appointment fields.
    target.pop("ReminderSet", None)
    target["ReminderSet"] = False
    target["MeetingStatus"] = 0
    return target


def field_values_equal(field, expected, actual, separator):
    if field == "Body":
        return (expected or "").rstrip() == (actual or "").rstrip()
    if field == "Categories":
        return {name.casefold() for name in parse_categories(expected, separator)} == {
            name.casefold() for name in parse_categories(actual, separator)
        }
    return expected == actual
