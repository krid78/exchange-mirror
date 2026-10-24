"""Pure synchronization rules; no Outlook or configuration-file access."""


def parse_categories(value, separator):
    """Parse complete names using Outlook's regional list separator."""
    if not separator:
        raise ValueError("Category separator must not be empty")
    return [name.strip() for name in (value or "").split(separator) if name.strip()]


def build_target_fields(source_fields, *, subject_prefix, mirror_category,
                        separator):
    """Calculate the normal mirror values for comparison and writing."""
    target = dict(source_fields)
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
