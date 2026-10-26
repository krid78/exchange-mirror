"""Bounded HTTPS publication downloads; no credentials, cache or Outlook access."""

from dataclasses import dataclass
from datetime import datetime, timezone
from hashlib import sha256
import math
import ssl
import time
from urllib.parse import urljoin, urlsplit
from urllib.request import HTTPRedirectHandler, HTTPSHandler, Request, build_opener

from ics_source import _check_envelope


class FeedDownloadError(ValueError):
    """Safe diagnostic which never includes the feed URL or raw response."""


def validate_feed_url(url):
    """Return the HTTPS origin; reject credentials and ambiguous URL syntax."""
    if not isinstance(url, str) or any(ord(c) <= 32 or ord(c) == 127 for c in url):
        raise FeedDownloadError("Expected an HTTPS URL without whitespace")
    try:
        parts = urlsplit(url)
        if (parts.scheme != "https" or not parts.hostname or parts.username is not None
                or parts.password is not None or parts.fragment):
            raise ValueError
        return parts.hostname.lower(), parts.port or 443
    except ValueError:
        raise FeedDownloadError("Expected an HTTPS URL without credentials or fragment") from None


class _SameOriginRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if validate_feed_url(urljoin(req.full_url, newurl)) != validate_feed_url(req.full_url):
            raise FeedDownloadError("Cross-origin redirect refused")
        return super().redirect_request(req, fp, code, msg, headers, newurl)


@dataclass(frozen=True)
class FeedSnapshot:
    data: bytes
    source_host: str
    fetched_at: datetime
    digest: str
    etag: str | None = None
    last_modified: str | None = None


def fetch_ics(url, *, timeout=30.0, max_bytes=5_000_000, opener=None):
    """Read a complete response or fail, never substitute an empty calendar.

    An injected opener must obey the same HTTPS/redirect policy; intended for
    deterministic transport tests. This validates the envelope, not all events.
    """
    origin = validate_feed_url(url)
    if isinstance(timeout, bool) or not isinstance(timeout, (int, float)) or not math.isfinite(timeout) or timeout <= 0:
        raise ValueError("timeout must be finite and positive")
    if isinstance(max_bytes, bool) or not isinstance(max_bytes, int) or max_bytes <= 0:
        raise ValueError("max_bytes must be a positive integer")
    try:
        if opener is None:
            opener = build_opener(_SameOriginRedirect(), HTTPSHandler(context=ssl.create_default_context()))
        request = Request(url, headers={"Accept": "text/calendar", "Accept-Encoding": "identity",
                                       "User-Agent": "Exchange-Mirror-ICS-Importer"})
        deadline = time.monotonic() + timeout
        with opener.open(request, timeout=timeout) as response:
            if response.status != 200:
                raise FeedDownloadError("Feed requires a full HTTP 200 response")
            if validate_feed_url(response.geturl()) != origin:
                raise FeedDownloadError("Unexpected response origin")
            media_type = response.headers.get_content_type()
            if media_type not in ("text/calendar", "text/plain", "application/octet-stream"):
                raise FeedDownloadError("Response is not a calendar download")
            if response.headers.get("Content-Encoding", "identity").lower() != "identity":
                raise FeedDownloadError("Unexpected response compression")
            length = response.headers.get("Content-Length")
            expected = int(length) if length is not None else None
            if expected is not None and (expected < 0 or expected > max_bytes):
                raise FeedDownloadError("Invalid or oversized response length")
            chunks, size = [], 0
            while True:
                chunk = response.read(min(65536, max_bytes + 1 - size))
                if time.monotonic() > deadline:
                    raise FeedDownloadError("Feed download deadline exceeded")
                if not chunk:
                    break
                chunks.append(chunk)
                size += len(chunk)
                if size > max_bytes:
                    raise FeedDownloadError("Feed exceeds size limit")
            if expected is not None and size != expected:
                raise FeedDownloadError("Incomplete response body")
            data = b"".join(chunks)
            _check_envelope(data.decode("utf-8-sig"))
            return FeedSnapshot(data, origin[0], datetime.now(timezone.utc), sha256(data).hexdigest(),
                                response.headers.get("ETag"), response.headers.get("Last-Modified"))
    except FeedDownloadError:
        raise
    except Exception:
        # HTTPError/URLError/SSL errors can echo URLs including the query token.
        raise FeedDownloadError("Calendar download failed or response was invalid") from None
