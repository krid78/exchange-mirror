"""HTTP policy with injected responses: no DNS, sockets or live feeds."""

from email.message import Message
from io import BytesIO
from unittest import TestCase
from unittest.mock import patch
from urllib.request import Request

from ics_http import FeedDownloadError, _SameOriginRedirect, fetch_ics
from test_source import calendar, EVENT


URL = "https://calendar.example.invalid/feed?ShareKey=DO-NOT-LOG"
BODY = calendar(EVENT).encode()


class Response(BytesIO):
    def __init__(self, body=BODY, *, status=200, media="text/calendar", length=None, url=URL):
        super().__init__(body)
        self.status, self.url = status, url
        self.headers = Message()
        self.headers["Content-Type"] = media
        self.headers["ETag"] = '"synthetic"'
        if length is not None:
            self.headers["Content-Length"] = str(length)

    def geturl(self):
        return self.url


class Opener:
    def __init__(self, response):
        self.response = response

    def open(self, request, timeout):
        self.request, self.timeout = request, timeout
        return self.response


class HttpTests(TestCase):
    def test_complete_download_and_metadata(self):
        opener = Opener(Response(length=len(BODY)))
        snapshot = fetch_ics(URL, opener=opener)
        self.assertEqual(snapshot.data, BODY)
        self.assertEqual(snapshot.source_host, "calendar.example.invalid")
        self.assertEqual(snapshot.etag, '"synthetic"')
        self.assertIsNotNone(snapshot.fetched_at.utcoffset())
        self.assertEqual(len(snapshot.digest), 64)
        self.assertFalse(opener.request.has_header("If-none-match"))
        self.assertEqual(opener.timeout, 30)

    def test_invalid_status_type_body_length_and_compression(self):
        for response in (Response(status=304), Response(status=500), Response(status=206),
                         Response(media="text/html"), Response(b"<html>login</html>"),
                         Response(BODY[:-20]), Response(length=len(BODY) + 1),
                         Response(length="invalid"), Response(length=-1),
                         Response(url="https://other.example.invalid/feed")):
            with self.subTest(response=response), self.assertRaises(FeedDownloadError):
                fetch_ics(URL, opener=Opener(response))
        response = Response()
        response.headers["Content-Encoding"] = "gzip"
        with self.assertRaises(FeedDownloadError):
            fetch_ics(URL, opener=Opener(response))

    def test_size_and_deadline_limits(self):
        for response in (Response(), Response(length=len(BODY))):
            with self.assertRaises(FeedDownloadError):
                fetch_ics(URL, max_bytes=10, opener=Opener(response))
        with patch("ics_http.time.monotonic", side_effect=[0, 31]):
            with self.assertRaises(FeedDownloadError):
                fetch_ics(URL, opener=Opener(Response()))

    def test_errors_do_not_echo_sensitive_url(self):
        for error in (TimeoutError(URL), OSError(URL), ValueError(URL)):
            with patch.object(Opener, "open", side_effect=error):
                with self.assertRaises(FeedDownloadError) as caught:
                    fetch_ics(URL, opener=Opener(Response()))
                self.assertNotIn("DO-NOT-LOG", str(caught.exception))
                self.assertTrue(caught.exception.__suppress_context__)

    def test_redirect_policy(self):
        handler = _SameOriginRedirect()
        req = Request(URL)
        redirected = handler.redirect_request(req, None, 302, "redirect", {},
                                              "https://calendar.example.invalid/new")
        self.assertEqual(redirected.host, req.host)
        for target in ("http://calendar.example.invalid/feed", "https://other.example.invalid/",
                       "https://calendar.example.invalid:444/", "https://user:pass@calendar.example.invalid/"):
            with self.subTest(target=target), self.assertRaises(FeedDownloadError):
                handler.redirect_request(req, None, 302, "redirect", {}, target)

    def test_invalid_request_options(self):
        for url in ("http://example.invalid/", "https://user:pass@example.invalid/",
                    "https://example.invalid/#fragment", "https://example.invalid/\n"):
            with self.subTest(url=url), self.assertRaises(FeedDownloadError):
                fetch_ics(url, opener=Opener(Response()))
        for options in ({"max_bytes": 0}, {"timeout": float("inf")}, {"timeout": 0}):
            with self.subTest(options=options), self.assertRaises(ValueError):
                fetch_ics(URL, opener=Opener(Response()), **options)

    def test_default_transport_uses_verifying_tls_context(self):
        with patch("ics_http.ssl.create_default_context", wraps=__import__("ssl").create_default_context) as context:
            with patch("ics_http.build_opener", return_value=Opener(Response())):
                fetch_ics(URL)
            context.assert_called_once()
