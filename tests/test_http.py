import httpx
import pytest

from timetable_sync.feed import download
from timetable_sync.graph import KEY_PROPERTY, SOURCE_PROPERTY, Graph, verify_owned
from timetable_sync.models import SyncError


def test_graph_throttle_retries_same_creation_identifier():
    calls = []

    def handler(request):
        calls.append(request)
        return (
            httpx.Response(429, headers={"Retry-After": "1"})
            if len(calls) == 1
            else httpx.Response(201, json={"id": "created"})
        )

    graph = Graph(
        lambda: "token",
        "test",
        client=httpx.Client(transport=httpx.MockTransport(handler)),
        sleep=lambda _: None,
    )
    assert graph.create({"showAs": "busy"}, "key", "same-transaction")["id"] == "created"
    assert calls[0].content == calls[1].content
    assert len(calls) == 2


def test_pagination_cannot_send_token_to_another_host():
    graph = Graph(lambda: "secret", "test")
    with pytest.raises(SyncError):
        graph.request("GET", "https://attacker.example/v1.0/events")


@pytest.mark.parametrize(
    "event",
    [
        {"id": "x"},
        {
            "singleValueExtendedProperties": [
                {"id": SOURCE_PROPERTY, "value": "test"},
                {"id": KEY_PROPERTY, "value": "key"},
            ],
            "attendees": [{"emailAddress": {"address": "x"}}],
        },
    ],
)
def test_ownership_and_meeting_guard(event):
    with pytest.raises(SyncError):
        verify_owned(event, "test", "key")


def test_etag_conflict_aborts():
    graph = Graph(
        lambda: "token",
        "test",
        client=httpx.Client(transport=httpx.MockTransport(lambda _: httpx.Response(412))),
    )
    with pytest.raises(SyncError, match="changed during sync"):
        graph.request("PATCH", "/me/calendar/events/x", json={})


def test_feed_errors_redact_private_url():
    url = "https://example.com/private-feed-token/calendar.ics"

    def fail(request):
        raise httpx.ConnectError(url, request=request)

    with pytest.raises(SyncError) as error:
        download(url, httpx.Client(transport=httpx.MockTransport(fail)))
    assert "private-feed-token" not in str(error.value)


def test_no_redirect_to_login_page():
    with pytest.raises(SyncError, match="HTTP 302"):
        download(
            "https://example.com/calendar.ics",
            httpx.Client(
                transport=httpx.MockTransport(
                    lambda _: httpx.Response(302, headers={"Location": "/login"})
                )
            ),
        )
