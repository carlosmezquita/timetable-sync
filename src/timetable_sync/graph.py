from __future__ import annotations

import time
from urllib.parse import quote, urlparse

import httpx

from .models import SyncError

ROOT = "https://graph.microsoft.com/v1.0"
GUID = "a63d1a09-d00b-48f3-8ba0-8b3d2caefec2"
SOURCE_PROPERTY = f"String {{{GUID}}} Name TimetableSyncSource"
KEY_PROPERTY = f"String {{{GUID}}} Name TimetableSyncKey"


def properties(event):
    return {
        item["id"]: item.get("value", "") for item in event.get("singleValueExtendedProperties", [])
    }


def verify_owned(event, source, key=None):
    props = properties(event)
    if props.get(SOURCE_PROPERTY) != source or (key and props.get(KEY_PROPERTY) != key):
        raise SyncError("Refusing to change an event without the expected ownership marker.")
    if event.get("attendees") or event.get("isOnlineMeeting"):
        raise SyncError("A managed event has become a meeting. Review it before syncing.")


class Graph:
    def __init__(self, token, source, calendar="primary", client=None, sleep=time.sleep):
        self.token = token
        self.source = source
        self.client = client or httpx.Client(timeout=30)
        self.sleep = sleep
        self.path = (
            "/me/calendar/events"
            if calendar == "primary"
            else "/me/calendars/" + quote(calendar, safe="") + "/events"
        )

    def request(self, method, path, **kwargs):
        url = path if path.startswith("https://") else ROOT + path
        parsed = urlparse(url)
        if (
            parsed.scheme != "https"
            or parsed.netloc != "graph.microsoft.com"
            or not parsed.path.startswith("/v1.0/")
        ):
            raise SyncError("Refusing an unexpected Microsoft Graph pagination URL.")
        headers = dict(kwargs.pop("headers", {}))
        for attempt in range(4):
            headers["Authorization"] = "Bearer " + self.token()
            try:
                response = self.client.request(method, url, headers=headers, **kwargs)
            except httpx.HTTPError as exc:
                raise SyncError(
                    "Microsoft Graph could not be reached. Retry the sync; its creation identifiers are retained."
                ) from exc
            if response.status_code in {429, 502, 503, 504} and attempt < 3:
                try:
                    delay = min(max(float(response.headers.get("Retry-After", 2**attempt)), 1), 30)
                except ValueError:
                    delay = 2**attempt
                self.sleep(delay)
                continue
            if response.status_code == 412:
                raise SyncError(
                    "An Outlook event changed during sync. Retry to preserve the latest availability choice."
                )
            if response.status_code >= 400:
                raise SyncError(
                    f"Microsoft Graph returned HTTP {response.status_code}. Check calendar access or retry."
                )
            return response.json() if response.content else {}
        raise SyncError("Microsoft Graph retries exhausted.")

    def list_owned(self):
        source = self.source.replace("'", "''")
        params = {
            "$filter": f"singleValueExtendedProperties/Any(ep: ep/id eq '{SOURCE_PROPERTY}' and ep/value eq '{source}')",
            "$expand": "singleValueExtendedProperties",
            "$top": "100",
        }
        result = {}
        path = self.path
        pages = 0
        while path:
            page = self.request("GET", path, params=params)
            params = None
            for event in page.get("value", []):
                verify_owned(event, self.source)
                key = properties(event).get(KEY_PROPERTY)
                if not key or key in result:
                    raise SyncError(
                        "A managed event has a missing or duplicate identity. Review Outlook before syncing."
                    )
                result[key] = event
            path = page.get("@odata.nextLink")
            pages += 1
            if pages > 100:
                raise SyncError("Too many managed calendar pages.")
        return result

    def create(self, payload, key, transaction):
        body = dict(payload)
        body["transactionId"] = transaction
        body["singleValueExtendedProperties"] = [
            {"id": SOURCE_PROPERTY, "value": self.source},
            {"id": KEY_PROPERTY, "value": key},
        ]
        return self.request("POST", self.path, json=body)

    def change(self, event, key, payload=None):
        verify_owned(event, self.source, key)
        etag = event.get("@odata.etag")
        if not etag:
            raise SyncError("Outlook did not provide an event version; refusing an unsafe update.")
        path = self.path + "/" + quote(event["id"], safe="")
        return self.request(
            "DELETE" if payload is None else "PATCH",
            path,
            headers={"If-Match": etag},
            **({"json": payload} if payload else {}),
        )
