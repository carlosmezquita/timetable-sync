from __future__ import annotations

from urllib.parse import urlparse

import httpx

from .models import SyncError


def download(url, client=None):
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password:
        raise SyncError("The timetable feed must use an HTTPS URL without embedded credentials.")
    client = client or httpx.Client(timeout=30, follow_redirects=False)
    try:
        with client.stream("GET", url, headers={"Cache-Control": "no-cache"}) as response:
            if response.status_code != 200:
                raise SyncError(
                    f"Timetable download returned HTTP {response.status_code}; no calendar changes were made."
                )
            raw = bytearray()
            for chunk in response.iter_bytes():
                raw.extend(chunk)
                if len(raw) > 5_000_000:
                    raise SyncError("Timetable feed exceeds the size limit.")
            return bytes(raw)
    except httpx.HTTPError as exc:
        raise SyncError("Timetable download failed; no calendar changes were made.") from exc
