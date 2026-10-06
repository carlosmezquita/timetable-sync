# Architecture

Pipeline: fetch -> validate/expand -> reconcile -> Outlook writes -> checkpoint.

SQLite stores the last applied occurrence, Outlook ID, availability preference and missing-event evidence. The parser uses recurrence family and original date for CMIS identity, and UID plus original recurrence time for generic feeds. CMIS numbered UID prefixes describe segments, not individual classes. Conflicting identities abort parsing.

Plans precede writes. Creation transaction IDs are committed before API calls. Ownership discovery recovers an accepted creation whose response was lost. Successful individual writes are checkpointed immediately; missing-event evidence and sync health advance only after the write pass succeeds.

Appointments carry source and key extended properties. The worker lists its own source, rejects duplicate identities and refuses to modify meetings. PATCH/DELETE use server ETags. Availability preferences remain separate from timetable content.

University data owns title/time/room/description; you own availability. Unrelated personal events remain in Outlook. No bidirectional calendar or Google API sync is implemented.

## Limits and tradeoffs

CMIS does not provide an independent cancellation channel in the observed feed. Missing events need confirmation; ambiguous date changes need an identity link. Empty schedules are conservatively rejected. Changed generic UIDs or CMIS internal IDs cannot always preserve preferences.

OAuth refresh can be revoked or demand interactive login. Monitor stale syncs. Creation recovery depends on Graph transaction handling and ownership discovery; duplicate ownership stops the next sync for review.

SQLite contains private timetable details. Protect it and backups. The current interface is a CLI, with no web dashboard or built-in notification transport.
