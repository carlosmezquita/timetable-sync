from __future__ import annotations

DEFAULTS = {"self directive studies": "free", "self directed studies": "free"}


def availability(session, rules, override=None):
    if override:
        return override
    selectors = [
        "event:" + session.key,
        "family:" + session.family,
        "module-activity:" + session.module.casefold() + "|" + session.activity.casefold(),
        "module:" + session.module.casefold(),
        "activity:" + session.activity.casefold(),
        "default",
    ]
    for selector in selectors:
        if selector in rules:
            return rules[selector]
    return DEFAULTS.get(session.activity.casefold(), "busy")
