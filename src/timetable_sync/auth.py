from __future__ import annotations

import os

import msal
from cryptography.fernet import Fernet, InvalidToken

from .models import SyncError

SCOPES = ["Calendars.ReadWrite"]


class Authentication:
    def __init__(self, config):
        if not config.client_id:
            raise SyncError(
                "Set TSYNC_CLIENT_ID to your Entra public-client application ID, then run auth."
            )
        if not config.token_key:
            raise SyncError("Set TSYNC_TOKEN_KEY to a Fernet key (the init command generates one).")
        try:
            self.cipher = Fernet(config.token_key.encode())
        except (ValueError, TypeError) as exc:
            raise SyncError("TSYNC_TOKEN_KEY is invalid.") from exc
        self.path = config.state_dir / "tokens.enc"
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.cache = msal.SerializableTokenCache()
        if self.path.exists():
            try:
                self.cache.deserialize(self.cipher.decrypt(self.path.read_bytes()).decode())
            except (InvalidToken, ValueError) as exc:
                raise SyncError("Cannot decrypt the token cache. Check TSYNC_TOKEN_KEY.") from exc
        self.app = msal.PublicClientApplication(
            config.client_id,
            authority="https://login.microsoftonline.com/" + config.tenant_id,
            token_cache=self.cache,
        )
        self.username = config.account_username

    def save(self):
        if self.cache.has_state_changed:
            temporary = self.path.with_suffix(".tmp")
            temporary.write_bytes(self.cipher.encrypt(self.cache.serialize().encode()))
            if os.name != "nt":
                temporary.chmod(0o600)
            temporary.replace(self.path)

    def account(self):
        accounts = self.app.get_accounts(username=self.username or None)
        if len(accounts) != 1:
            raise SyncError(
                "Run auth first. Set TSYNC_ACCOUNT_USERNAME if multiple accounts are cached."
            )
        return accounts[0]

    def token(self):
        try:
            result = self.app.acquire_token_silent_with_error(SCOPES, account=self.account())
        finally:
            self.save()
        if not result or "access_token" not in result:
            raise SyncError(
                "Calendar sign-in expired or consent is missing. Run auth interactively."
            )
        return result["access_token"]

    def login(self):
        flow = self.app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise SyncError("Device sign-in could not start. Enable public-client flows in Entra.")
        print(flow["message"], flush=True)
        try:
            result = self.app.acquire_token_by_device_flow(flow)
        finally:
            self.save()
        if "access_token" not in result:
            raise SyncError("Sign-in failed. Your school may need to approve Calendars.ReadWrite.")
        return self.account().get("username", "")
