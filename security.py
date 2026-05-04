"""API key authentication."""

import os
import secrets

from fastapi import Header, HTTPException, status


def get_expected_api_key() -> str:
    key = os.environ.get("API_KEY")
    if not key:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Service not configured (API_KEY missing)",
        )
    return key


async def require_api_key(x_api_key: str | None = Header(default=None)):
    expected = get_expected_api_key()
    if x_api_key is None or not secrets.compare_digest(x_api_key, expected):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or missing API key",
        )
