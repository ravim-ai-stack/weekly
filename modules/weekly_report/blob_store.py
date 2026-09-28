"""Thin wrapper over Vercel Blob (private store) - the durable, cross-instance
replacement for everything that used to live under the ephemeral, per-instance
/tmp on Vercel (see WRITABLE_DIR in app.py, updates_store.py and
medtronic_timesheet.py: Vercel's deployment filesystem is read-only outside of
/tmp, and /tmp itself is wiped between deploys and isn't shared across
serverless instances, so anything written there could vanish, or be invisible
to the very next request, without warning).

Access is private (not public) since these are internal client-engagement
documents - reads and writes both require BLOB_READ_WRITE_TOKEN rather than
being reachable by anyone who finds/guesses the URL."""

from vercel.blob import BlobClient
from vercel.blob.errors import BlobNotFoundError


def put(pathname: str, data: bytes, content_type: str = "application/octet-stream") -> str:
    """Uploads to `pathname`, overwriting any existing object there. Returns
    the blob's URL, needed for later get()/delete() calls."""
    result = BlobClient().put(pathname, data, access="private", content_type=content_type, overwrite=True)
    return result.url


def get(url: str) -> bytes | None:
    """Fetches a blob's bytes by its URL (as returned by put() or
    list_objects()). Returns None if the object doesn't exist.

    use_cache=False: every pathname we read is either a JSON store
    (updates_store.py, medtronic_timesheet.py) overwritten on nearly every
    request, or a single-pathname generated file (e.g.
    Medtronic_Time_Sheet.xlsx) overwritten on every save - Vercel's private-
    blob CDN cache can serve a stale/previous version for up to 60s after an
    overwrite otherwise, which silently drops the just-written change on the
    very next read."""
    try:
        return BlobClient().get(url, access="private", use_cache=False).content
    except BlobNotFoundError:
        return None


def delete(url: str) -> None:
    BlobClient().delete(url)


def list_objects(prefix: str = None) -> list:
    """Blob items (.pathname, .url, .uploaded_at, ...) under `prefix`."""
    return BlobClient().list_objects(prefix=prefix).blobs
