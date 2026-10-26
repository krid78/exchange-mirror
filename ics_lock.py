"""Nonblocking OS lock per importer profile and concrete destination."""

from contextlib import contextmanager
import errno
from hashlib import sha256
import json
import os
from pathlib import Path

from ics_profiles import default_profiles_path


class RunBusyError(RuntimeError):
    pass


class RunLockError(RuntimeError):
    pass


@contextmanager
def target_lock(profile, *, directory=None):
    root = Path(directory) if directory is not None else default_profiles_path().parent / "locks"
    identity = [str(profile.profile_id), profile.destination.store_id.casefold(),
                profile.destination.folder_id.casefold()]
    name = sha256(json.dumps(identity, separators=(",", ":")).encode()).hexdigest() + ".lock"
    try:
        root.mkdir(parents=True, exist_ok=True)
        stream = (root / name).open("a+b")
    except OSError:
        raise RunLockError("Cannot open the importer run lock") from None
    locked = False
    try:
        try:
            stream.seek(0)
            if os.name == "nt":
                import msvcrt
                # The byte region may extend beyond EOF; no lock-file payload is needed.
                msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            locked = True
        except OSError as error:
            if error.errno in (errno.EACCES, errno.EAGAIN, errno.EDEADLK):
                raise RunBusyError("Another importer run owns this profile/destination") from None
            raise RunLockError("Cannot acquire the importer run lock") from None
        yield
    finally:
        try:
            if locked:
                stream.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(stream.fileno(), fcntl.LOCK_UN)
        finally:
            stream.close()
        # Keep the file: unlinking a locked inode permits split-lock races on Unix.
        # OS handles release locks on process exit; an empty file is not a stale lock.
