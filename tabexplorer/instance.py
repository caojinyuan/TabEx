"""Single-writer ownership and bounded, acknowledged local instance requests."""

from collections import OrderedDict, deque
import hashlib
import json
import os
from pathlib import Path
import secrets
import socket
import struct
import tempfile
import threading
import time

from PyQt5.QtCore import QLockFile


MAX_MESSAGE_BYTES = 262144


def _receive_message(connection, timeout=1.0):
    deadline = time.monotonic() + timeout

    def receive_exact(length):
        chunks = bytearray()
        while len(chunks) < length:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutError('Instance request timed out')
            connection.settimeout(remaining)
            chunk = connection.recv(length - len(chunks))
            if not chunk:
                raise ConnectionError('Incomplete instance request')
            chunks.extend(chunk)
        return bytes(chunks)

    length = struct.unpack('!I', receive_exact(4))[0]
    if length <= 0 or length > MAX_MESSAGE_BYTES:
        raise ValueError('Invalid instance message size')
    message = json.loads(receive_exact(length).decode('utf-8'))
    if not isinstance(message, dict):
        raise ValueError('Invalid instance message')
    return message


def _send_message(connection, message):
    encoded = json.dumps(message, ensure_ascii=True).encode('utf-8')
    if len(encoded) > MAX_MESSAGE_BYTES:
        raise ValueError('Instance message is too large')
    connection.sendall(struct.pack('!I', len(encoded)) + encoded)


class InstanceCoordinator:
    def __init__(self, app_directory, state_directory=None):
        identity = os.path.normcase(os.path.realpath(app_directory))
        key = hashlib.sha256(identity.encode('utf-8')).hexdigest()[:24]
        root = Path(state_directory or os.path.join(
            os.environ.get('LOCALAPPDATA') or tempfile.gettempdir(), 'TabEx', 'ipc'))
        root.mkdir(parents=True, exist_ok=True)
        self.endpoint = root / (key + '.json')
        self.lock = QLockFile(str(root / (key + '.lock')))
        self.lock.setStaleLockTime(0)
        self.is_owner = False
        self._server = None
        self._connection = None
        self._thread = None
        self._stop = threading.Event()
        self._guard = threading.Lock()
        self._receiver = None
        self._pending = deque()
        self._seen = OrderedDict()
        self._token = secrets.token_hex(32)

    def acquire(self):
        if self.is_owner:
            return True
        if not self.lock.tryLock(0):
            if self.lock.error() != QLockFile.LockFailedError:
                raise OSError('Cannot acquire the application data lock')
            return False
        self.is_owner = True
        try:
            self._server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self._server.bind(('127.0.0.1', 0))
            self._server.listen(8)
            self._server.settimeout(0.2)
            temporary = self.endpoint.with_suffix('.tmp')
            temporary.write_text(json.dumps({
                'port': self._server.getsockname()[1], 'token': self._token,
            }), encoding='utf-8')
            os.replace(temporary, self.endpoint)
            self._thread = threading.Thread(target=self._serve, name='InstanceServer', daemon=True)
            self._thread.start()
            return True
        except Exception:
            self.close()
            raise

    def set_receiver(self, receiver):
        with self._guard:
            self._receiver = receiver
            while self._pending:
                receiver(self._pending.popleft())

    def _accept_request(self, message):
        token = message.get('token')
        request_id = message.get('id')
        path = message.get('path')
        if (message.get('version') != 1 or not isinstance(token, str)
                or not secrets.compare_digest(token, self._token)
                or not isinstance(request_id, str) or len(request_id) != 32
                or not isinstance(path, str) or len(path) > 32768 or '\x00' in path):
            raise ValueError('Invalid instance request')
        with self._guard:
            if request_id in self._seen:
                return
            if self._receiver is None:
                if len(self._pending) >= 128:
                    raise ValueError('Too many pending instance requests')
                self._pending.append(path)
            else:
                self._receiver(path)
            self._seen[request_id] = None
            while len(self._seen) > 256:
                self._seen.popitem(last=False)

    def _serve(self):
        while not self._stop.is_set():
            try:
                connection, _address = self._server.accept()
            except socket.timeout:
                continue
            except OSError:
                break
            self._connection = connection
            try:
                with connection:
                    message = _receive_message(connection, timeout=0.5)
                    self._accept_request(message)
                    _send_message(connection, {'id': message['id'], 'accepted': True})
            except (OSError, ValueError, ConnectionError):
                pass
            finally:
                self._connection = None

    def send(self, path='', timeout=5.0, request_id=None):
        request_id = request_id or secrets.token_hex(16)
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            try:
                with self.endpoint.open('r', encoding='utf-8') as source:
                    endpoint = json.loads(source.read(4096))
                port = int(endpoint['port'])
                remaining = max(0.01, min(1.0, deadline - time.monotonic()))
                with socket.create_connection(('127.0.0.1', port), timeout=remaining) as connection:
                    _send_message(connection, {
                        'version': 1, 'id': request_id, 'token': endpoint['token'], 'path': path or '',
                    })
                    response = _receive_message(connection, timeout=remaining)
                if response.get('id') == request_id and response.get('accepted') is True:
                    return True
            except (OSError, ValueError, KeyError, TypeError):
                pass
            self._stop.wait(min(0.05, max(0, deadline - time.monotonic())))
        return False

    def close(self):
        if not self.is_owner:
            return
        self._stop.set()
        for connection in (self._connection, self._server):
            if connection is not None:
                try:
                    connection.close()
                except OSError:
                    pass
        if self._thread is not None:
            self._thread.join(timeout=1.0)
        try:
            self.endpoint.unlink(missing_ok=True)
        finally:
            self.lock.unlock()
            self.is_owner = False