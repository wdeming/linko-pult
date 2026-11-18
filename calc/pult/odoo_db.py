# -*- coding: utf-8 -*-
"""OdooDB для пульта: только чтение. SSH-ключ linko_ro умеет лишь открыть туннель к PostgreSQL; роль БД linko_ro — только SELECT."""
import os, sys, json, select, threading, socketserver, atexit
import paramiko, psycopg2
from psycopg2 import extras
_D = os.path.dirname(os.path.abspath(__file__))
_A = json.load(open(os.path.join(_D, 'access.json'), encoding='utf-8'))
_cli = paramiko.SSHClient(); _cli.set_missing_host_key_policy(paramiko.AutoAddPolicy())
_cli.connect(_A['ssh_host'], username=_A['ssh_user'], pkey=paramiko.Ed25519Key.from_private_key_file(os.path.join(_D, _A['ssh_key'])),
             timeout=20, look_for_keys=False, allow_agent=False)
_tr = _cli.get_transport(); _tr.set_keepalive(20)
class _H(socketserver.BaseRequestHandler):
    def handle(self):
        try: ch = _tr.open_channel('direct-tcpip', ('127.0.0.1', 5432), self.request.getpeername())
        except Exception as e: sys.stderr.write('tunnel open failed: %r\n' % e); return
        try:
            while True:
                r, _, _ = select.select([self.request, ch], [], [], 60)
                if self.request in r:
                    d = self.request.recv(65536)
                    if not d: break
                    ch.sendall(d)
                if ch in r:
                    d = ch.recv(65536)
                    if not d: break
                    self.request.sendall(d)
        finally:
            ch.close(); self.request.close()
class _S(socketserver.ThreadingTCPServer):
    daemon_threads = True; allow_reuse_address = True
_srv = _S(('127.0.0.1', 0), _H)
threading.Thread(target=_srv.serve_forever, daemon=True).start()
atexit.register(lambda: (_srv.shutdown(), _cli.close()))
DB_CONFIG = dict(host='127.0.0.1', port=_srv.server_address[1], dbname=_A['db_name'], user=_A['db_user'], password=_A['db_password'], connect_timeout=20)
class OdooDB:
    def __init__(self, config=None):
        self._conn = psycopg2.connect(**(config or DB_CONFIG)); self._conn.autocommit = True
    def query(self, sql_query, params=None):
        with self._conn.cursor(cursor_factory=extras.RealDictCursor) as cur:
            cur.execute(sql_query, params); return cur.fetchall()
    def query_one(self, sql_query, params=None):
        r = self.query(sql_query, params); return r[0] if r else None
    def close(self):
        try: self._conn.close()
        except Exception: pass
