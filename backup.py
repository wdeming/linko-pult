import sqlite3, datetime, os, glob, time
os.chdir('/opt/linko-pult/data'); os.makedirs('backups', exist_ok=True)
s = sqlite3.connect('pult.db'); d = sqlite3.connect('backups/pult-%s.db' % datetime.date.today()); s.backup(d); d.close(); s.close()
for f in glob.glob('backups/pult-*.db'):
    if time.time() - os.path.getmtime(f) > 14 * 86400: os.remove(f)
