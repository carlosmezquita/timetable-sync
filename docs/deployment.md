# Micro deployment

Python + SQLite worker; no Docker or resident web server. The timer runs every five minutes. The included 192 MiB memory ceiling is provisional; measure live runs before relying on it. This implementation has not been deployed to Micro.

Complete Entra authentication and a reviewed live preview first. Ubuntu 22.04 Python 3.10 is supported.

~~~sh
sudo apt-get update
sudo apt-get install -y python3-venv
sudo useradd --system --create-home --home-dir /var/lib/timetable-sync --shell /usr/sbin/nologin timetable-sync
sudo install -d -o timetable-sync -g timetable-sync -m 700 /var/lib/timetable-sync/state
sudo install -d -m 755 /opt/timetable-sync/releases
~~~

Transfer a Git archive of the reviewed commit into /opt/timetable-sync/releases/COMMIT. From that release:

~~~sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.lock
.venv/bin/python -m pip install --no-deps .
~~~

Create /etc/timetable-sync.env owned root:timetable-sync mode 640. Use TSYNC_STATE_DIR=/var/lib/timetable-sync/state and preserve the local TSYNC_TOKEN_KEY. Copy tokens.enc and a SQLite backup from local state, owned timetable-sync:timetable-sync mode 600. The database includes the account binding needed for --apply, rules and transaction records. Rename backup.sqlite3 to state.sqlite3 on the server.

Stop any existing local writer before handover. Never run independent writers against the same source. Keep credentials and feed URLs out of Git/build artifacts.

Point /opt/timetable-sync/current at the new release. Install units:

~~~sh
sudo install -m 644 deploy/timetable-sync.service /etc/systemd/system/
sudo install -m 644 deploy/timetable-sync.timer /etc/systemd/system/
sudo systemctl daemon-reload
sudo -u timetable-sync /opt/timetable-sync/current/.venv/bin/timetable-sync --env /etc/timetable-sync.env sync --outlook
~~~

Review preview before starting a write, then inspect the run and enable scheduling:

~~~sh
sudo systemctl start timetable-sync.service
sudo journalctl -u timetable-sync.service --since '10 minutes ago'
sudo systemctl enable --now timetable-sync.timer
systemctl list-timers timetable-sync.timer
sudo -u timetable-sync /opt/timetable-sync/current/.venv/bin/timetable-sync --env /etc/timetable-sync.env health
~~~

The manual preview runs outside the service memory ceiling. Configure an external dead-man monitor with TSYNC_HEARTBEAT_URL; check health for held changes.

## Releases

Stop timer, back up state, upload an immutable commit release, install dependencies, run tests/preview, switch current, restart timer. Retain prior release and backup. Roll back code by restoring the previous symlink with the timer stopped. Restoring SQLite alone does not undo Outlook writes.

GitHub Actions validates commits; release transfer is initially manual. Add SSH deployment after account setup and server resource measurements are verified.
