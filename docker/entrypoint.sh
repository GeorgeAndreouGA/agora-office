#!/bin/sh
# Greek Office container start-up: turns GREEK_ALLOWED_IPS into the geo table that nginx.conf
# includes, then hands over to nginx. The root filesystem is read-only, so the table goes in the
# /tmp tmpfs and is rebuilt on every start.
#
# The value ends up inside nginx config, so every entry must match a strict pattern first. A stray
# ';', a hostname or a typo stops the container: that cannot inject config, and it cannot quietly
# let the wrong people in. A missing or empty list stops it too, so losing the setting never
# silently opens the app. Letting everyone in takes an explicit `all`.
set -eu

table=/tmp/allowed-ips.conf

fail() {
    echo "greek-office: $1" >&2
    echo "greek-office: set GREEK_ALLOWED_IPS in .env to comma-separated addresses or CIDR ranges, or to 'all' (see .env.example)." >&2
    exit 1
}

octet='(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9]?[0-9])'
ipv4="^$octet(\\.$octet){3}(/(3[0-2]|[12]?[0-9]))?\$"
# IPv6 is checked by shape only: these characters cannot break out of the table, and nginx itself
# refuses a malformed address when it starts (which stops the container the same way).
ipv6='^[0-9A-Fa-f:.]*:[0-9A-Fa-f:.]*(/(12[0-8]|1[01][0-9]|[1-9]?[0-9]))?$'

set -f # an entry such as * must stay literal, not expand to file names
entries=$(printf '%s' "${GREEK_ALLOWED_IPS:-}" | tr ',' ' ')

open=
count=0
for entry in $entries; do
    count=$((count + 1))
    if [ "$entry" = all ]; then
        open=yes
    elif ! printf '%s\n' "$entry" | grep -Eq "$ipv4|$ipv6"; then
        fail "'$entry' is not an IP address or CIDR range."
    fi
done

[ "$count" -gt 0 ] || fail 'GREEK_ALLOWED_IPS is missing or empty.'

if [ -n "$open" ]; then
    [ "$count" -eq 1 ] || fail "'all' cannot be combined with addresses."
    echo 'default 1;' > "$table"
    echo 'greek-office: IP allow-list off (GREEK_ALLOWED_IPS=all), every client may enter.'
else
    {
        echo 'default 0;'
        for entry in $entries; do
            echo "$entry 1;"
        done
    } > "$table"
    echo "greek-office: IP allow-list on, allowed:" $entries
fi

exec "$@"
