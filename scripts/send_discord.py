# -*- coding: utf-8 -*-
"""
send_discord.py - Kirim file PNG (foto tabel harga) ke channel Discord via webhook.

Config webhook dibaca dari discord_config.json di folder repo:
{
  "harga-lm-emas":  {"webhook_url": "https://discord.com/api/webhooks/..."},
  "harga-lm-perak": {"webhook_url": "..."},
  ...
}

Pemakaian:
  send_discord.py <file.png> --channel harga-lm-emas [--caption "25 AGUSTUS 2026"]
  send_discord.py <file.png> --webhook https://discord.com/api/webhooks/... [--caption "..."]
"""
import argparse
import json
import mimetypes
import os
import uuid

try:
    import requests
    _HAVE_REQUESTS = True
except Exception:
    _HAVE_REQUESTS = False


def load_webhook(channel):
    cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "discord_config.json")
    if not os.path.exists(cfg_path):
        raise SystemExit("discord_config.json tidak ditemukan di %s" % cfg_path)
    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)
    entry = cfg.get(channel)
    if not entry:
        raise SystemExit("Channel '%s' tidak ada di discord_config.json. Ada: %s"
                         % (channel, ", ".join(cfg.keys())))
    url = entry.get("webhook_url")
    if not url:
        raise SystemExit("webhook_url kosong untuk channel '%s'" % channel)
    return url


def send_requests(url, png_path, caption):
    import requests as rq
    if caption:
        payload = {"content": caption}
        files = {"file": (os.path.basename(png_path), open(png_path, "rb"), "image/png")}
        resp = rq.post(url, data=payload, files=files)
    else:
        files = {"file": (os.path.basename(png_path), open(png_path, "rb"), "image/png")}
        resp = rq.post(url, files=files)
    return resp


def send_urllib(url, png_path, caption):
    """Kirim multipart/form-data tanpa dependency pihak ketiga."""
    boundary = "----WebhookFormBoundary" + uuid.uuid4().hex
    filename = os.path.basename(png_path)
    with open(png_path, "rb") as f:
        filedata = f.read()

    lines = []
    if caption:
        lines.append("--%s" % boundary)
        lines.append('Content-Disposition: form-data; name="content"')
        lines.append("")
        lines.append(caption)
    lines.append("--%s" % boundary)
    lines.append('Content-Disposition: form-data; name="file"; filename="%s"' % filename)
    lines.append("Content-Type: image/png")
    lines.append("")
    head = "\r\n".join(lines).encode("utf-8")
    tail = ("\r\n--%s--\r\n" % boundary).encode("utf-8")
    body = head + b"\r\n" + filedata + tail

    import urllib.request
    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": "multipart/form-data; boundary=%s" % boundary,
            "User-Agent": "Mozilla/5.0 (toko-mas-lia-price-bot)",
        },
        method="POST",
    )
    return urllib.request.urlopen(req, timeout=60)


def main():
    ap = argparse.ArgumentParser(description="Kirim PNG tabel harga ke Discord webhook.")
    ap.add_argument("file", help="Path ke file PNG.")
    ap.add_argument("--channel", default=None, help="Nama channel di discord_config.json.")
    ap.add_argument("--webhook", default=None, help="URL webhook langsung (ganti --channel).")
    ap.add_argument("--caption", default=None, help="Teks pesan sebelum gambar.")
    ap.add_argument("--config", default=None, help="Path manual ke discord_config.json.")
    args = ap.parse_args()

    if not os.path.exists(args.file):
        raise SystemExit("File tidak ditemukan: %s" % args.file)

    if args.webhook:
        url = args.webhook
    elif args.channel:
        # Izinkan override path config
        cfg_path = args.config or os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "..", "discord_config.json")
        if not os.path.exists(cfg_path):
            raise SystemExit("discord_config.json tidak ditemukan: %s" % cfg_path)
        with open(cfg_path, encoding="utf-8") as f:
            cfg = json.load(f)
        entry = cfg.get(args.channel)
        if not entry:
            raise SystemExit("Channel '%s' tidak ada. Ada: %s" % (args.channel, ", ".join(cfg)))
        url = entry.get("webhook_url")
        if not url:
            raise SystemExit("webhook_url kosong untuk channel '%s'" % args.channel)
    else:
        raise SystemExit("Butuh --channel atau --webhook.")

    if _HAVE_REQUESTS:
        resp = send_requests(url, args.file, args.caption)
        code = resp.status_code
        body = resp.text
    else:
        resp = send_urllib(url, args.file, args.caption)
        code = resp.status
        body = resp.read().decode("utf-8", "replace")
    ok = 200 <= code < 300
    print("[discord] status %s -> %s" % (code, "OK" if ok else "GAGAL"))
    if not ok:
        print("  resp:", body)
        raise SystemExit("Gagal mengirim ke Discord (status %s)." % code)


if __name__ == "__main__":
    main()
