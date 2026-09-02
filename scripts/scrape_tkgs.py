#!/usr/bin/env python3
"""
TKGS Türksat Kanal Listesi Scraper
===================================
Türksat 42.0°E uydusundaki TV kanallarını Türksat resmi Excel API'sinden
çekerek `tkgs_turksat.json` dosyasını üretir.

Veri kaynağı:
  https://www.turksat.com.tr/turksat-api/download/xlsx/frequency-list

Excel kolonları:
  Sıra | Kanal Adı | Frekans | Polarizasyon | Kapsama | SR | FEC |
  V-PID | A-PID | Uydu | Format | Paket Adı | Şifre Durumu

Kullanım:
    python scrape_tkgs.py [--output PATH]
"""

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

import openpyxl
import requests

# ---------------------------------------------------------------------------
# Türksat resmi Excel API
# ---------------------------------------------------------------------------
TURKSAT_EXCEL_URL = (
    "https://www.turksat.com.tr/turksat-api/download/xlsx/frequency-list"
)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
}

# ---------------------------------------------------------------------------
# Resmi TKGS LCN sıralaması (RTÜK / Türksat yayın sırası)
# Kanal adı (UPPER) → LCN eşleştirmesi
# ---------------------------------------------------------------------------
TKGS_LCN_MAP: dict[str, int] = {
    "TRT 1 HD": 1,
    "TRT 1": 1,
    "TRT 2 HD": 2,
    "TRT 2": 2,
    "TRT WORLD HD": 3,
    "TRT WORLD": 3,
    "TRT HABER HD": 4,
    "TRT HABER": 4,
    "TRT BELGESEL HD": 5,
    "TRT BELGESEL": 5,
    "TRT SPOR HD": 6,
    "TRT SPOR": 6,
    "TRT SPOR YILDIZ HD": 7,
    "TRT SPOR YILDIZ": 7,
    "TRT SPOR 2 HD": 7,
    "TRT MÜZIK HD": 8,
    "TRT MUZIK HD": 8,
    "TRT MÜZIK": 8,
    "TRT ÇOCUK HD": 9,
    "TRT COCUK HD": 9,
    "TRT ÇOCUK": 9,
    "TRT KURDÎ HD": 10,
    "TRT KURDI HD": 10,
    "TRT KURDÎ": 10,
    "TRT TÜRK HD": 11,
    "TRT TURK HD": 11,
    "TRT TÜRK": 11,
    "TRT ARABI HD": 12,
    "TRT ARABI": 12,
    "TBMM TV HD": 14,
    "TBM TV HD": 14,
    "TBMM TV": 14,
    "TRT AVAZ HD": 13,
    "TRT AVAZ": 13,
    "DİYANET TV HD": 15,
    "DIYANET TV HD": 15,
    "DİYANET TV": 15,
    "DIYANET TV": 15,
    "ATV HD": 16,
    "ATV": 16,
    "KANAL D HD": 17,
    "KANAL D": 17,
    "STAR TV HD": 18,
    "STAR TV": 18,
    "SHOW TV HD": 19,
    "SHOW TV": 19,
    "FOX TV HD": 20,
    "FOX TV": 20,
    "FOX": 20,
    "NOW TV HD": 20,
    "NOW HD": 20,
    "NOW TV": 20,
    "TV8 HD": 21,
    "TV8": 21,
    "KANAL 7 HD": 22,
    "KANAL 7": 22,
    "360 HD": 23,
    "360": 23,
    "HABERTÜRK TV HD": 24,
    "HABERTURK TV HD": 24,
    "HABERTÜRK TV": 24,
    "BLOOMBERG HT HD": 25,
    "BLOOMBERG HT": 25,
    "NTV HD": 26,
    "NTV": 26,
    "CNN TÜRK HD": 27,
    "CNN TURK HD": 27,
    "CNN TÜRK": 27,
    "A HABER HD": 28,
    "A HABER": 28,
    "TGRT HABER HD": 29,
    "TGRT HABER": 29,
    "HABER GLOBAL HD": 30,
    "HABER GLOBAL": 30,
    "TV 100 HD": 31,
    "TV100 HD": 31,
    "TV100": 31,
    "TV 100": 31,
    "BENGÜ TÜRK HD": 32,
    "BENGU TURK HD": 32,
    "BENGÜ TÜRK": 32,
    "ÜLKE TV HD": 33,
    "ULKE TV HD": 33,
    "ÜLKE TV": 33,
    "TELE1 HD": 34,
    "TELE1": 34,
    "KRT HD": 35,
    "KRT": 35,
    "FLASH HABER HD": 36,
    "FLASH HABER": 36,
    "FLASH TV": 36,
    "TVNET HD": 37,
    "TV NET HD": 37,
    "TVNET": 37,
    "BEYAZ TV HD": 38,
    "BEYAZ TV": 38,
    "A2 TV HD": 39,
    "A2 HD": 39,
    "A2 TV": 39,
    "TV 360 HD": 40,
    "TV360 HD": 40,
    "TV360": 40,
    "EKOL TV HD": 41,
    "EKOL TV": 41,
    "TRT EBA TV İLKOKUL": 42,
    "TRT EBA TV ILKOKUL": 42,
    "TRT EBA TV ORTAOKUL": 43,
    "TRT EBA TV LİSE": 44,
    "TRT EBA TV LISE": 44,
    "CARTOON NETWORK": 45,
    "MİNİKA ÇOCUK": 46,
    "MINIKA COCUK": 46,
    "MİNİKA GO": 47,
    "MINIKA GO": 47,
    "PLANET ÇOCUK": 48,
    "PLANET COCUK": 48,
    "BABY TV": 49,
    "TLC HD": 50,
    "TLC": 50,
    "TV8,5 HD": 51,
    "TV8.5 HD": 51,
    "TV8,5": 51,
    "TV8.5": 51,
    "TV 8,5 HD": 51,
    "NUMBER1 TV HD": 52,
    "NUMBER1 TV": 52,
    "POWER TÜRK TV HD": 53,
    "POWER TURK TV HD": 53,
    "POWER TÜRK TV": 53,
    "KRAL POP TV": 54,
    "KRAL TV": 55,
    "DREAM TÜRK HD": 56,
    "DREAM TURK HD": 56,
    "DREAM TÜRK": 56,
    "DREAM TURK": 56,
    "DREAM TV HD": 57,
    "DREAM TV": 57,
    "TRT DİYANET ÇOCUK": 58,
    "TRT DIYANET COCUK": 58,
    "TRT GENÇ": 59,
    "TRT GENC": 59,
    "24 TV": 60,
    "24 HD": 60,
}


# ---------------------------------------------------------------------------
# Name normalization & LCN matching
# ---------------------------------------------------------------------------
def _normalize_ascii(name: str) -> str:
    """Normalize Turkish characters to ASCII for fuzzy matching."""
    n = name.strip().upper()
    n = re.sub(r"\s+", " ", n)
    tr_map = {
        "İ": "I", "Ş": "S", "Ğ": "G", "Ü": "U", "Ö": "O", "Ç": "C",
        "ı": "I", "ş": "S", "ğ": "G", "ü": "U", "ö": "O", "ç": "C",
        "Î": "I", "î": "I",
    }
    for tr_char, en_char in tr_map.items():
        n = n.replace(tr_char, en_char)
    return n


def _try_match_lcn(name: str) -> int | None:
    """Try to find an LCN for the given channel name."""
    upper = name.strip().upper()

    # Direct match
    if upper in TKGS_LCN_MAP:
        return TKGS_LCN_MAP[upper]

    # Try with/without HD suffix
    if upper.endswith(" HD"):
        without = upper[:-3].strip()
        if without in TKGS_LCN_MAP:
            return TKGS_LCN_MAP[without]
    else:
        with_hd = upper + " HD"
        if with_hd in TKGS_LCN_MAP:
            return TKGS_LCN_MAP[with_hd]

    # ASCII normalized match
    n_ascii = _normalize_ascii(name)
    for map_name, lcn in TKGS_LCN_MAP.items():
        if _normalize_ascii(map_name) == n_ascii:
            return lcn

    return None


# ---------------------------------------------------------------------------
# Türksat Excel scraper
# ---------------------------------------------------------------------------
def scrape_turksat() -> list[dict]:
    """Download and parse the official Turksat frequency Excel file.

    Excel columns (row 0 = header):
      0: Sıra
      1: Kanal Adı
      2: Frekans
      3: Polarizasyon  (e.g. "V - Dikey", "H - Yatay")
      4: Kapsama
      5: SR
      6: FEC
      7: V-PID         (Video PID — None for radio)
      8: A-PID
      9: Uydu           (T3A, T4A, T5B, T6A)
     10: Format         (HD, SD, RD, UHD)
     11: Paket Adı
     12: Şifre Durumu   ("Şifresiz" or "Şifreli")
    """
    print("  Türksat Excel indiriliyor...")
    resp = requests.get(TURKSAT_EXCEL_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    print(f"  İndirildi: {len(resp.content) / 1024:.0f} KB")

    wb = openpyxl.load_workbook(BytesIO(resp.content), read_only=True)
    ws = wb.active
    if ws is None:
        print("  HATA: Excel dosyasında aktif sayfa bulunamadı.")
        return []

    channels: list[dict] = []
    header_found = False

    for row in ws.iter_rows(values_only=True):
        # Skip header row
        if not header_found:
            if row and str(row[0]).strip().lower() in ("sıra", "sira", "no"):
                header_found = True
            continue

        if not row or row[0] is None:
            continue

        try:
            name = str(row[1]).strip() if row[1] else ""
            if not name:
                continue

            freq = int(row[2]) if row[2] else 0
            pol_raw = str(row[3]).strip() if row[3] else ""
            pol = pol_raw[0] if pol_raw and pol_raw[0] in ("V", "H") else ""
            coverage = str(row[4]).strip() if row[4] else ""
            sr = int(row[5]) if row[5] else 0
            vpid = row[7]  # None for radio
            satellite = str(row[9]).strip() if row[9] else ""
            fmt = str(row[10]).strip() if row[10] else ""
            encryption = str(row[12]).strip() if row[12] else ""

            # Only FTA (Şifresiz) TV channels (skip radio: V-PID is None)
            is_fta = "şifresiz" in encryption.lower() or "sifresiz" in encryption.lower()
            is_tv = vpid is not None and str(vpid).strip() != ""
            is_radio = fmt.upper() == "RD"

            if not is_fta:
                continue
            if is_radio or not is_tv:
                continue

            channels.append({
                "name": name,
                "sid": 0,  # Türksat Excel'de SID yok
                "tsid": 0,
                "onid": 1070,
                "freq": freq,
                "pol": pol,
                "sr": sr,
                "satellite": satellite,
                "format": fmt,
                "coverage": coverage,
            })
        except (ValueError, IndexError, TypeError):
            continue

    return channels


# ---------------------------------------------------------------------------
# Build JSON
# ---------------------------------------------------------------------------
def build_tkgs_json(channels: list[dict]) -> dict:
    """Assign LCN numbers and build the final JSON structure."""
    lcn_channels = []
    unmatched_channels = []
    seen_lcns: set[int] = set()

    for ch in channels:
        lcn = _try_match_lcn(ch["name"])
        if lcn is not None and lcn not in seen_lcns:
            seen_lcns.add(lcn)
            ch["lcn"] = lcn
            lcn_channels.append(ch)
        else:
            unmatched_channels.append(ch)

    # Sort matched by LCN
    lcn_channels.sort(key=lambda c: c["lcn"])

    # Assign offset LCN (1000+) to remaining
    offset = 1000
    for ch in unmatched_channels:
        while offset in seen_lcns:
            offset += 1
        ch["lcn"] = offset
        seen_lcns.add(offset)
        lcn_channels.append(ch)
        offset += 1

    now = datetime.now(timezone.utc)
    return {
        "version": now.strftime("%Y.%m.%d"),
        "satellite": "Turksat 42.0E",
        "description": "TKGS Türksat LCN Sıralama Şablonu – Türksat resmi verisi",
        "updated_at": now.isoformat(),
        "source": "turksat.com.tr",
        "channel_count": len(lcn_channels),
        "channels": [
            {
                "lcn": ch["lcn"],
                "name": ch["name"],
                "sid": ch.get("sid", 0),
                "tsid": ch.get("tsid", 0),
                "onid": ch.get("onid", 1070),
                "freq": ch.get("freq", 0),
                "pol": ch.get("pol", ""),
            }
            for ch in lcn_channels
        ],
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(
        description="TKGS Türksat kanal listesi scraper (Türksat resmi Excel)"
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Çıktı dosya yolu (varsayılan: repo kökünde tkgs_turksat.json)",
    )
    args = parser.parse_args()

    if args.output:
        output_path = Path(args.output)
    else:
        script_dir = Path(__file__).resolve().parent
        repo_root = script_dir.parent
        output_path = repo_root / "tkgs_turksat.json"

    print("=" * 60)
    print("TKGS Türksat Kanal Listesi Scraper")
    print("Kaynak: turksat.com.tr (Resmi Excel API)")
    print("=" * 60)

    # Step 1: Download & parse
    print("\n[1/2] Türksat resmi frekans listesi çekiliyor...")
    channels = scrape_turksat()
    print(f"  → {len(channels)} FTA TV kanalı bulundu")

    if not channels:
        print("\nHATA: Hiç kanal verisi çekilemedi!")
        sys.exit(1)

    # Step 2: Build JSON
    print("\n[2/2] JSON dosyası oluşturuluyor...")
    result = build_tkgs_json(channels)

    matched = sum(1 for c in result["channels"] if c["lcn"] < 1000)
    unmatched = result["channel_count"] - matched
    print(f"  → {matched} kanal TKGS LCN eşleştirildi")
    print(f"  → {unmatched} kanal offset bölgesinde (1000+)")

    # Write
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Dosya yazıldı: {output_path}")
    print(f"   Toplam: {result['channel_count']} kanal")
    print(f"   Sürüm: {result['version']}")


if __name__ == "__main__":
    main()
