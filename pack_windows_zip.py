"""Pack dist/ into a Windows zip with correct UTF-8 Cyrillic filenames.

PowerShell Compress-Archive historically stores non-ASCII names in a way that
turns «ИНСТРУКЦИЯ» into mojibake (╨Ш╨Э╨б…). This helper writes ZipInfo with the
UTF-8 flag so Explorer / 7-Zip / WinRAR show the real names.
Also ships ASCII INSTRUCTION.* copies for tools that ignore the UTF-8 flag.
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path


def pack(dist: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    files = sorted(
        p for p in dist.rglob("*") if p.is_file() and p.name != zip_path.name
    )
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in files:
            arcname = path.relative_to(dist).as_posix()
            # Drop accidental mojibake duplicates if present
            if "╨" in arcname or "┬" in arcname:
                print(f"skip mojibake path: {arcname}")
                continue
            info = zipfile.ZipInfo(arcname)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.flag_bits |= 0x800  # UTF-8 filenames
            # Preserve mtime roughly
            st = path.stat()
            import time as _time

            info.date_time = _time.localtime(st.st_mtime)[:6]
            with path.open("rb") as src:
                zf.writestr(info, src.read())
            print(f"+ {arcname}")
    print(f"Wrote {zip_path} ({zip_path.stat().st_size} bytes)")


def main() -> int:
    root = Path(__file__).resolve().parent
    version = (root / "VERSION").read_text(encoding="utf-8").strip()
    dist = root / "dist"
    if len(sys.argv) > 1:
        dist = Path(sys.argv[1])
    zip_path = root / f"media-monitor-v{version}-windows.zip"
    if len(sys.argv) > 2:
        zip_path = Path(sys.argv[2])
    if not dist.is_dir():
        print(f"ERROR: dist not found: {dist}", file=sys.stderr)
        return 1
    pack(dist, zip_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
