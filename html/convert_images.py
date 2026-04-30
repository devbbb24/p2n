"""
images/ 폴더의 PNG/JPG를 WebP로 변환합니다.

  foo.png  ->  foo.webp      (원본 해상도, 데스크톱용)
  foo.png  ->  foo-sm.webp   (가로 800px 이하, 모바일용)

원본 이미지는 그대로 두고 WebP 파일만 추가합니다. 원본보다 WebP가 새 파일이면
재변환을 건너뜁니다 (idempotent).
"""

import sys
from pathlib import Path

try:
    from PIL import Image
except ImportError:
    print("Pillow가 설치되어 있지 않습니다. 다음 명령어로 설치하세요:")
    print("  py -3 -m pip install Pillow")
    sys.exit(1)

SCRIPT_DIR = Path(__file__).resolve().parent
IMAGES_DIR = SCRIPT_DIR / "images"

SM_MAX_WIDTH = 800
QUALITY_FULL = 85
QUALITY_SM = 78


def needs_update(src: Path, dst: Path) -> bool:
    if not dst.exists():
        return True
    return src.stat().st_mtime > dst.stat().st_mtime


def normalize_mode(im: "Image.Image") -> "Image.Image":
    if im.mode in ("RGB", "RGBA"):
        return im
    if im.mode == "P" and "transparency" in im.info:
        return im.convert("RGBA")
    if im.mode in ("LA", "PA"):
        return im.convert("RGBA")
    return im.convert("RGB")


def save_webp(im: "Image.Image", dst: Path, quality: int) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    normalize_mode(im).save(dst, "WEBP", quality=quality, method=6)


def convert_one(src: Path) -> int:
    base = src.with_suffix("")
    full_dst = base.with_suffix(".webp")
    sm_dst = base.parent / f"{base.name}-sm.webp"

    converted = 0

    if needs_update(src, full_dst):
        with Image.open(src) as im:
            save_webp(im, full_dst, QUALITY_FULL)
        converted += 1

    if needs_update(src, sm_dst):
        with Image.open(src) as im:
            w, h = im.size
            if w > SM_MAX_WIDTH:
                ratio = SM_MAX_WIDTH / w
                im = im.resize((SM_MAX_WIDTH, int(h * ratio)), Image.LANCZOS)
            save_webp(im, sm_dst, QUALITY_SM)
        converted += 1

    return converted


def main() -> None:
    if not IMAGES_DIR.exists():
        print(f"images 폴더를 찾을 수 없습니다: {IMAGES_DIR}")
        sys.exit(1)

    targets = []
    for pattern in ("*.png", "*.PNG", "*.jpg", "*.JPG", "*.jpeg", "*.JPEG"):
        targets.extend(IMAGES_DIR.rglob(pattern))

    targets = [p for p in targets if not p.name.endswith("-sm.webp")]

    print(f"원본 이미지 {len(targets)}개 검사 중...")

    new_files = 0
    skipped = 0
    failed = []

    for src in targets:
        try:
            n = convert_one(src)
            if n > 0:
                rel = src.relative_to(IMAGES_DIR)
                print(f"  + {rel}  ({n}개 생성/갱신)")
                new_files += n
            else:
                skipped += 1
        except Exception as e:
            failed.append((src, e))

    print()
    print(f"완료: {new_files}개 생성/갱신, {skipped}개 최신상태로 건너뜀")
    if failed:
        print(f"실패 {len(failed)}개:")
        for src, err in failed:
            print(f"  - {src.relative_to(IMAGES_DIR)}: {err}")
        sys.exit(1)


if __name__ == "__main__":
    main()
