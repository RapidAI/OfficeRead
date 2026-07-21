#!/usr/bin/env python3
import argparse
import concurrent.futures
import csv
import hashlib
import os
import re
import shutil
import ssl
import struct
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
import zlib
from pathlib import Path


EXTS = (".doc", ".docx", ".ppt", ".pptx", ".xls", ".xlsx")

MSX_LIST_URL = "https://roussev.net/msx-13/msx-13--name-hash-size-url.txt"
GOVDOCS_BY_TYPE = {
    ".doc": ["https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/doc.zip"],
    ".ppt": [
        "https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/ppt-part1.zip",
        "https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/ppt-part2.zip",
        "https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/ppt-part3.zip",
        "https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/ppt-part4.zip",
    ],
    ".xls": ["https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/xls.zip"],
    ".docx": ["https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/docx.zip"],
    ".pptx": ["https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/pptx.zip"],
    ".xlsx": ["https://digitalcorpora.s3.amazonaws.com/corpora/files/govdocs1/by_type/xlsx.zip"],
}


def ssl_context():
    ctx = ssl.create_default_context()
    if os.environ.get("OFFICEREAD_INSECURE_TLS") == "1":
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
    return ctx


CTX = ssl_context()


def request(url, headers=None, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": "officeread-compat-test", **(headers or {})})
    return urllib.request.urlopen(req, timeout=timeout, context=CTX)


def read_range(url: str, start: int, size: int, timeout: int = 60) -> bytes:
    if size <= 0:
        return b""
    with request(url, headers={"Range": f"bytes={start}-{start + size - 1}"}, timeout=timeout) as r:
        return r.read()


def ensure_msx_list(root: Path) -> Path:
    path = root / "sources" / "msx-13--name-hash-size-url.txt"
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists() and path.stat().st_size > 1000:
        return path
    with request(MSX_LIST_URL, timeout=60) as r, path.open("wb") as f:
        shutil.copyfileobj(r, f)
    return path


def parse_msx_list(path: Path, ext: str, max_size: int):
    pat = re.compile(r'^(?P<name>\S+%s)\s+(?P<sha1>[0-9a-f]{40})\s+(?P<size>\d+)\s+"(?P<url>.*)"$' % re.escape(ext))
    out = []
    for line in path.read_text(errors="replace").splitlines():
        m = pat.match(line.strip())
        if not m:
            continue
        size = int(m.group("size"))
        if size <= max_size:
            out.append((m.group("name"), m.group("sha1"), size, m.group("url")))
    return out


def valid_ext(path: Path, ext: str) -> bool:
    if not path.exists() or path.stat().st_size == 0:
        return False
    try:
        head = path.read_bytes()[:8]
    except OSError:
        return False
    if ext in (".docx", ".pptx", ".xlsx"):
        return head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06") or head.startswith(b"PK\x07\x08")
    if ext in (".doc", ".ppt", ".xls"):
        return head.startswith(bytes.fromhex("d0cf11e0a1b11ae1"))
    return True


def download_one_msx(item, ext_dir: Path, ext: str, timeout: int, max_size: int):
    name, sha1, size, url = item
    dest = ext_dir / name
    if valid_ext(dest, ext):
        return ("ok-existing", name, "")
    tmp = dest.with_suffix(dest.suffix + ".part")
    try:
        with request(url, timeout=timeout) as r:
            clen = r.headers.get("Content-Length")
            if clen and int(clen) > max_size:
                return ("skip-large", name, clen)
            h = hashlib.sha1()
            total = 0
            with tmp.open("wb") as f:
                while True:
                    chunk = r.read(1024 * 256)
                    if not chunk:
                        break
                    total += len(chunk)
                    if total > max_size:
                        return ("skip-large", name, str(total))
                    h.update(chunk)
                    f.write(chunk)
        if sha1 and h.hexdigest().lower() != sha1.lower():
            tmp.unlink(missing_ok=True)
            return ("bad-sha1", name, "")
        if not valid_ext(tmp, ext):
            tmp.unlink(missing_ok=True)
            return ("bad-format", name, "")
        tmp.replace(dest)
        return ("ok", name, "")
    except Exception as e:
        tmp.unlink(missing_ok=True)
        return ("error", name, str(e)[:180])


class RemoteHTTPFile:
    def __init__(self, url, timeout=60):
        self.url = url
        self.timeout = timeout
        req = urllib.request.Request(url, method="HEAD", headers={"User-Agent": "officeread-compat-test"})
        with urllib.request.urlopen(req, timeout=timeout, context=CTX) as r:
            self.size = int(r.headers["Content-Length"])
        self.pos = 0

    def seekable(self):
        return True

    def readable(self):
        return True

    def tell(self):
        return self.pos

    def seek(self, offset, whence=0):
        if whence == 0:
            self.pos = offset
        elif whence == 1:
            self.pos += offset
        elif whence == 2:
            self.pos = self.size + offset
        else:
            raise ValueError("invalid whence")
        return self.pos

    def read(self, n=-1):
        if n is None or n < 0:
            n = self.size - self.pos
        if n == 0 or self.pos >= self.size:
            return b""
        start = self.pos
        end = min(self.size - 1, start + n - 1)
        with request(self.url, headers={"Range": f"bytes={start}-{end}"}, timeout=self.timeout) as r:
            data = r.read()
        self.pos += len(data)
        return data

    def close(self):
        pass


def extract_from_remote_zip(url: str, ext_dir: Path, ext: str, target: int, max_size: int, log):
    made = 0
    existing = len([p for p in ext_dir.glob(f"*{ext}") if valid_ext(p, ext)])
    if existing >= target:
        return 0
    rf = RemoteHTTPFile(url)
    with zipfile.ZipFile(rf) as zf:
        infos = [i for i in zf.infolist() if not i.is_dir() and i.filename.lower().endswith(ext)]
        planned = []
        for info in infos:
            if existing + len(planned) >= target:
                break
            if info.file_size <= 0 or info.file_size > max_size:
                log.writerow([ext, "skip-large", info.filename, info.file_size, url])
                continue
            dest = destination_for_info(ext_dir, info, ext)
            if dest.exists() and valid_ext(dest, ext):
                continue
            planned.append((info, dest))
        with concurrent.futures.ThreadPoolExecutor(max_workers=8) as pool:
            futs = [pool.submit(extract_one_remote_zip_member, url, info, dest, ext) for info, dest in planned]
            for fut in concurrent.futures.as_completed(futs):
                status, name, detail, size = fut.result()
                if status == "ok":
                    made += 1
                log.writerow([ext, status, name, size if detail == "" else detail, url])
                if made % 50 == 0 or status != "ok":
                    print(f"{ext}: +{made} from current ZIP, total={count_samples(ext_dir.parent.parent, ext)}", flush=True)
    return made


def destination_for_info(ext_dir: Path, info: zipfile.ZipInfo, ext: str) -> Path:
    out_name = Path(info.filename).name
    if not out_name.lower().endswith(ext):
        out_name += ext
    dest = ext_dir / out_name
    if dest.exists():
        stem = dest.stem
        idx = 2
        while dest.exists():
            dest = ext_dir / f"{stem}-{idx}{ext}"
            idx += 1
    return dest


def local_data_offset(url: str, info: zipfile.ZipInfo) -> int:
    header = read_range(url, info.header_offset, 30)
    if len(header) != 30 or header[:4] != b"PK\x03\x04":
        raise ValueError("bad local file header")
    name_len, extra_len = struct.unpack_from("<HH", header, 26)
    return info.header_offset + 30 + name_len + extra_len


def decompress_zip_payload(info: zipfile.ZipInfo, payload: bytes) -> bytes:
    if info.compress_type == zipfile.ZIP_STORED:
        return payload
    if info.compress_type == zipfile.ZIP_DEFLATED:
        return zlib.decompress(payload, -15)
    raise NotImplementedError(f"zip compression method {info.compress_type} is not supported by fast extractor")


def extract_one_remote_zip_member(url: str, info: zipfile.ZipInfo, dest: Path, ext: str):
    tmp = dest.with_suffix(dest.suffix + ".part")
    try:
        data_offset = local_data_offset(url, info)
        payload = read_range(url, data_offset, info.compress_size)
        data = decompress_zip_payload(info, payload)
        if len(data) != info.file_size:
            return ("bad-size", info.filename, f"{len(data)} != {info.file_size}", info.file_size)
        tmp.write_bytes(data)
        if not valid_ext(tmp, ext):
            tmp.unlink(missing_ok=True)
            return ("bad-format", info.filename, "", info.file_size)
        tmp.replace(dest)
        return ("ok", info.filename, "", info.file_size)
    except Exception as e:
        tmp.unlink(missing_ok=True)
        return ("error", info.filename, str(e)[:180], info.file_size)


def count_samples(root: Path, ext: str) -> int:
    ext_dir = root / "samples" / ext[1:]
    return len([p for p in ext_dir.glob(f"*{ext}") if valid_ext(p, ext)])


def download_msx_batched(candidates, ext_dir: Path, root: Path, ext: str, target: int, timeout: int, max_size: int, workers: int, log, log_file):
    idx = 0
    batch_size = max(workers * 4, 16)
    while count_samples(root, ext) < target and idx < len(candidates):
        batch = candidates[idx:idx + batch_size]
        idx += batch_size
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
            futs = [pool.submit(download_one_msx, item, ext_dir, ext, timeout, max_size) for item in batch]
            for fut in concurrent.futures.as_completed(futs):
                status, name, detail = fut.result()
                log.writerow([ext, status, name, detail, "msx-13"])
                log_file.flush()
        print(f"{ext}: have {count_samples(root, ext)} after {idx} MSX candidates", flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default="testdata/web-samples")
    ap.add_argument("--target", type=int, default=1000)
    ap.add_argument("--max-size-mb", type=int, default=25)
    ap.add_argument("--timeout", type=int, default=20)
    ap.add_argument("--workers", type=int, default=16)
    ap.add_argument("--ext", action="append", choices=[e[1:] for e in EXTS])
    ap.add_argument("--prefer-msx", action="store_true", help="try MSX-13 direct URLs before GOVDOCS1 zip sampling for OOXML")
    args = ap.parse_args()

    root = Path(args.root)
    root.mkdir(parents=True, exist_ok=True)
    max_size = args.max_size_mb * 1024 * 1024
    exts = tuple("." + e for e in args.ext) if args.ext else EXTS
    log_path = root / "download-log.csv"

    with log_path.open("a", newline="", encoding="utf-8") as f:
        log = csv.writer(f)
        if f.tell() == 0:
            log.writerow(["ext", "status", "name", "size_or_detail", "source"])

        if args.prefer_msx:
            msx = ensure_msx_list(root)
            for ext in [e for e in exts if e in (".docx", ".pptx", ".xlsx")]:
                ext_dir = root / "samples" / ext[1:]
                ext_dir.mkdir(parents=True, exist_ok=True)
                need = args.target - count_samples(root, ext)
                if need <= 0:
                    continue
                candidates = parse_msx_list(msx, ext, max_size)
                print(f"{ext}: trying MSX-13 direct URLs; need={need}, candidates={len(candidates)}", flush=True)
                download_msx_batched(candidates, ext_dir, root, ext, args.target, args.timeout, max_size, args.workers, log, f)

        for ext in exts:
            ext_dir = root / "samples" / ext[1:]
            ext_dir.mkdir(parents=True, exist_ok=True)
            for url in GOVDOCS_BY_TYPE.get(ext, []):
                if count_samples(root, ext) >= args.target:
                    break
                print(f"{ext}: sampling GOVDOCS1 {url}", flush=True)
                extract_from_remote_zip(url, ext_dir, ext, args.target, max_size, log)
                f.flush()

    for ext in EXTS:
        print(f"{ext}: {count_samples(root, ext)}")


if __name__ == "__main__":
    main()
