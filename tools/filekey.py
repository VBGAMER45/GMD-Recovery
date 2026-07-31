"""
Offline encryption-key recovery for GameMaker executables.

The VB6 tool finds the key by launching the game, dumping its memory and
scanning for a byte signature. That is why it breaks on anything but WinXP --
the signature offsets move. This does the job from the .exe on disk.

Why it is possible at all
-------------------------
GameMaker's protection is a monoalphabetic byte substitution: one fixed
256-byte permutation applied to the whole appended data stream. DecompileExe
in frmMain.frm builds the inverse with

    NewArray(Asc(Mid$(key, i + 1, 1))) = i
    plaintext = NewArray(ciphertext)

That leaks its own key, because the plaintext start is known: the data always
begins with the dword 1234321 (91 D5 12 00) followed by a version dword.

Header detection assumes nothing about the key. A substitution is injective,
so the equality pattern of those eight bytes alone identifies the version, and
E(0)=0 holds in every sample seen.

Key families (measured from the header deltas)
----------------------------------------------
  identity     data is not encrypted at all
  adjacent     every byte maps to itself +/-1; the whole keyspace is 8001
               tables, so the key is found by exhaustive search   (GM 5.0, 6.x)
  near-shift   a constant shift plus per-value jitter             (GM 5.1)
  full         an unstructured permutation; needs a real substitution
               solver, not implemented here                       (GM 5.3/5.3a)

Usage:  python filekey.py <game.exe> [...] [-o OUTDIR]
Requires numpy + scipy.
"""
import os
import sys
import glob
import struct
import argparse

import numpy as np
from scipy.optimize import linear_sum_assignment

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VERSIONS = {430: 0x1AE, 500: 0x1F4, 510: 0x1FE, 520: 0x208, 530: 0x212,
            540: 0x21C, 600: 0x258, 610: 0x262, 700: 0x2BC, 800: 0x320,
            810: 0x32A}
ZERO_AT = (3, 6, 7, 8, 9, 10, 11)
NONZERO_AT = (0, 1, 2, 4, 5)
PRIOR_BONUS = 1e7
SHIFT_TIEBREAK = 1.0


def signed(x):
    return x - 256 if x > 127 else x


# ------------------------------------------------------------------ pe layout
def overlay_start(d):
    """Offset where appended (non-PE) data begins, or 0 if there is none."""
    try:
        e = struct.unpack_from('<I', d, 0x3C)[0]
        if d[e:e + 4] != b'PE\0\0':
            return 0
        nsec = struct.unpack_from('<H', d, e + 6)[0]
        base = e + 24 + struct.unpack_from('<H', d, e + 20)[0]
        end = max(struct.unpack_from('<I', d, base + i * 40 + 16)[0]
                  + struct.unpack_from('<I', d, base + i * 40 + 20)[0]
                  for i in range(nsec))
        return end if 0 < end < len(d) else 0
    except Exception:
        return 0


# -------------------------------------------------------------------- anchors
def classify_header(ct):
    """Tightest (rank, version, family, deltas) consistent with these 8 bytes."""
    best = None
    for v, w in VERSIONS.items():
        pt = [0x91, 0xD5, 0x12, 0x00, w & 0xFF, w >> 8, 0x00, 0x00]
        if not all((pt[i] == pt[j]) == (ct[i] == ct[j])
                   for i in range(8) for j in range(i + 1, 8)):
            continue
        ds = sorted({signed((ct[k] - pt[k]) & 0xFF) for k in NONZERO_AT})
        if ds == [0]:
            fam, rank = "identity", 0
        elif set(ds) <= {-1, 0, 1}:
            fam, rank = "adjacent", 1
        elif max(ds) - min(ds) <= 1:
            fam, rank = "near-shift", 2
        else:
            fam, rank = "full", 3
        if best is None or rank < best[0]:
            best = (rank, v, fam, ds)
    return best


def find_headers(blob):
    out = []
    for i in range(len(blob) - 12):
        if any(blob[i + k] for k in ZERO_AT):
            continue
        if any(blob[i + k] == 0 for k in NONZERO_AT):
            continue
        r = classify_header(blob[i:i + 8])
        if r:
            out.append((i, r))
    out.sort(key=lambda x: (x[1][0], x[0]))
    return out


# ------------------------------------------------ family 1: adjacent-swap keys
def swap_table(a, c):
    t = bytearray(range(256))
    for i in range(127):
        if ((i * a + c) % 127) < 94:
            p = 2 * i + 1
            t[p], t[p + 1] = t[p + 1], t[p]
    return bytes(t)


_SPACE = None


def keyspace():
    """All 8001 distinct adjacent-swap tables."""
    global _SPACE
    if _SPACE is None:
        _SPACE = {}
        for a in range(1, 127):
            for c in range(127):
                _SPACE.setdefault(swap_table(a, c), (a, c))
    return _SPACE


def string_consistency(plain):
    """Count dwords 'L 00 00 00' followed by exactly L printable bytes."""
    good, i, n = 0, 0, len(plain)
    while i < n - 8:
        if not (plain[i + 1] or plain[i + 2] or plain[i + 3]):
            L = plain[i]
            if 4 <= L <= 64 and i + 4 + L < n:
                seg = plain[i + 4:i + 4 + L]
                if all(32 <= ch < 127 for ch in seg) \
                        and not 32 <= plain[i + 4 + L] < 127:
                    good += 1
                    i += 4 + L
                    continue
        i += 1
    return good


def solve_adjacent(data, ver, LM):
    ct = data[:8]
    w = VERSIONS[ver]
    want = bytes([0x91, 0xD5, 0x12, 0x00, w & 0xFF, w >> 8, 0, 0])
    hits = [(a, c, t) for t, (a, c) in keyspace().items()
            if all(t[want[k]] == ct[k] for k in range(8))]
    if not hits:
        return None, 0
    scored = []
    for a, c, t in hits:
        D = bytearray(256)
        for p in range(256):
            D[t[p]] = p
        D = bytes(D)
        plain = data.translate(D)
        sc = string_consistency(plain[:400000])
        arr = np.frombuffer(plain[:400000], dtype=np.uint8)
        pr = (arr >= 32) & (arr < 127)
        m = pr[:-1] & pr[1:]
        lm = float(LM[arr[:-1][m], arr[1:][m]].mean()) if m.any() else -99.0
        scored.append((sc, lm, a, c, D))
    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return scored[0][4], len(hits)


# ------------------------------------------------------ family 2: near-shift
def solve_near_shift(data, shift, pins, LM, rounds=6):
    arr = np.frombuffer(data, dtype=np.uint8)
    lo, hi = (0x20 + shift) & 0xFF, (0x7F + shift) & 0xFF
    a, b = arr[:-1], arr[1:]
    band = ((a >= lo) & (a < hi) & (b >= lo) & (b < hi)) if lo < hi \
        else np.ones(len(a), dtype=bool)
    A = np.zeros((256, 256))
    np.add.at(A, (a[band], b[band]), 1.0)

    prior = np.full((256, 256), -PRIOR_BONUS)
    for c in range(256):
        prior[c, (c - shift) & 0xFF] = SHIFT_TIEBREAK
        prior[c, (c - shift - 1) & 0xFF] = 0.0
    for c, v in pins.items():
        prior[c, :] = -1e9
        prior[c, v] = 1e9

    D = np.array([(c - shift) & 0xFF for c in range(256)], dtype=np.int64)
    for c, v in pins.items():
        D[c] = v
    for _ in range(rounds):
        _, col = linear_sum_assignment(-((A.T @ LM[D]) + (A @ LM[:, D].T) + prior))
        if np.array_equal(col, D):
            break
        D = col
    return bytes(int(x) for x in D)


# ----------------------------------------------------------------- utilities
def build_lm():
    """Byte bigram log-probs. Prefers a real GameMaker project file if present."""
    blob = bytearray()
    for p in glob.glob(os.path.join(REPO, "samples", "*.gm6")) + \
             glob.glob(os.path.join(REPO, "samples", "*.gmd")):
        try:
            blob += open(p, 'rb').read()
        except OSError:
            pass
    if not blob:
        for pat in ("*.bas", "*.frm", "*.txt"):
            for p in glob.glob(os.path.join(REPO, pat)):
                try:
                    blob += open(p, 'rb').read()
                except OSError:
                    pass
    A = np.full((256, 256), 0.5)
    if blob:
        arr = np.frombuffer(bytes(blob), dtype=np.uint8)
        np.add.at(A, (arr[:-1], arr[1:]), 1.0)
    return np.log(A / A.sum())


def to_vb_key(D):
    """VB6 stores key[plaintext] = ciphertext and inverts it itself."""
    K = bytearray(256)
    for c in range(256):
        K[D[c]] = c
    return bytes(K)


def strings_of(p, limit=400000, minlen=8):
    out, cur = [], []
    for ch in p[:limit]:
        if 32 <= ch < 127:
            cur.append(chr(ch))
        else:
            if len(cur) >= minlen:
                out.append("".join(cur))
            cur = []
    return out


# -------------------------------------------------------------------- driver
def recover(path, LM, outdir):
    d = open(path, 'rb').read()
    ov = overlay_start(d)
    blob = d[ov:] if ov else d
    name = os.path.basename(path)
    print(f"\n=== {name}  ({len(d)} bytes; appended data at {ov or 'n/a'})")

    heads = find_headers(blob)
    if not heads:
        print("      no GameMaker data header found")
        return
    off, (rank, ver, fam, ds) = heads[0]
    print(f"      header   : file offset {ov + off} (overlay+{off})")
    print(f"      version  : {ver}")
    print(f"      key family: {fam}   header deltas {ds}")

    data = blob[off:]
    if fam == "identity":
        # An unencrypted header is usually the "included files" block (a
        # filename plus a zlib stream), not the game data itself.
        head = data[:160]
        run, names = [], []
        for ch in head:
            if 32 <= ch < 127:
                run.append(chr(ch))
            else:
                if len(run) >= 5:
                    names.append("".join(run))
                run = []
        if names and b'\x78\x9c' in head:
            print(f"      this is an included-files block ({names[0]!r}, zlib) --")
            print("      NOT the game data; the game data is elsewhere and is encrypted")
        else:
            print("      data is stored UNENCRYPTED -- no key needed")
        return
    if fam == "full":
        print("      unstructured permutation: needs a substitution solver "
              "(not implemented)")
        return

    if fam == "adjacent":
        D, n = solve_adjacent(data, ver, LM)
        print(f"      searched 8001 candidate keys -> {n} matched the header")
        if D is None:
            print("      no key matched; the family guess was wrong")
            return
    else:
        shift = min(ds)
        w = VERSIONS[ver]
        pins = {0: 0, data[0]: 0x91, data[1]: 0xD5, data[2]: 0x12,
                data[4]: w & 0xFF, data[5]: w >> 8}
        D = solve_near_shift(data, shift, pins, LM)

    plain = data.translate(D)
    print(f"      verify   : magic={plain[:4].hex(' ')} "
          f"version={int.from_bytes(plain[4:8], 'little')} "
          f"bijection={sorted(D) == list(range(256))}")

    os.makedirs(outdir, exist_ok=True)
    kp = os.path.join(outdir, name + ".key.txt")
    gp = os.path.join(outdir, name + ".gmd")
    open(kp, 'wb').write(to_vb_key(D))
    open(gp, 'wb').write(plain)
    print(f"      wrote    : {os.path.relpath(kp, REPO)}, {os.path.relpath(gp, REPO)}")
    found = strings_of(plain)
    for s in found[:6]:
        print(f"         {s!r}")


def main():
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("exe", nargs="*")
    ap.add_argument("-o", "--outdir", default=os.path.join(REPO, "recovered"))
    args = ap.parse_args()
    targets = args.exe or sorted(glob.glob(os.path.join(REPO, "samples", "*.exe")))
    if not targets:
        ap.error("no input files")
    LM = build_lm()
    for p in targets:
        recover(p, LM, args.outdir)


if __name__ == '__main__':
    main()
