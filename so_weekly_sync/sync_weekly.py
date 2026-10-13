#!/usr/bin/env python3
"""Weekly Sell-Out sync: Google Sheet (per-retailer tabs) -> Feishu spreadsheet.

Data flow per tab:
  1. Download ONLY that tab as CSV via the gviz endpoint (fast; no full-workbook export).
  2. Locate the header row (longest ascending 1,2,3... run), the EAN column
     (column with the most 8-14 digit codes) and the target week column
     (header == week and left neighbour == week-1, which disambiguates from the
     trailing cumulative/summary column that repeats the current week number).
  3. Read the target Feishu tab, match rows BY EAN (never by row position),
     and write the week value into the W<week> column (created at the far right
     if it does not exist yet -> historical accumulation).
  4. EANs present in the target but absent in the source are left untouched and
     reported in the missing-EAN log. Configurable anchors (known EAN->value
     pairs) are verified BEFORE any write; a mismatch aborts the run so that a
     mis-detected column can never silently corrupt the target.

Backends:
  larkcli - uses `lark-cli` (user identity on a MiCo Desktop runtime). Default
            when lark-cli is on PATH.
  openapi - Lark Open API with app credentials from env LARK_APP_ID /
            LARK_APP_SECRET (tenant_access_token). For GitHub Actions / CI.
  auto    - larkcli if available, else openapi.

Stdlib only; no third-party dependencies.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import os
import re
import subprocess
import sys
import urllib.parse
import urllib.request
from typing import Any

EAN_RE = re.compile(r"^\d{8,14}$")
GA_ID = "https://open.feishu.cn"
UA = {"User-Agent": "Mozilla/5.0 (so-weekly-sync)"}


# --------------------------------------------------------------------------- utils

def log(msg: str) -> None:
    print(msg, file=sys.stderr, flush=True)


def die(msg: str, code: int = 2) -> "NoReturn":  # type: ignore[valid-type]
    log(f"FATAL: {msg}")
    sys.exit(code)


def parse_num(s: str) -> Any:
    """' 0 ' -> 0, '1058' -> 1058, '37,5' -> 37.5, '' -> None, junk -> None."""
    if s is None:
        return None
    t = str(s).replace("\u00a0", "").replace(" ", "")
    if t == "":
        return None
    if re.fullmatch(r"-?\d+", t):
        return int(t)
    t2 = t.replace(",", ".")
    try:
        f = float(t2)
    except ValueError:
        return None
    return int(f) if f.is_integer() else f


def col_letter(idx0: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA."""
    out = ""
    n = idx0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def last_week_complete(today: dt.date | None = None) -> int:
    """ISO week number of the last *complete* ISO week (today's week - 1)."""
    today = today or dt.date.today()
    iso = today.isocalendar()
    return iso.week - 1 if iso.week > 1 else 52


# --------------------------------------------------------------------------- source (Google Sheet)

def gviz_url(sheet_id: str, tab: str) -> str:
    return (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq"
        f"?tqx=out:csv&sheet={urllib.parse.quote(tab)}"
    )


def fetch_tab_csv(sheet_id: str, tab: str) -> list[list[str]]:
    req = urllib.request.Request(gviz_url(sheet_id, tab), headers=UA)
    with urllib.request.urlopen(req, timeout=60) as resp:
        raw = resp.read().decode("utf-8-sig", "replace")
    return list(csv.reader(io.StringIO(raw)))


def find_header_row(rows: list[list[str]], max_scan: int = 10) -> int:
    """Header row = row holding the longest ascending run 1,2,3,... (week columns).

    Robust against lost label text and against the 'Ultima week: <n>' banner row,
    whose bare week number is not part of an ascending run.
    """
    best_row, best_len = -1, 0
    for i in range(min(max_scan, len(rows))):
        row = [c.strip() for c in rows[i]]
        run = 0
        for j in range(1, len(row)):
            if row[j].isdigit() and row[j - 1].isdigit() and int(row[j]) == int(row[j - 1]) + 1:
                run = max(run, 2)
            elif row[j] == "1" and (j == 0 or not row[j - 1].isdigit()):
                # potential run start; count forward
                k, n = j, 1
                while k + 1 < len(row) and row[k + 1].isdigit() and int(row[k + 1]) == n + 1:
                    k += 1
                    n += 1
                if n > run:
                    run = n
        if run > best_len:
            best_row, best_len = i, run
    if best_row < 0 or best_len < 4:
        die(f"could not locate a week header row (best ascending run length={best_len})")
    return best_row


def week_columns(rows: list[list[str]], hidx: int) -> dict[int, int]:
    """week number -> column index, from the ascending run in the header row."""
    row = [c.strip() for c in rows[hidx]]
    weeks: dict[int, int] = {}
    for j, v in enumerate(row):
        if v == "1" and (j == 0 or not row[j - 1].isdigit()):
            k, n = j, 1
            weeks[1] = j
            while k + 1 < len(row) and row[k + 1].isdigit() and int(row[k + 1]) == n + 1:
                k += 1
                n += 1
                weeks[n] = k
            if len(weeks) >= 4:
                return weeks
        weeks = {}
    die("no ascending week run found in header row")
    return weeks


def ean_column(rows: list[list[str]], hidx: int) -> int:
    best_col, best_cnt = -1, 0
    width = max(len(r) for r in rows)
    for j in range(width):
        cnt = 0
        for r in rows[hidx + 1:]:
            if j < len(r) and EAN_RE.match(r[j].strip()):
                cnt += 1
        if cnt > best_cnt:
            best_col, best_cnt = j, cnt
    if best_col < 0 or best_cnt < 5:
        die(f"EAN column detection failed (best count={best_cnt})")
    return best_col


def load_source_tab(sheet_id: str, tab: str, week: int) -> tuple[dict[str, Any], dict[str, Any], list[list[str]]]:
    """Return (ean -> week value, diagnostics, raw rows). Aborts on undetectable structure."""
    rows = fetch_tab_csv(sheet_id, tab)
    head_text = " ".join(c for r in rows[:3] for c in r)
    if "ltima week" not in head_text:
        # Some tabs (e.g. Expert) have a blank banner cell; structural checks
        # below still guard against gviz's silent fallback to the first sheet.
        log(f"[warn] tab '{tab}': no 'Ultima week' banner in first rows; "
            f"relying on structural checks only")
    hidx = find_header_row(rows)
    weeks = week_columns(rows, hidx)
    if week not in weeks:
        die(f"tab '{tab}': target week {week} not present in header (weeks {min(weeks)}..{max(weeks)})")
    ecol = ean_column(rows, hidx)
    wcol = weeks[week]
    # secondary disambiguation check required by spec: left neighbour must be week-1
    hdr = [c.strip() for c in rows[hidx]]
    if week > 1 and hdr[wcol - 1] != str(week - 1):
        die(f"tab '{tab}': week {week} column {col_letter(wcol)} has left neighbour "
            f"{hdr[wcol-1]!r}, expected '{week-1}' - refusing to guess")
    data: dict[str, Any] = {}
    for r in rows[hidx + 1:]:
        if len(r) > ecol:
            ean = r[ecol].strip()
            if EAN_RE.match(ean):
                val = parse_num(r[wcol]) if wcol < len(r) else None
                data[ean] = val
    if not data:
        die(f"tab '{tab}': no EAN rows found")
    diag = {"tab": tab, "header_row0": hidx, "ean_col": col_letter(ecol),
            "week_col": col_letter(wcol), "source_eans": len(data),
            "nonempty": sum(1 for v in data.values() if v is not None)}
    log(f"[source] {tab}: header@row{hidx + 1} EAN={col_letter(ecol)} "
        f"W{week}={col_letter(wcol)} eans={diag['source_eans']} values={diag['nonempty']}")
    return data, diag, rows


def cell_value(rows: list[list[str]], hidx: int, ecol: int, weeks: dict[int, int],
               ean: str, week: int) -> Any:
    wcol = weeks[week]
    for r in rows[hidx + 1:]:
        if len(r) > ecol and r[ecol].strip() == ean:
            return parse_num(r[wcol]) if wcol < len(r) else None
    return None


def check_anchors(tab: str, rows: list[list[str]], anchors: list[dict]) -> None:
    """Verify pinned (EAN, week) -> value pairs on the *source* before any write.

    Anchors pin a completed historical week so they stay valid week over week;
    any mismatch means column detection drifted -> abort.
    """
    if not anchors:
        return
    hidx = find_header_row(rows)
    weeks = week_columns(rows, hidx)
    ecol = ean_column(rows, hidx)
    for a in anchors:
        ean, w, expected = a["ean"], a["week"], a["value"]
        if w not in weeks:
            die(f"anchor failed for tab '{tab}': week {w} not in source header")
        got = cell_value(rows, hidx, ecol, weeks, ean, w)
        if got != expected:
            die(f"anchor failed for tab '{tab}': EAN {ean} W{w} expected {expected}, got {got!r} "
                f"- column detection is off; aborting BEFORE any write")
    log(f"[anchor] {tab}: {len(anchors)} anchor(s) OK")


# --------------------------------------------------------------------------- target (Feishu)

class LarkClient:
    def __init__(self, backend: str, wiki_url: str):
        self.backend = backend
        self.wiki_url = wiki_url
        self.token = wiki_url
        self.sheet_meta: dict[str, dict[str, Any]] = {}
        if backend == "openapi":
            self.app_id = os.environ.get("LARK_APP_ID", "")
            self.app_secret = os.environ.get("LARK_APP_SECRET", "")
            if not (self.app_id and self.app_secret):
                die("openapi backend requires env LARK_APP_ID and LARK_APP_SECRET")
            self._tenant = self._http("POST", "/open-apis/auth/v3/tenant_access_token/internal",
                                      body={"app_id": self.app_id, "app_secret": self.app_secret},
                                      auth=False)["tenant_access_token"]
            m = re.search(r"/wiki/([A-Za-z0-9]+)", wiki_url)
            if not m:
                die(f"target URL is not a wiki URL: {wiki_url}")
            node = self._http("GET", f"/open-apis/wiki/v2/spaces/get_node?token={m.group(1)}")
            self.token = node["node"]["obj_token"]
        else:
            if not shutil_which("lark-cli"):
                die("larkcli backend requested but lark-cli not found on PATH")
        self._load_sheets()

    # -- low level ---------------------------------------------------------
    def _http(self, method: str, path: str, body: dict | None = None,
              auth: bool = True) -> dict:
        url = GA_ID + path
        data = json.dumps(body).encode() if body is not None else None
        headers = {"Content-Type": "application/json; charset=utf-8"}
        if auth:
            headers["Authorization"] = f"Bearer {self._tenant}"
        req = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                out = json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            die(f"Open API HTTP {e.code}: {e.read().decode(errors='replace')[:300]}")
        if out.get("code") != 0:
            die(f"Open API error {out.get('code')}: {out.get('msg')}")
        return out.get("data", out)

    def _cli(self, *args: str) -> dict:
        proc = subprocess.run(["lark-cli", *args], capture_output=True, text=True,
                              encoding="utf-8", errors="replace")
        if proc.returncode != 0:
            die(f"lark-cli {' '.join(args[:2])} failed (rc={proc.returncode}): "
                f"{proc.stderr.strip()[:400]}")
        try:
            return json.loads(proc.stdout)
        except json.JSONDecodeError:
            die(f"lark-cli returned non-JSON stdout: {proc.stdout[:200]}")

    # -- sheets ------------------------------------------------------------
    def _load_sheets(self) -> None:
        if self.backend == "larkcli":
            out = self._cli("sheets", "+workbook-info", "--url", self.wiki_url)
            sheets = out["data"]["sheets"]
        else:
            sheets = self._http("GET", f"/open-apis/sheets/v2/spreadsheets/{self.token}/metainfo")["sheets"]
        for s in sheets:
            name = s.get("sheet_name") or s.get("title")
            self.sheet_meta[name] = {"id": s["sheet_id"],
                                     "rows": s.get("grid_properties", s.get("gridProperties", {})).get("row_count", s.get("row_count", 200)),
                                     "cols": s.get("grid_properties", s.get("gridProperties", {})).get("column_count", s.get("column_count", 26))}

    CHUNK = 40  # rows per cells-get call; keeps each response well under output limits

    def read(self, tab: str, rng: str) -> list[list[str]]:
        """Return a dense matrix of string values for the range (blank -> '').

        The larkcli path paginates by rows and asserts full coverage, so a
        truncated response can never be mistaken for the whole range.
        """
        m = re.match(r"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", rng)
        a, b, c, d = m.groups()
        r0, r1 = int(b), int(d)
        if self.backend == "larkcli":
            out_rows: list[list[str]] = []
            cur = r0
            while cur <= r1:
                end = min(cur + self.CHUNK - 1, r1)
                part = f"{a}{cur}:{c}{end}"
                out = self._cli("sheets", "+cells-get", "--url", self.wiki_url,
                                "--sheet-name", tab, "--range", part, "--include", "value")
                data = out["data"]
                rr = data["ranges"][0]
                if data.get("has_more") or rr.get("truncated"):
                    die(f"cells-get returned a truncated page for {tab}!{part}; "
                        f"refusing to sync from a partial read")
                cols_idx = rr["col_indices"]
                width = len(cols_idx)
                for i, real_row in enumerate(rr["row_indices"]):
                    if real_row != cur + i:
                        die(f"row gap in {tab}!{part}: expected row {cur + i}, got {real_row}")
                    row_vals = []
                    for cell in rr["cells"][i][:width]:
                        v = cell.get("value") if isinstance(cell, dict) else None
                        row_vals.append("" if v is None else str(v).strip())
                    out_rows.append(row_vals + [""] * (width - len(row_vals)))
                cur = end + 1
            return out_rows
        path = (f"/open-apis/sheets/v2/spreadsheets/{self.token}/values"
                f"/{self.sheet_meta[tab]['id']}!{a}{b}:{c}{d}?valueRenderOption=ToString")
        vals = self._http("GET", path).get("valueRange", {}).get("values", [])
        out = []
        for row in vals:
            out.append(["" if v is None else str(v).strip() for v in row] + [""] * 0)
        width = col_index(c) - col_index(a) + 1
        return [row[:width] + [""] * (width - len(row)) for row in out]

    def write_runs(self, tab: str, runs: list[tuple[str, list[list[Any]]]]) -> None:
        """runs: list of (A1 range, [[value], ...]) - only non-empty runs are passed."""
        if not runs:
            return
        if self.backend == "larkcli":
            for rng, vals in runs:
                cells = json.dumps([[{"value": v} for v in col] for col in vals])
                self._cli("sheets", "+cells-set", "--url", self.wiki_url,
                          "--sheet-name", tab, "--range", rng, "--cells", cells)
        else:
            payload = {"valueRanges": [
                {"range": f"{self.sheet_meta[tab]['id']}!{rng}", "values": vals}
                for rng, vals in runs]}
            self._http("POST", f"/open-apis/sheets/v2/spreadsheets/{self.token}/values_batch_update",
                       body=payload)

    def insert_col(self, tab: str, at_letter: str) -> None:
        idx = col_index(at_letter)
        if self.backend == "larkcli":
            self._cli("sheets", "+dim-insert", "--url", self.wiki_url, "--sheet-name", tab,
                      "--position", at_letter, "--count", "1", "--inherit-style", "before")
        else:
            self._http("POST", f"/open-apis/sheets/v2/spreadsheets/{self.token}/dimension_range",
                       body={"dimension": {"sheetId": self.sheet_meta[tab]["id"],
                                           "majorDimension": "COLUMNS",
                                           "startIndex": idx, "endIndex": idx + 1}})


def shutil_which(cmd: str) -> str | None:
    for p in os.environ.get("PATH", "").split(os.pathsep):
        f = os.path.join(p, cmd)
        if os.path.isfile(f) and (os.name == "nt" or os.access(f, os.X_OK)):
            return f
    return None


# --------------------------------------------------------------------------- sync core

def sync_tab(client: LarkClient, tab: str, src: dict[str, Any], week: int,
             dry_run: bool) -> dict[str, Any]:
    meta = client.sheet_meta[tab]
    width = min(meta["cols"], 52)
    grid = client.read(tab, f"A1:{col_letter(width - 1)}{meta['rows']}")

    header = grid[0]
    week_hdr = f"W{week}"
    if week_hdr in header:
        wcol = header.index(week_hdr)
        created = False
    else:
        # append at far right: after the last W<n> header (historical accumulation)
        wcols = [i for i, v in enumerate(header) if re.fullmatch(r"W\d{1,2}", v)]
        wcol = (max(wcols) + 1) if wcols else 2
        if wcol >= meta["cols"]:
            client.insert_col(tab, col_letter(wcol))
            meta["cols"] += 1
        created = True

    # rows: skip blanks in col A; EAN compare is stripped (MI tab has trailing spaces)
    plan: list[tuple[int, Any]] = []       # (row1, value)
    missing: list[str] = []
    matched = 0
    for i in range(1, len(grid)):
        ean = grid[i][0].strip()
        if not ean:
            continue
        if ean in src and src[ean] is not None:
            matched += 1
            plan.append((i + 1, src[ean]))
        elif ean not in src:
            missing.append(ean)

    # contiguous runs of planned rows (values only; we never write blanks,
    # so untouched cells keep their previous content)
    runs: list[tuple[str, list[list[Any]]]] = []
    if created:
        runs.append((f"{col_letter(wcol)}1", [[week_hdr]]))
    cur_rows: list[int] = []
    cur_vals: list[Any] = []
    for row1, val in plan:
        if cur_rows and row1 == cur_rows[-1] + 1:
            cur_rows.append(row1)
            cur_vals.append(val)
        else:
            if cur_rows:
                runs.append((f"{col_letter(wcol)}{cur_rows[0]}:{col_letter(wcol)}{cur_rows[-1]}",
                             [[v] for v in cur_vals]))
            cur_rows, cur_vals = [row1], [val]
    if cur_rows:
        runs.append((f"{col_letter(wcol)}{cur_rows[0]}:{col_letter(wcol)}{cur_rows[-1]}",
                     [[v] for v in cur_vals]))

    log(f"[plan] {tab}: W{week} -> column {col_letter(wcol)} "
        f"({'new' if created else 'existing'}), {matched} values in {len(runs)} write run(s), "
        f"{len(missing)} missing EAN(s)")
    if not dry_run:
        client.write_runs(tab, runs)
    return {"tab": tab, "week_column": col_letter(wcol), "created": created,
            "written": matched, "runs": len(runs), "missing": missing,
            "dry_run": dry_run}


def verify_tab(client: LarkClient, tab: str, src: dict[str, Any], res: dict, week: int) -> int:
    wcol = col_index(res["week_column"])
    meta = client.sheet_meta[tab]
    grid = client.read(tab, f"A1:{res['week_column']}{meta['rows']}")
    bad = 0
    for i in range(1, len(grid)):
        ean = grid[i][0].strip()
        if ean and ean in src and src[ean] is not None:
            got = parse_num(grid[i][wcol])
            if got != src[ean]:
                bad += 1
                log(f"[verify] {tab} row{i + 1} EAN {ean}: expected {src[ean]}, got {grid[i][wcol]!r}")
    log(f"[verify] {tab}: {'OK' if bad == 0 else f'{bad} MISMATCH(ES)'}")
    return bad


# --------------------------------------------------------------------------- main

def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json"),
                    help="config JSON path (default: config.json next to this script)")
    ap.add_argument("--week", type=int, default=None,
                    help="ISO week number to sync (default: config.week, else last complete ISO week)")
    ap.add_argument("--tabs", nargs="*", default=None, help="subset of tabs to sync")
    ap.add_argument("--dry-run", action="store_true", help="plan + anchors, no writes")
    ap.add_argument("--backend", choices=["auto", "larkcli", "openapi"], default="auto")
    args = ap.parse_args()

    with open(args.config, encoding="utf-8") as f:
        cfg = json.load(f)

    week = args.week or cfg.get("week") or last_week_complete()
    sheet_id = cfg["source_sheet_id"]
    default_source = sheet_id == cfg.get("default_source_sheet_id")
    tabs = args.tabs or cfg["tabs"]
    anchors_all = cfg.get("anchors", {}) if default_source else {}

    backend = args.backend
    if backend == "auto":
        backend = "larkcli" if shutil_which("lark-cli") else "openapi"
    log(f"[run] week=W{week} backend={backend} tabs={','.join(tabs)} dry_run={args.dry_run}")

    client = LarkClient(backend, cfg["target_url"])
    for t in tabs:
        if t not in client.sheet_meta:
            die(f"target workbook has no tab named '{t}' (tabs: {sorted(client.sheet_meta)})")

    # Phase 1: load every source tab and pass ALL anchors before any write,
    # so a bad column detection on a later tab cannot leave earlier tabs half-synced.
    loaded: dict[str, tuple[dict, list]] = {}
    for t in tabs:
        src, _diag, rows = load_source_tab(sheet_id, t, week)
        check_anchors(t, rows, anchors_all.get(t, []))
        loaded[t] = (src, rows)

    # Phase 2: plan + write.
    total_bad = 0
    results = [sync_tab(client, t, loaded[t][0], week, args.dry_run) for t in tabs]

    if not args.dry_run:
        for res, t in zip(results, tabs):
            total_bad += verify_tab(client, t, loaded[t][0], res, week)

    summary = {"week": week, "backend": backend, "dry_run": args.dry_run,
               "tabs": [{k: v for k, v in r.items() if k != "missing"} for r in results],
               "missing_eans": {r["tab"]: r["missing"] for r in results if r["missing"]},
               "verify_mismatches": total_bad}
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if total_bad:
        die(f"verification found {total_bad} mismatch(es)", code=3)
    log("[done] sync complete" + (" (dry-run)" if args.dry_run else ""))


if __name__ == "__main__":
    main()
