# Modules/Transformer.py
from __future__ import annotations
from pathlib import Path
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Any, Optional
from Modules.General import safe_name
from Modules.IslandFilter import should_comment_branch, should_drop_branch, drop_mode_enabled

# ------------------------
# Small helpers
# ------------------------
def _t(x: Optional[str]) -> str:
    return "" if x is None else x.strip()

def _f(x: Optional[str]) -> Optional[float]:
    try:
        if x is None:
            return None
        xs = x.strip()
        if xs == "":
            return None
        # handle "330deg" / "30Â°" patterns if ever needed
        if xs.lower().endswith("deg"):
            xs = xs[:-3]
        return float(xs)
    except Exception:
        return None


def _phase_count(phase_str: str) -> int:
    s = (phase_str or "").upper()
    ph = {p for p in s if p in "ABC"}
    return len(ph) if ph else 3


def _decode_conn(code: str) -> Tuple[str, str]:
    """
    Map CYME connection string to (primary, secondary) textual forms.
    e.g. "Yg_Yg", "Y_D", "D_Yg", "D_Y"
    Only topology is reported (wye / delta).
    """
    if not code:
        return "", ""
    parts = code.replace("-", "_").split("_")
    if len(parts) < 2:
        parts = (parts + [""])[:2]

    def one(s: str) -> str:
        s = s.upper()
        if s.startswith("D"):
            return "delta"
        if s.startswith("Y"):
            return "wye"
        return ""

    return one(parts[0]), one(parts[1])


def _bus_labels(bus: str, phase_str: str) -> Tuple[str, str, str]:
    # sanitize bus before composing phase-specific labels
    bus = safe_name(bus)
    s = (phase_str or "ABC").upper()
    labs = []
    for p in "ABC":
        labs.append(f"{bus}_{p.lower()}" if p in s else "")
    return tuple(labs)


def _compute_x_rw(z_percent: Optional[float],
                  xr_ratio: Optional[float]) -> Tuple[Optional[float], Optional[float], Optional[float]]:
    """
    From PositiveSequenceImpedancePercent (Z%) and XRRatio, compute:
      X (pu), RW1 (pu), RW2 (pu).  If inputs missing/invalid â†’ all None.
    """
    if z_percent is None or xr_ratio is None:
        return None, None, None
    try:
        z_pu = float(z_percent) / 100.0
        xr = float(xr_ratio)
    except Exception:
        return None, None, None
    if z_pu <= 0 or xr <= 0:
        return None, None, None

    denom = (1.0 + xr * xr) ** 0.5
    r_total = z_pu / denom
    x_pu = z_pu * xr / denom
    rw = r_total / 2.0
    return round(x_pu, 5), round(rw, 8), round(rw, 8)


def _bounds_from_ntaps(ntaps: Optional[float]) -> tuple[Optional[int], Optional[int]]:
    """
    Symmetric bounds around 0 based on number of taps.
    Examples:
      16 -> (-8, +8)
      17 -> (-8, +8)
       1 -> (0, 0)
    Returns (low, high) as integers, or (None, None) if unavailable.
    """
    if ntaps is None:
        return None, None
    try:
        n = int(round(float(ntaps)))
    except Exception:
        return None, None
    if n <= 0:
        return None, None
    half = n // 2
    return -half, half


# ------------------------
# Read RegulatorDB (for modeling regulators as 1:1 transformers)
# ------------------------
def _read_regulator_db(root: ET.Element) -> Dict[str, Dict[str, Any]]:
    out: Dict[str, Dict[str, Any]] = {}
    for rdb in root.findall(".//RegulatorDB"):
        eid = _t(rdb.findtext("EquipmentID"))
        if not eid:
            continue
        rated_kvln = _f(rdb.findtext("RatedKVLN"))
        rated_kva  = _f(rdb.findtext("RatedKVA")) or _f(rdb.findtext("FirstRatedKVA"))
        maxreg = _f(rdb.findtext("MaximumBoost"))
        minreg = _f(rdb.findtext("MaximumBuck"))
        ntaps  = _f(rdb.findtext("NumberOfTaps"))
        out[eid] = {
            "kvln": rated_kvln,
            "kva": rated_kva,
            "ntaps": ntaps,
            "minreg": minreg,
            "maxreg": maxreg,
        }
    return out


# ------------------------
# Active bus discovery (from Bus sheet)
# ------------------------
_PHASE_SUFFIX_RE = re.compile(r"_(a|b|c)$")

def _active_bus_bases_from_bus_sheet(input_path: Path) -> set[str]:
    """
    Consider a bus active if it appears on the Bus sheet and does NOT start with '//'.
    Return base names (without trailing _a/_b/_c).
    """
    # Lazy import to avoid circular imports
    from Modules.Bus import extract_bus_data

    bases: set[str] = set()
    for row in extract_bus_data(input_path):
        bus = str(row.get("Bus", "")).strip()
        if not bus or bus.startswith("//"):
            continue
        base = _PHASE_SUFFIX_RE.sub("", bus)
        bases.add(base)
    return bases


# ------------------------
# Read TransformerDB
# ------------------------
def _read_transformer_db(root: ET.Element) -> Dict[str, Dict[str, Any]]:
    """
    TransformerDB block: key by EquipmentID.
    We pick only what we actually need; anything not present stays None.
    """
    out: Dict[str, Dict[str, Any]] = {}
    for tdb in root.findall(".//TransformerDB"):
        eid = _t(tdb.findtext("EquipmentID"))
        if not eid:
            continue

        kvp = _f(tdb.findtext("PrimaryVoltage")) or _f(tdb.findtext("PrimaryKV"))
        kvs = _f(tdb.findtext("SecondaryVoltage")) or _f(tdb.findtext("SecondaryKV"))
        kva = _f(tdb.findtext("NominalRatingKVA")) or _f(tdb.findtext("NominalRating"))
        zpct = _f(tdb.findtext("PositiveSequenceImpedancePercent"))
        xr   = _f(tdb.findtext("XRRatio"))
        conn = _t(tdb.findtext("TransformerConnection") or tdb.findtext("Connection"))

        # Pull LTC info from DB
        ltc = tdb.find("./LoadTapChanger")
        ntaps  = _f(ltc.findtext("NumberOfTaps")) if ltc is not None else None
        minreg = _f(ltc.findtext("MinimumRegulationRange")) if ltc is not None else None
        maxreg = _f(ltc.findtext("MaximumRegulationRange")) if ltc is not None else None

        out[eid] = {
            "kvp": kvp, "kvs": kvs, "kva": kva,
            "z_pct": zpct, "xr": xr, "conn": conn,
            "ntaps": ntaps, "minreg": minreg, "maxreg": maxreg,
        }
    return out


# ------------------------
# Parse sections â†’ rows
# ------------------------
def _ltc_fields(xf: ET.Element) -> Dict[str, Optional[float]]:
    """
    Pull tap/range values. Convert *_TapSettingPercent fields from % to pu.
    """
    # --- Tap settings by winding (percent â†’ per-unit) ---
    tap1_pct = _f(xf.findtext("PrimaryTapSettingPercent"))
    tap2_pct = _f(xf.findtext("SecondaryTapSettingPercent"))
    tap3_pct = _f(xf.findtext("TertiaryTapSettingPercent"))

    tap1 = (tap1_pct / 100.0) if tap1_pct is not None else None
    tap2 = (tap2_pct / 100.0) if tap2_pct is not None else None
    tap3 = (tap3_pct / 100.0) if tap3_pct is not None else 1.0  # default 1.0 if tertiary missing

    # LTC block (optional)
    ltc = xf.find("./LTCSettings")
    tap_setting = _f(ltc.findtext("TapSetting")) if ltc is not None else None
    boost = _f(ltc.findtext("Boost")) if ltc is not None else None
    buck  = _f(ltc.findtext("Buck")) if ltc is not None else None

    # If only LTC TapSetting exists (a position, not %), keep it as-is on Tap 2
    if tap2 is None and tap_setting is not None:
        tap2 = tap_setting

    lowest  = _f(xf.findtext("LowestTap")) or _f(xf.findtext("LowestTapPosition"))
    highest = _f(xf.findtext("HighestTap")) or _f(xf.findtext("HighestTapPosition"))

    return {
        "tap1": tap1, "tap2": tap2, "tap3": tap3,
        "low": lowest, "high": highest,
        "min_rng": buck, "max_rng": boost,
    }


def _parse_multiphase_2w_rows(input_path: Path) -> List[List[Any]]:
    """
    Parse <Section><Devices><Transformer> entries and build Multiphase 2W rows.
    No defaults are injected; missing data â†’ empty cells.

    NEW: If either endpoint bus is NOT active on the Bus sheet, we prefix the
         transformer ID with '//' so the row is commented out (avoids dangling devices).
    """
    root = ET.fromstring(input_path.read_text(encoding="utf-8", errors="ignore"))
    # Build island -> source primary per-unit (LN) map from EquivalentSource operating voltage
    from Modules.General import get_island_context
    ctx = get_island_context() or {}
    bus_to_island: Dict[str, int] = dict(ctx.get("bus_to_island", {}))
    # phase-agnostic magnitude per island (use OperatingVoltage1/ KVLL/sqrt(3))
    import math
    island_vpri_pu: Dict[int, float] = {}
    for topo in root.findall(".//Topo"):
        ntype = (topo.findtext("NetworkType") or "").strip().lower()
        eq_mode = (topo.findtext("EquivalentMode") or "").strip()
        if ntype != "substation" or eq_mode == "1":
            continue
        srcs_parent = topo.find("./Sources")
        if srcs_parent is None:
            continue
        for src in srcs_parent.findall("./Source"):
            node_raw = _t(src.findtext("SourceNodeID"))
            node = safe_name(node_raw)
            if not node:
                continue
            eq = src.find("./EquivalentSourceModels/EquivalentSourceModel/EquivalentSource")
            if eq is None:
                continue
            kvll = _f(eq.findtext("KVLL"))
            v1 = _f(eq.findtext("OperatingVoltage1"))
            if kvll is None or v1 is None or kvll == 0:
                continue
            isl = bus_to_island.get(node)
            if isl is None:
                continue
            vbase_ln = float(kvll) / math.sqrt(3.0)
            island_vpri_pu[isl] = float(v1) / vbase_ln

    tdb = _read_transformer_db(root)
    rdb = _read_regulator_db(root)

    # Active bus bases from Bus sheet
    active_bases = _active_bus_bases_from_bus_sheet(input_path)

    rows: List[List[Any]] = []

    for sec in root.findall(".//Section"):
        xf = sec.find(".//Devices/Transformer")
        reg = sec.find(".//Devices/Regulator")
        if xf is None and reg is None:
            continue

        # raw values
        from_bus_raw = _t(sec.findtext("FromNodeID"))
        to_bus_raw   = _t(sec.findtext("ToNodeID"))
        phase        = _t(sec.findtext("Phase") or "ABC").upper()

        # sanitized for output / IDs
        from_bus = safe_name(from_bus_raw)
        to_bus   = safe_name(to_bus_raw)

        # Regulator-only section: export as 1:1 transformer and continue
        if xf is None and reg is not None:
            status_text_r = _t(reg.findtext("ConnectionStatus") or "Connected").lower()
            status_r = 1 if status_text_r == "connected" else 0
            dev_id_r = _t(reg.findtext("DeviceID"))
            conn_code_r = _t(reg.findtext("ConnectionConfiguration"))
            conn_pr, conn_sr = _decode_conn(conn_code_r)

            info_r = rdb.get(dev_id_r, {})
            kvln = info_r.get("kvln")
            kva_r = info_r.get("kva")
            kvp_r = kvln
            kvs_r = kvln

            b1a, b1b, b1c = _bus_labels(from_bus, phase)
            b2a, b2b, b2c = _bus_labels(to_bus, phase)

            rid_r = safe_name(f"VR1_{from_bus}_{to_bus}")
            from_active = from_bus in active_bases
            to_active   = to_bus in active_bases
            island_exclude = should_comment_branch(from_bus, to_bus)
            comment_r = (not from_active) or (not to_active) or island_exclude
            if not (drop_mode_enabled() and (should_drop_branch(from_bus, to_bus) or (not from_active) or (not to_active))):
                rid_out_r = f"//{rid_r}" if comment_r else rid_r
                rows.append([
                    rid_out_r, status_r, _phase_count(phase),
                    b1a, b1b, b1c, kvp_r, kva_r, conn_pr,
                    b2a, b2b, b2c, kvs_r, kva_r, conn_sr,
                    None, None, None,
                    None, None,
                    None, None,
                    None, None, None
                ])
            continue

        # If transformer element missing, skip; Pylance narrowing
        if xf is None:
            continue
        status_text = _t(xf.findtext("ConnectionStatus") or "Connected").lower()
        status = 1 if status_text == "connected" else 0

        dev_id = _t(xf.findtext("DeviceID"))  # keep raw for DB lookup keys

        # Connection type (prefer section value, else DB)
        conn_code = _t(xf.findtext("TransformerConnection"))
        if not conn_code and dev_id in tdb:
            conn_code = _t(tdb[dev_id].get("conn"))
        conn_p, conn_s = _decode_conn(conn_code)

        # Ratings / impedance from DB
        info = tdb.get(dev_id, {})
        kvp  = info.get("kvp")
        kvs  = info.get("kvs")
        kva  = info.get("kva")
        x_pu, rw1, rw2 = _compute_x_rw(info.get("z_pct"), info.get("xr"))

        # Tap / ranges from device
        taps = _ltc_fields(xf)

        # If Lowest/Highest not on the device, compute from DB NumberOfTaps
        low_db, high_db = _bounds_from_ntaps(info.get("ntaps"))
        if taps["low"] is None and low_db is not None:
            taps["low"] = low_db
        if taps["high"] is None and high_db is not None:
            taps["high"] = high_db

        # If Min/Max Range (%) not on the device, use DB regulation ranges
        if taps["min_rng"] is None:
            taps["min_rng"] = info.get("minreg")
        if taps["max_rng"] is None:
            taps["max_rng"] = info.get("maxreg")

        # Compute initial LTC integer tap (applies to all phases equally)
        chosen_n: Optional[int] = None
        try:
            ltc_blk = xf.find("./LTCSettings")
            setpoint = _f(ltc_blk.findtext("SetPoint")) if ltc_blk is not None else None
            tap_loc = _t(ltc_blk.findtext("TapLocation") if ltc_blk is not None else "")
            low = taps.get("low")
            high = taps.get("high")
            # Determine island from either endpoint (prefer from_bus)
            isl_id = bus_to_island.get(from_bus) or bus_to_island.get(to_bus)
            vpri_pu = island_vpri_pu.get(isl_id) if isl_id is not None else None

            # step size from max range and highest tap count
            max_rng = taps.get("max_rng") or taps.get("min_rng")
            if (setpoint is not None) and (low is not None) and (high is not None) and (max_rng is not None) and vpri_pu:
                # Only Voltage120V supported for now
                vset_pu = float(setpoint) / 120.0
                # Narrow types for static analysis; ensure numeric
                if not isinstance(low, (int, float)) or not isinstance(high, (int, float)):
                    raise ValueError("Invalid tap bounds")
                low_i = int(round(low))
                high_i = int(round(high))
                steps = max(abs(high_i), abs(low_i))
                if steps > 0:
                    step = (float(max_rng) / 100.0) / float(steps)
                    if step > 0:
                        # Bias toward the higher resulting secondary voltage relative to setpoint
                        # Secondary control: choose ceil(n_est) so Vsec >= Vset if possible
                        # Primary control: increasing n lowers Vsec; choose floor(n_est) for Vsec >= Vset
                        import math
                        def clamp(nv: int) -> int:
                            return max(low_i, min(high_i, int(nv)))
                        sec_ctrl = not tap_loc.lower().startswith("primary")
                        if sec_ctrl:
                            target_ratio = (vset_pu / float(vpri_pu))
                            n_est = ((target_ratio - 1.0) / step)
                            n_bias = math.ceil(n_est)
                        else:
                            target_ratio = (1.0 / (vset_pu / float(vpri_pu)))
                            n_est = ((target_ratio - 1.0) / step)
                            n_bias = math.floor(n_est)
                        chosen_n = clamp(int(n_bias))
        except Exception:
            pass

        # Bus labels (respect the stated phases) â€” bus names already sanitized
        b1a, b1b, b1c = _bus_labels(from_bus, phase)
        b2a, b2b, b2c = _bus_labels(to_bus, phase)

        # sanitized row ID
        rid = safe_name(f"TR1_{from_bus}_{to_bus}")

        # Comment out if either endpoint is inactive on Bus sheet OR island policy excludes this branch
        from_active = from_bus in active_bases
        to_active   = to_bus in active_bases
        island_exclude = should_comment_branch(from_bus, to_bus)
        comment = (not from_active) or (not to_active) or island_exclude
        # Drop entirely if policy says remove
        if drop_mode_enabled() and (should_drop_branch(from_bus, to_bus) or (not from_active) or (not to_active)):
            continue
        rid_out = f"//{rid}" if comment else rid

        # Final Tap values:
        # LTC controls all three phases together -> set Tap1/2/3 = n
        if chosen_n is not None:
            tap1_out = chosen_n
            tap2_out = chosen_n
            tap3_out = chosen_n
        else:
            tap1_out = taps["tap1"]
            tap2_out = taps["tap2"]
            tap3_out = taps["tap3"]

        rows.append([
            rid_out, status, _phase_count(phase),
            b1a, b1b, b1c, kvp, kva, conn_p,
            b2a, b2b, b2c, kvs, kva, conn_s,
            tap1_out, tap2_out, tap3_out,
            taps["low"], taps["high"],
            taps["min_rng"], taps["max_rng"],
            x_pu, rw1, rw2
        ])

        # --- Regulators modeled as 1:1 transformers ---
        if reg is not None:
            status_text_r = _t(reg.findtext("ConnectionStatus") or "Connected").lower()
            status_r = 1 if status_text_r == "connected" else 0
            dev_id_r = _t(reg.findtext("DeviceID"))
            conn_code_r = _t(reg.findtext("ConnectionConfiguration"))
            conn_pr, conn_sr = _decode_conn(conn_code_r)

            info_r = rdb.get(dev_id_r, {})
            kvln = info_r.get("kvln")
            kva_r = info_r.get("kva")
            kvp_r = kvln
            kvs_r = kvln

            b1a, b1b, b1c = _bus_labels(from_bus, phase)
            b2a, b2b, b2c = _bus_labels(to_bus, phase)

            rid_r = safe_name(f"VR1_{from_bus}_{to_bus}")
            from_active = from_bus in active_bases
            to_active   = to_bus in active_bases
            island_exclude = should_comment_branch(from_bus, to_bus)
            comment_r = (not from_active) or (not to_active) or island_exclude
            if drop_mode_enabled() and (should_drop_branch(from_bus, to_bus) or (not from_active) or (not to_active)):
                pass
            else:
                rid_out_r = f"//{rid_r}" if comment_r else rid_r
                rows.append([
                    rid_out_r, status_r, _phase_count(phase),
                    b1a, b1b, b1c, kvp_r, kva_r, conn_pr,
                    b2a, b2b, b2c, kvs_r, kva_r, conn_sr,
                    None, None, None,
                    None, None,
                    None, None,
                    None, None, None
                ])

    return rows


# ------------------------
# Sheet writer
# ------------------------
def write_transformer_sheet(xw, input_path: Path) -> None:
    """
    Build the 'Transformer' sheet.
    Absolutely no hard-coded engineering values are inserted:
    if a datum is missing in the file/DB, the cell is left blank.
    """
    wb = xw.book
    ws = wb.add_worksheet("Transformer")
    xw.sheets["Transformer"] = ws

    # Formats
    bold = wb.add_format({"bold": True})
    link_fmt = wb.add_format({"font_color": "blue", "underline": 1})
    notes_hdr = wb.add_format({"bold": True, "font_color": "yellow", "bg_color": "#595959", "align": "left"})
    notes_txt = wb.add_format({"font_color": "yellow", "bg_color": "#595959"})
    th   = wb.add_format({"bold": True, "bottom": 1})
    f0   = wb.add_format({"num_format": "0"})
    f2   = wb.add_format({"num_format": "0.00"})
    f5   = wb.add_format({"num_format": "0.00000"})
    f8   = wb.add_format({"num_format": "0.00000000"})

    # Column widths
    widths = [22, 8, 16, 12, 12, 12, 8, 12, 12, 12, 12, 12, 8, 12, 12,
              10, 10, 10, 12, 12, 14, 14, 10, 12, 12]
    for c, w in enumerate(widths):
        ws.set_column(c, c, w)

    # Data
    rows_mp2w = _parse_multiphase_2w_rows(input_path)

    # Anchors
    r = 10
    b1_t, b1_h, b1_e = r, r+1, r+2; r = b1_e + 2                # PosSeq 2W (empty)
    b2_t, b2_h, b2_e = r, r+1, r+2; r = b2_e + 2                # PosSeq 3W (empty)
    b3_t, b3_h, b3_first = r, r+1, r+2; b3_e = b3_first + len(rows_mp2w); r = b3_e + 2
    b4_t, b4_h, b4_e = r, r+1, r+2                               # MP 2W with Mutual (empty)

    # Top links + notes
    ws.write(0, 0, "Type", bold)
    ws.write_url(1, 0, f"internal:'Transformer'!A{b1_h+1}:K{b1_e+1}", link_fmt, "PositiveSeq2wXF")
    ws.write_url(2, 0, f"internal:'Transformer'!A{b2_h+1}:S{b2_e+1}", link_fmt, "PositiveSeq3wXF")
    ws.write_url(3, 0, f"internal:'Transformer'!A{b3_h+1}:Y{b3_e+1}", link_fmt, "Multiphase2wXF")
    ws.write_url(4, 0, f"internal:'Transformer'!A{b4_h+1}:AA{b4_e+1}", link_fmt, "Multiphase2wXFMutual")

    ws.merge_range(7, 0, 7, 7, "Important notes:", notes_hdr)
    ws.merge_range(8, 0, 8, 7, "Default order of blocks and columns after row 11 must not change", notes_txt)
    ws.merge_range(9, 0, 9, 7, "One empty row between End of each block and the next block is mandatory; otherwise, empty rows are NOT allowed", notes_txt)

    go_top = "internal:'Transformer'!A1"

    # ---- Block 1: Positive-Sequence 2W (empty) ----
    ws.merge_range(b1_t, 0, b1_t, 2, "Positive-Sequence 2W-Transformer", bold)
    ws.write_url(b1_t, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(b1_h, 0, ["ID","Status","From bus","To bus","R (pu)","Xl (pu)","Gmag (pu)","Bmag (pu)","Ratio W1 (pu)","Ratio W2 (pu)","Phase Shift (deg)"], th)
    ws.merge_range(b1_e, 0, b1_e, 2, "End of Positive-Sequence 2W-Transformer")

    # ---- Block 2: Positive-Sequence 3W (empty) ----
    ws.merge_range(b2_t, 0, b2_t, 2, "Positive-Sequence 3W-Transformer", bold)
    ws.write_url(b2_t, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(b2_h, 0, ["ID","Status","Bus1","Bus2","Bus3","R_12 (pu)","Xl_12 (pu)","R_23 (pu)","Xl_23 (pu)","R_31 (pu)","Xl_31 (pu)","Gmag (pu)","Bmag (pu)","Ratio W1 (pu)","Ratio W2 (pu)","Ratio W3 (pu)","Phase Shift W1 (deg)","Phase Shift W2 (deg)","Phase Shift W3 (deg)"], th)
    ws.merge_range(b2_e, 0, b2_e, 2, "End of Positive-Sequence 3W-Transformer")

    # ---- Block 3: Multiphase 2W (parsed) ----
    ws.merge_range(b3_t, 0, b3_t, 2, "Multiphase 2W-Transformer", bold)
    ws.write_url(b3_t, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(
        b3_h, 0,
        ["ID","Status","Number of phases",
         "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
         "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
         "Tap 1","Tap 2","Tap 3","Lowest Tap","Highest Tap","Min Range (%)","Max Range (%)",
         "X (pu)","RW1 (pu)","RW2 (pu)"],
        th
    )

    rr = b3_first
    for row in rows_mp2w:
        # helper to write optional numbers / text
        def wnum(c, v, fmt):
            if v is None or v == "":
                ws.write(rr, c, "")
            else:
                ws.write_number(rr, c, v, fmt)

        ws.write(rr, 0, row[0])                 # ID (may be prefixed with //)
        wnum(1, row[1],  f0)                    # Status
        wnum(2, row[2],  f0)                    # Number of phases

        ws.write(rr, 3, row[3]); ws.write(rr, 4, row[4]); ws.write(rr, 5, row[5])
        wnum(6,  row[6],  f2)                   # Vp
        wnum(7,  row[7],  f2)                   # Sbase p
        ws.write(rr, 8, row[8])                 # Conn p

        ws.write(rr, 9, row[9]); ws.write(rr,10, row[10]); ws.write(rr,11, row[11])
        wnum(12, row[12], f2)                   # Vs
        wnum(13, row[13], f2)                   # Sbase s
        ws.write(rr,14, row[14])                # Conn s

        wnum(15, row[15], f2)                   # Tap1
        wnum(16, row[16], f2)                   # Tap2
        wnum(17, row[17], f2)                   # Tap3
        wnum(18, row[18], f2)                   # Lowest tap
        wnum(19, row[19], f2)                   # Highest tap
        wnum(20, row[20], f2)                   # Min range (%)
        wnum(21, row[21], f2)                   # Max range (%)

        wnum(22, row[22], f5)                   # X (pu)
        wnum(23, row[23], f8)                   # RW1 (pu)
        wnum(24, row[24], f8)                   # RW2 (pu)

        rr += 1

    ws.merge_range(b3_e, 0, b3_e, 3, "End of Multiphase 2W-Transformer")

    # ---- Block 4: Multiphase 2W with Mutual (template only) ----
    ws.merge_range(b4_t, 0, b4_t, 2, "Multiphase 2W-Transformer with Mutual Impedance", bold)
    ws.write_url(b4_t, 3, go_top, link_fmt, "Go to Type List")
    ws.write_row(
        b4_h, 0,
            ["ID","Status","Number of phases",
            "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
            "Bus1","Bus2","Bus3","V (kV)","S_base (kVA)","Conn. type",
            "Tap 1","Tap 2","Tap 3","Lowest Tap","Highest Tap","Min Range (%)","Max Range (%)",
            "Z0 leakage (pu)","Z1 leakage (pu)","X0/R0","X1/R1","No Load Loss (kW)"],
        th
    )
    ws.merge_range(b4_e, 0, b4_e, 3, "End of Multiphase 2W-Transformer with Mutual Impedance")
