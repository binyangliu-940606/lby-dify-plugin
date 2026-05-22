import json, math
import numpy as np

def truncated_normal(n, mean, sd, low, high, rng):
    # rejection sampling for simplicity
    if sd <= 0:
        return np.clip(np.full(n, mean), low, high)
    out = []
    while len(out) < n:
        x = rng.normal(mean, sd, size=n)
        x = x[(x >= low) & (x <= high)]
        out.extend(x.tolist())
    return np.array(out[:n])

def calibrate_mean_sd(x, target_mean, target_sd, low, high):
    # linear transform to match mean/sd then clip; repeat a few times
    x = np.array(x, dtype=float)
    for _ in range(5):
        m = float(np.mean(x))
        s = float(np.std(x, ddof=0))
        if s == 0:
            x = np.clip(x, low, high)
            continue
        x = (x - m) * (target_sd / s) + target_mean
        x = np.clip(x, low, high)
    return x

def set_continuous_by_group(rows, group_var, gname, var, mean, sd, rng, ranges, do_calibrate=True):
    idxs = [i for i,r in enumerate(rows) if r[group_var]==gname]
    n = len(idxs)
    low, high = ranges.get(var, [None, None])
    low = -1e9 if low is None else low
    high = 1e9 if high is None else high

    if mean is None or sd is None:
        # fallback random within range
        if not math.isfinite(low) or not math.isfinite(high) or low >= high:
            low, high = 0, 1
        x = rng.uniform(low, high, size=n)
    else:
        x = truncated_normal(n, mean, sd, low, high, rng)
        if do_calibrate:
            x = calibrate_mean_sd(x, mean, sd, low, high)

    for i,val in zip(idxs, x.tolist()):
        rows[i][var] = float(val)

def derive_fields(row):
    # consistent fields (examples)
    if "sex" in row:
        row["sex_male"] = 1 if row["sex"] == "male" else 0
    # good_outcome from gos_group
    if "gos_group" in row:
        row["good_outcome"] = 1 if row["gos_group"] == "high" else 0
    # gos_score: sample within group
    if "gos_group" in row and "gos_score" not in row:
        if row["gos_group"] == "low":
            row["gos_score"] = int(np.random.choice([1,2,3]))
        else:
            row["gos_score"] = int(np.random.choice([4,5]))
    return row

def apply_threshold_rules(row, rules):
    # rules: {field: {"expr": "..."}}
    # safe eval with limited names
    env = {k: row.get(k) for k in row.keys()}
    for out_field, meta in rules.items():
        expr = meta.get("expr")
        if not expr:
            continue
        try:
            row[out_field] = 1 if eval(expr, {"__builtins__": {}}, env) else 0
        except:
            # if missing vars, set null
            row[out_field] = None
    return row

def generate_dataset(spec, frame, seed=42):
    rng = np.random.default_rng(seed)
    rows = json.loads(frame) if isinstance(frame, str) else frame
    group_var = spec["study"].get("group_var","gos_group")

    cont = spec.get("continuous", {})
    ranges = spec.get("ranges", {})

    # for each continuous variable, set by group
    for var, meta in cont.items():
        by_group = (meta or {}).get("by_group") or {}
        for gname in spec["study"]["groups"].keys():
            gstats = by_group.get(gname) or {}
            mean = gstats.get("mean")
            sd = gstats.get("sd")
            set_continuous_by_group(rows, group_var, gname, var, mean, sd, rng, ranges, do_calibrate=True)

    # derive fields + threshold rules
    rules = spec.get("threshold_rules", {})
    for r in rows:
        r = derive_fields(r)
        apply_threshold_rules(r, rules)

        # treatment consistency example (optional)
        if "treatment_method" in r and "treatment" not in r:
            r["treatment"] = "surgical" if r["treatment_method"] == "surgery" else "conservative"

    field_schema = spec.get("field_schema", {})
    postprocess_types_and_rounding(rows, field_schema)
    return rows
















def normalize_yes_no(v):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip().lower()
        if s in {"yes", "y", "1", "true", "t"}:
            return "yes"
        if s in {"no", "n", "0", "false", "f"}:
            return "no"
        # unknown string: keep as-is
        return v
    if isinstance(v, (int, float)):
        return "yes" if float(v) >= 0.5 else "no"
    if isinstance(v, bool):
        return "yes" if v else "no"
    return v

def clip_value(x, mn, mx):
    if x is None:
        return None
    try:
        xf = float(x)
    except:
        return None
    if mn is not None:
        xf = max(xf, float(mn))
    if mx is not None:
        xf = min(xf, float(mx))
    return xf

def round_value(x, decimals):
    if x is None:
        return None
    if decimals is None:
        return x
    try:
        return round(float(x), int(decimals))
    except:
        return x

def cast_value(x, typ, decimals=None):
    if x is None:
        return None
    if typ == "int":
        try:
            # 先按 decimals round 再 int，避免 1.999 这种
            y = round_value(x, decimals if decimals is not None else 0)
            return int(float(y))
        except:
            return None
    if typ == "float":
        try:
            y = round_value(x, decimals if decimals is not None else 2)
            return float(y)
        except:
            return None
    if typ == "str":
        return str(x)
    if typ == "bool":
        # 不建议用bool，这里给兜底
        if isinstance(x, str):
            return normalize_yes_no(x) == "yes"
        return bool(x)
    return x

def postprocess_types_and_rounding(rows, field_schema):
    """
    field_schema: constraint_spec["field_schema"]
    for each field:
      - if type=str and allowed_values includes yes/no => normalize to yes/no
      - if numeric => clip to range, round decimals, cast type
    """
    for r in rows:
        for f, meta in field_schema.items():
            if f not in r:
                continue
            typ = meta.get("type")
            decimals = meta.get("decimals", None)
            rng = meta.get("range") or {}
            mn, mx = rng.get("min"), rng.get("max")
            allowed = meta.get("allowed_values")

            # yes/no normalization for original binary fields
            if typ == "str" and isinstance(allowed, list) and set([a.lower() for a in allowed]) == {"yes","no"}:
                r[f] = normalize_yes_no(r.get(f))
                continue

            # numeric handling
            if typ in {"int","float"}:
                v = r.get(f)
                v = clip_value(v, mn, mx)
                v = cast_value(v, typ, decimals=decimals)
                r[f] = v

    return rows