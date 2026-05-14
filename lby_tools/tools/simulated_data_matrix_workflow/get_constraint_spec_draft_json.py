def default_ranges():
    # 可按你数据字典扩展
    return {
        "age_years": [18, 100],
        "length_of_stay_days": [1, 90],
        "sbp_mmhg": [80, 260],
        "dbp_mmhg": [40, 160],
        "hematoma_volume_ml": [1, 200],
        "gcs_score": [3, 15],
        "glucose_onset_mmol_L": [1.5, 30],
        "fasting_blood_glucose_mmol_L": [1.5, 30],
        "uric_acid_umol_L": [50, 800],
        "total_cholesterol_mmol_L": [1.5, 15],
        "triglycerides_mmol_L": [0.1, 15],
        "serum_albumin_g_L": [20, 55],
    }

def ensure_groups(spec):
    N = spec.get("study", {}).get("N")
    groups = spec.get("study", {}).get("groups", {})
    if not groups or not isinstance(groups, dict):
        # fallback if missing
        spec.setdefault("study", {})
        spec["study"]["groups"] = {"low": None, "high": None}
    # try infer N
    if N is None:
        gl = spec["study"]["groups"].get("low")
        gh = spec["study"]["groups"].get("high")
        if isinstance(gl, int) and isinstance(gh, int):
            spec["study"]["N"] = gl + gh
    return spec

def attach_tolerances(spec, mean_tol=0.2, sd_tol=0.3):
    spec.setdefault("tolerances", {})
    spec["tolerances"]["mean_tol"] = float(mean_tol)
    spec["tolerances"]["sd_tol"] = float(sd_tol)
    return spec

def finalize_spec(spec, mean_tol=0.2, sd_tol=0.3):
    spec = ensure_groups(spec)
    spec = attach_tolerances(spec, mean_tol, sd_tol)
    spec.setdefault("ranges", {})
    # merge default ranges without overwriting user-provided
    dr = default_ranges()
    for k,v in dr.items():
        spec["ranges"].setdefault(k, v)
    # basic required categorical defaults
    spec.setdefault("categorical", {})
    # if sex counts missing but paper likely has, keep as is; generator will use what exists
    spec.setdefault("warnings", [])
    return spec

