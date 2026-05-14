
from collections import Counter
import numpy as np

def mean_sd(arr):
    arr = np.array(arr, dtype=float)
    return float(np.mean(arr)), float(np.std(arr, ddof=0))

def validate(spec, dataset,log):
    groups = spec["study"]["groups"]
    group_var = spec["study"].get("group_var","gos_group")
    mean_tol = spec.get("tolerances", {}).get("mean_tol", 0.2)
    sd_tol = spec.get("tolerances", {}).get("sd_tol", 0.3)

    report = {
        "hard": {"passed": True, "checks": []},
        "soft": {"passed": True, "checks": []},
        "notes": []
    }

    # hard: N & group counts
    N = spec["study"]["N"]
    if len(dataset) != N:
        report["hard"]["passed"] = False
        report["hard"]["checks"].append({"item":"N", "target":N, "actual":len(dataset), "passed":False})
    else:
        report["hard"]["checks"].append({"item":"N", "target":N, "actual":len(dataset), "passed":True})

    # group counts
    c = Counter([r.get(group_var) for r in dataset])
    for gname, gn in groups.items():
        ok = (c.get(gname,0) == gn)
        report["hard"]["passed"] &= ok
        report["hard"]["checks"].append({"item":f"group_count:{gname}", "target":gn, "actual":c.get(gname,0), "passed":ok})
        if not ok:
            log.append(f'item:group_count:{gname}, target:{gn}, actual:{c.get(gname,0)}, passed:Failed')

    # categorical by group counts
    cat = spec.get("categorical", {})
    for var, meta in cat.items():
        cbg = (meta or {}).get("counts_by_group")
        if not (isinstance(cbg, dict) and all(isinstance(cbg.get(g), dict) for g in groups.keys())):
            continue
        for gname in groups.keys():
            target = cbg[gname]
            actual = Counter([r.get(var) for r in dataset if r.get(group_var)==gname])
            for level, tv in target.items():
                ok = (actual.get(level,0) == tv)
                report["hard"]["passed"] &= ok
                report["hard"]["checks"].append({
                    "item": f"{var}:{gname}:{level}",
                    "target": tv,
                    "actual": actual.get(level,0),
                    "passed": ok
                })
                if not ok:
                    log.append(f'item:{var}:{gname}:{level}:{gname}, target:{tv}, actual:{actual.get(level,0)}, passed:Failed')

    # soft: continuous mean/sd by group
    cont = spec.get("continuous", {})
    for var, meta in cont.items():
        by_group = (meta or {}).get("by_group") or {}
        for gname in groups.keys():
            gstats = by_group.get(gname) or {}
            tm, ts = gstats.get("mean"), gstats.get("sd")
            if tm is None or ts is None:
                continue  # missing => no soft check
            vals = [r.get(var) for r in dataset if r.get(group_var)==gname and r.get(var) is not None]
            if len(vals) == 0:
                report["soft"]["passed"] = False
                report["soft"]["checks"].append({"item":f"{var}:{gname}", "passed":False, "reason":"no values"})
                continue
            am, astd = mean_sd(vals)
            ok_m = abs(am - tm) <= mean_tol
            ok_s = abs(astd - ts) <= sd_tol
            ok = ok_m and ok_s
            report["soft"]["passed"] &= ok
            report["soft"]["checks"].append({
                "item": f"{var}:{gname}",
                "target": {"mean":tm, "sd":ts},
                "actual": {"mean":am, "sd":astd},
                "passed": ok,
                "mean_diff": am-tm,
                "sd_diff": astd-ts
            })
            if not ok:
                log.append(f'item:{var}:{gname}, target:mean:{tm}、sd:{ts}, actual:mean:{am}、sd:{astd},mean_diff:{am-tm},sd_diff:{astd-ts}, passed:Failed')


    return report

