import random

def expand_quota(level_counts):
    arr = []
    for k,v in level_counts.items():
        if v is None: 
            continue
        arr += [k]*int(v)
    return arr

# def assign_by_group(rows, group_var, group_name, var_name, counts_by_group):
#     # counts_by_group: {"male":39,"female":18}
#     idxs = [i for i,r in enumerate(rows) if r[group_var]==group_name]
#     quota = expand_quota(counts_by_group)
#     if len(quota) != len(idxs):
#         raise ValueError(f"Quota size mismatch for {var_name} in group {group_name}: quota={len(quota)} rows={len(idxs)}")
#     random.shuffle(quota)
#     for i,val in zip(idxs, quota):
#         rows[i][var_name] = val


def adjust_quota(quota, target_n,log):
    """将 quota 微调到 target_n：不够就随机补，超了就截断。"""
    if target_n <= 0:
        return []

    if len(quota) == 0:
        # 极端情况：没配额，随便给一个占位（也可以改成 None）
        return [None] * target_n

    if len(quota) < target_n:
        # 不够：从已有类别里随机抽样补齐
        add_or_del = random.choices(quota, k=target_n - len(quota))
        quota = quota + add_or_del
        loginfo = f'不够：从已有类别里随机抽样补齐，{add_or_del}'
    elif len(quota) > target_n:
        # 超了：随机打乱后截断（等价于随机删掉多余的）
        add_or_del = quota[target_n+1:]
        quota = quota[:target_n]
        random.shuffle(quota)
        loginfo = f'超了：随机打乱后截断（等价于随机删掉多余的）,{add_or_del}'
    log.append(loginfo)
    return quota

def assign_by_group(rows, group_var, group_name, var_name, counts_by_group,log):
    idxs = [i for i, r in enumerate(rows) if r[group_var] == group_name]
    quota = expand_quota(counts_by_group)

    # 核心修改：不匹配就微调，不报错终止
    if len(quota) != len(idxs):
        quota = adjust_quota(quota, len(idxs),log)

    random.shuffle(quota)
    for i, val in zip(idxs, quota):
        rows[i][var_name] = val

def make_patient_ids(N, prefix="SHICH"):
    return [f"{prefix}-{i:03d}" for i in range(1, N+1)]

def build_frame(spec, seed=42):
    random.seed(seed)
    N = spec["study"]["N"]
    groups = spec["study"]["groups"]
    group_var = spec["study"].get("group_var","gos_group")

    rows = [{"patient_id": pid} for pid in make_patient_ids(N, prefix="SHICH")]
    # assign groups
    quota = []
    for gname, gn in groups.items():
        quota += [gname]*int(gn)
    if len(quota) != N:
        raise ValueError(f"Group quota != N: {len(quota)} vs {N}")
    random.shuffle(quota)
    for r,g in zip(rows, quota):
        r[group_var] = g

    cat = spec.get("categorical", {})
    log = []
    # helper: assign if counts_by_group exists
    for var_name, meta in cat.items():
        cbg = meta.get("counts_by_group")
        if isinstance(cbg, dict) and all(isinstance(cbg.get(g), dict) for g in groups.keys()):
            for gname in groups.keys():
                assign_by_group(rows, group_var, gname, var_name, cbg[gname],log)

    return rows,log

