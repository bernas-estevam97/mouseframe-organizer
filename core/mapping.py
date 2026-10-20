import os
import re
import pandas as pd

_SPLIT_RE = re.compile(r"[_\s\-]+")


def load_mapping_file(map_path):
    """
    Loads mapping data from a CSV or Excel file.
    Expects a header row and 3 columns: [Subject ID, Genotype, Sex]

    Returns:
        mapping_dict: dict mapping subject ID -> (genotype, sex)
        genotype_index: dict mapping normalized genotype -> integer index
    """
    if not map_path:
        return {}, {}
    if not os.path.isfile(map_path):
        raise FileNotFoundError(f"Mapping file not found: {map_path}")

    # dtype=str keeps IDs like '0123' intact and avoids '2199.0' when the column has blanks
    if map_path.lower().endswith(".csv"):
        df_map = pd.read_csv(map_path, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
    else:
        df_map = pd.read_excel(map_path, dtype=str)

    if df_map.shape[1] < 3:
        raise ValueError("Mapping file must contain at least 3 columns: [ID, Genotype, Sex]")

    df_map = df_map.iloc[:, :3].fillna("").apply(lambda col: col.str.strip())
    df_map = df_map[df_map.iloc[:, 0] != ""]
    if df_map.empty:
        raise ValueError("Mapping file contains no subject IDs.")

    mapping_dict = {}
    conflicts = []
    for key, geno, sex in df_map.itertuples(index=False, name=None):
        if key in mapping_dict and mapping_dict[key] != (geno, sex):
            conflicts.append(key)
        mapping_dict[key] = (geno, sex)
    if conflicts:
        raise ValueError(
            "Mapping file lists these IDs more than once with different Genotype/Sex:\n"
            + ", ".join(sorted(set(conflicts)))
        )

    unique_genotypes = dict.fromkeys(g for g, _ in mapping_dict.values() if g)
    genotype_index = {geno.lower(): i for i, geno in enumerate(unique_genotypes)}

    return mapping_dict, genotype_index


def match_key(table_id, mapping_dict):
    """
    Returns the mapping key that matches a Table ID, or None.

    A key matches when its tokens (split on '_', '-', spaces) are the leading tokens of the
    Table ID, compared case-insensitively. The longest matching key wins, so '2199_1R' beats
    '2199', and '219' never matches '2199'.
    """
    id_tokens = [t.lower() for t in _SPLIT_RE.split(str(table_id).strip()) if t]
    best, best_len = None, 0
    for key in mapping_dict:
        key_tokens = [t.lower() for t in _SPLIT_RE.split(key) if t]
        n = len(key_tokens)
        if n > best_len and id_tokens[:n] == key_tokens:
            best, best_len = key, n
    return best

