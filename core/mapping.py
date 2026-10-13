import os
import pandas as pd


def load_mapping_file(map_path):
    """
    Loads mapping data from a CSV or Excel file.
    Expects 3 columns: [Subject ID, Genotype, Sex]

    Returns:
        mapping_dict: dict mapping ID substring -> (genotype, sex)
        genotype_index: dict mapping normalized genotype -> integer index
        unique_genotypes: list of unique genotype strings
    """
    if not map_path or not os.path.isfile(map_path):
        return {}, {}, []

    if map_path.lower().endswith('.csv'):
        df_map = pd.read_csv(map_path)
    else:
        df_map = pd.read_excel(map_path)

    if df_map.shape[1] < 3:
        raise ValueError("Mapping file must contain at least 3 columns: [ID, Genotype, Sex]")

    # Normalize genotypes for color index assignment
    raw_genotypes = df_map.iloc[:, 1].dropna().astype(str).str.strip()
    unique_genotypes = list(dict.fromkeys(raw_genotypes))

    genotype_index = {geno.lower(): i for i, geno in enumerate(unique_genotypes)}

    mapping_dict = {}
    for _, row in df_map.iterrows():
        key = str(row.iloc[0]).strip()
        geno = str(row.iloc[1]).strip()
        sex = str(row.iloc[2]).strip()
        if key:
            mapping_dict[key] = (geno, sex)

    return mapping_dict, genotype_index, unique_genotypes


def match_table_id(table_id, mapping_dict):
    """
    Attempts to match a Table ID against keys in the mapping dictionary.
    """
    table_id_str = str(table_id).strip()
    matched_key = next((k for k in mapping_dict if k in table_id_str), None)
    if matched_key:
        return mapping_dict[matched_key]
    return None, None
