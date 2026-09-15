# clinical_data_curation

Curates raw PPMI clinical data into the analysis-ready tables used by `MRI_neurotransmitter_analysis/`. Run the notebooks in order:

1. **`01_load_and_merge_raw_data.ipynb`** — loads the raw PPMI and HIVE data cuts, fixes column types, resolves duplicate subject IDs, and merges the two sources.
2. **`02_clean_and_finalize_dataset.ipynb`** — loads the merged dataset, handles missing demographic/clinical values, resolves cohort/subgroup assignments, and removes outliers.
3. **`03_final_qc_and_exploration.ipynb`** — loads the cleaned dataset and produces sanity-check plots (subgroup sizes, missing-value heatmaps, clinical score distributions).

No PPMI data is included in this repository — see the top-level [README](../README.md#data--reproducibility) for how to obtain your own access and where to place the files.
