# Linking MRI Atrophy Patterns in Parkinson's Disease to Neurotransmitter Systems using NiSpace

MSc thesis project investigating structural brain atrophy in Parkinson's disease (PD) — from prodromal stages through de novo diagnosis — and how those atrophy patterns spatially align with neurotransmitter system maps, using [NiSpace](https://github.com/LeonDLotter/NiSpace). Clinical and imaging data are drawn from the [Parkinson's Progression Markers Initiative (PPMI)](https://www.ppmi-info.org).

## Repository structure

```
.
├── clinical_data_curation/       # Curates and merges raw PPMI clinical data into analysis-ready tables
└── MRI_neurotransmitter_analysis/ # Volume/thickness group comparisons, clinical correlations,
                                    # ggseg surface visualization, and NiSpace neurotransmitter colocalization
```

- **`clinical_data_curation/`** — Jupyter notebooks that load, clean, and merge the raw PPMI clinical data cut into the tables used downstream (subject diagnosis/cohort, demographics, cognitive and motor scores, biomarkers).
- **`MRI_neurotransmitter_analysis/`** — Imaging-derived analysis pipeline:
  - `notebooks/00_preprocessing` — outlier detection on volume/thickness measures
  - `notebooks/01_main_analysis` — case-control group comparisons (cortical/subcortical volume, cortical thickness)
  - `notebooks/02_clinical` — correlating atrophy with clinical variables
  - `notebooks/03_visualization` / `04_appendix` — figures and supplementary analyses, including brain-surface rendering with [ggseg](https://ggseg.github.io/ggseg/) and neurotransmitter colocalization via NiSpace
  - `scripts/` — helper scripts (e.g. compiling appendix result tables)

Most notebooks are Python; the `ggseg_*` notebooks under `03_visualization` and `04_appendix` run on an **R kernel (IRkernel)** for brain-surface plotting via the R `ggseg` package.

## Data & reproducibility

**No PPMI data — raw, merged, or derived — is included in this repository**, nor are any result tables or figures generated from it. This is due to the [PPMI Data Use Agreement](https://www.ppmi-info.org/access-data-specimens/download-data), which does not permit redistributing the data or data products outside the approved research use. Only analysis code and notebook structure (with all cell outputs stripped) are published here.

To reproduce this work:
1. Request your own PPMI data access at [ppmi-info.org](https://www.ppmi-info.org/access-data-specimens/download-data).
2. Place the relevant files under `clinical_data_curation/data/` and `MRI_neurotransmitter_analysis/data/` (both git-ignored) following the paths referenced in the notebooks.
3. Install dependencies and run the notebooks in the order implied by the folder numbering (`00_` → `04_`).

## Setup

**Python** — each subproject has its own `requirements.txt`:

```bash
# clinical_data_curation
cd clinical_data_curation
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

# MRI_neurotransmitter_analysis
cd MRI_neurotransmitter_analysis
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
```

**R** — the `ggseg_*` notebooks additionally require an R installation with [IRkernel](https://irkernel.github.io/) and the [`ggseg`](https://ggseg.github.io/ggseg/) / [`ggsegExtra`](https://github.com/ggseg/ggsegExtra) packages for cortical/subcortical surface plotting.

Notebook outputs are stripped automatically on commit via [`nbstripout`](https://github.com/kynan/nbstripout), configured as a git filter for this repository, so patient-level output cells never enter version control even if a notebook is committed with outputs still attached locally.
