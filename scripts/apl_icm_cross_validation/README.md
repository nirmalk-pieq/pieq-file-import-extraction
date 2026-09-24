# APL & ICM Cross Validation Script

This folder contains a script designed to cross-validate policy data directly between APL Statements and ICM (Parallel Run) payables.

## Script: `validate_apl_icm.py`

### What it does
The script compares policy, agent name, and commission amounts to ensure that what is recorded in the APL system matches the data on the ICM side.

It performs validation by bridging the datasets via a direct match:
1. **APL -> Parallel Run (ICM)**: The raw APL policy numbers are cleaned (stripping the leading `'` and removing any `OVR` prefix/suffix) and matched directly against the `policy_number` column in the Parallel Run Payables CSV to extract the ICM `Agent Name` and `commission_amount`.

### Validation Checks
- **Agent Name Matching**: Compares the APL `Writing Agent Name` against the ICM `Agent Name` (case-insensitive).
- **Amount Matching**: Compares the APL `Earned Commission` against the ICM `commission_amount` (strips `$`, `,` and allows small float variances).

### Column Configurations
The top of the script contains `Enum` blocks which define exactly which columns to use from the CSVs. 
If you ever need to use `Gross Commission` instead of `Earned Commission`, you can easily change it there without editing the core logic!

### Execution
Run the script natively using Python. It will interactively prompt you for the paths to the **two** input files:
```bash
cd apl_icm_cross_validation
python3 validate_apl_icm.py
```

### Outputs & Versioning
When the script runs, it automatically creates a new versioned folder (e.g., `versions/v1/`, `versions/v2/`) within the `apl_icm_cross_validation` directory.

Inside the version folder, you will find:
- **`input/`**: Contains exact copies of the input files you provided.
- **`output/`**: Contains the generated side-by-side validation results and error logs in both formats:
  - `Validation_Output.csv`
  - `Validation_Output.xlsx`
