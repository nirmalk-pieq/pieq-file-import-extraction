# APL Commission & ICM Commission Database Validation Engine

This package provides a standalone 2-stage validation and database verification engine for APL policy commission upload files against the ICM database (specifically scoped to `agency_id = 5`).

---

## 🚀 Key Features & 2-Stage Process

```
                      [ Raw APL Commission File / Folder ]
                                        │
                                        ▼
    ┌───────────────────────────────────────────────────────────────────────┐
    │                       STAGE 1: FILE VALIDATION                        │
    ├───────────────────────────────────────────────────────────────────────┤
    │  1. Product Normalization & Agent ID updates                          │
    │  2. Sequence Differential Rate / Fee Roll-ups                         │
    │  3. 100.00% Split Share Normalization (Percentage & Fixed Fee)        │
    │  4. Group Split Share Sum Verification (verifies sum == 100.00%)      │
    │  5. Policy Row Count Verification (verifies 3 to 4 rows per policy)   │
    │  6. Carrier Rate Benchmark Verification (Master Sheet Audit)          │
    └───────────────────────────────────────────────────────────────────────┘
                                        │
                     ┌──────────────────┴──────────────────┐
                     ▼                                     ▼
      [ file_validated_clean_<ts>.xlsx ]    [ file_validated_errors_<ts>.xlsx ]
                     │
                     ▼
    ┌───────────────────────────────────────────────────────────────────────┐
    │                    STAGE 2: DATABASE VERIFICATION                     │
    ├───────────────────────────────────────────────────────────────────────┤
    │  1. Query products table (agency_id=5, name) -> product_uuid          │
    │  2. Query policies table (agency_id=5, number, product_id)            │
    │  3. Query policy_commission table -> verify commission rules          │
    └───────────────────────────────────────────────────────────────────────┘
                                        │
                     ┌──────────────────┴──────────────────┐
                     ▼                                     ▼
       [ db_matched_commission_<ts>.xlsx ]  [ db_unmatched_commission_<ts>.xlsx ]
```

---

## 📁 Directory Structure

```
apl_commission_icm_commission_db_validation/
├── .env                                  # Local DB and environment credentials
├── .env.example                          # Environment template
├── README.md                             # Documentation
└── validate_policy_commission_db.py     # Main 2-stage validation script
```

---

## ⚙️ Virtual Environment & Setup

1. **Create and Activate Virtual Environment**:
   ```bash
   # Navigate to folder
   cd apl_commission_icm_commission_db_validation

   # Create virtual environment
   python3 -m venv venv

   # Activate virtual environment
   # On macOS/Linux:
   source venv/bin/activate
   # On Windows:
   # venv\Scripts\activate
   ```

2. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure `.env`**:
   Ensure `.env` in this directory contains your database credentials:
   ```env
   # Database Credentials
   DATABASE_URL=postgresql://localhost:5432/pieq_db
   DB_USER=your_db_username
   DB_PASSWORD=your_db_password
   AGENCY_ID=5

   # Input Upload File
   UPLOAD_FILE_PATH=/path/to/commission_upload.csv

   # Optional Master Rate Sheet, Product Normalization Map & Agent Roster
   RATE_SHEET_PATH=ProductWise Commission Rate.xlsx
   NORM_FILE_PATH=Final Product Normalization Local.xlsx
   AGENT_ROSTER_PATH=../(AUG 3)RAW MASTER DATA for POLICY COMMISSIONS IMPORT/Agent ID, Agent Name Roaster Final.xlsx
   ```

---

## 💻 Usage

### 1. Run in Dry-Run Mode (File Validation Only)
```bash
python validate_policy_commission_db.py --file /path/to/commission_upload.csv --dry-run
```

### 2. Run Full 2-Stage DB Verification
```bash
python validate_policy_commission_db.py --file /path/to/commission_upload.csv
```

### 3. Specify Target Directory
```bash
python validate_policy_commission_db.py --file /path/to/directory_with_csvs/
```

---

## 📊 Generated Reports

| Phase | Output Report | Description |
|-------|---------------|-------------|
| **Phase 1** | `file_validated_clean_<timestamp>.xlsx` | Clean records passing split normalization and carrier benchmark validation. |
| **Phase 1** | `file_validated_errors_<timestamp>.xlsx` | Records failing Stage 1 with explicit line error messages. |
| **Phase 2** | `db_matched_commission_<timestamp>.xlsx` | Records matching DB products, policies, and commission rules for `agency_id = 5`. |
| **Phase 2** | `db_unmatched_commission_<timestamp>.xlsx` | Records failing DB lookup (Missing Product, Missing Policy, Commission Discrepancy). |

---

## 🛠️ Auto-Fixing Discrepancies (generate_commission_update_script.py)

If your Phase 2 Database Verification identifies missing or mismatched agents/agencies in the database `commission_split` JSON array, you can automatically generate an SQL repair script using the validation summary markdown file.

### Logic & Rules:
1. **Order 1 (Agency):** If the file agent is `"MAIN LINE BENEFITS"` or `"MLB parallel run"`, the discrepancy is considered expected and is **ignored**. If it's an unrecognized agency, it looks up the agency `uuid` using the `AGENCY_ID` in your `.env`.
2. **Order > 1 (Agent):** Looks up the user `uuid` in the database `users` table based on the `File Agent (Code)` and automatically repairs the JSON object for that specific order.

### How to use:
```bash
python generate_commission_update_script.py --file <path_to_validation_summary.md>
```

**Output:**
This will generate an `update_commission_splits.sql` file containing the necessary `UPDATE` statements to fix the JSON properties in your database.
