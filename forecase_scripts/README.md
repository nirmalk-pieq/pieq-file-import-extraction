# Review-to-reconcile file forecast

Forecasts how one commission statement would move through file reconciliation and commission processing. The script reads the source database and writes the result on your machine. It does not update that database, and it does not POST or PUT to the API base URL.

## Setup

From this folder, pass the environment name. That loads `config_preprod` or `config_prod`:

```bash
pip install -r requirements.txt
python3 review_to_reconcile_file_forecast.py preprod <commission_statement_uuid> [<uuid> ...]
python3 review_to_reconcile_file_forecast.py prod review --agency_id 5
```

`config_preprod` and `config_prod` use the keys in `.env.example`. Values already set in the shell override the file.

Fill in `config_preprod` or `config_prod`:

| Variable | Purpose |
| --- | --- |
| `PIEQ_API_BASE_URL` | Policy API base URL. The script does not send POST or PUT here. |
| `AUTH_URL` | Keycloak host. The token is requested at `{AUTH_URL}/realms/pieq-sso/protocol/openid-connect/token`. |
| `client_id` | Keycloak client id. Sample value `pieq-app`. |
| `USERNAME` | SSO username. |
| `PASSWORD` | SSO password. |
| `grant_type` | OAuth grant. Sample value `password`. |
| `PIEQ_DB_HOST` | Database host. Preprod is `preprod-mlb-cluster.cluster-cshu8yiygpyj.us-east-1.rds.amazonaws.com`. |
| `PIEQ_DB_PORT` | Database port. Default `5432`. |
| `PIEQ_DB_NAME` | Database name. |
| `PIEQ_DB_USERNAME` | Database username. |
| `PIEQ_DB_PASSWORD` | Database password. |
| `APL_STATEMENT_PATH` | Optional. One carrier statement CSV or Excel compared with `commission_detail_report.csv`. |
| `APL_STATEMENT_DIR` | Optional. Folder of carrier statements. A file is used when its name matches the commission statement name. |
| `APL_VALIDATION_MODE` | Optional. APL cross-check driving side: `apl` (default, APL -> ICM) or `icm` (ICM -> APL). Forecast always passes this non-interactively. |

## Run

```bash
python3 review_to_reconcile_file_forecast.py preprod <commission_statement_uuid> [<uuid> ...]
python3 review_to_reconcile_file_forecast.py prod <uuid-1> <uuid-2>
python3 review_to_reconcile_file_forecast.py prod review --agency_id 5
```

`review --agency_id <id>` forecasts every `commission_statements` row for that agency whose `stage` is `3` (Review). Each file is still written to its own `forecast-out/<uuid> - <statement name>/` folder. Pass one or more statement UUIDs to forecast only those files. Commas between UUIDs in one argument are also accepted. A failure on one file is logged and the remaining files still run. The process exits with code 1 when any file fails.

Logs go to stderr. A successful database login looks like:

```text
13:20:01 INFO Connecting to database (host=db.example port=5432 db=icm user=readonly_user), read only
13:20:01 INFO Database connected: db=icm user=readonly_user server=10.0.0.5:5432 driver=psycopg read_only=on
```

If the host, password, or network is wrong, the log line starts with `Database connection failed` and the script exits with code 1. Passwords and tokens are not written to the log.

Values already set in the shell override the same keys in `config_preprod` or `config_prod`.

## Output

Files are written to `forecast-out/<commission_statement_uuid>/`:

- `forecast.sql` — lookup rows used by the report joins, then the rows the forecast would insert
- `reconciliation_exceptions.md` — statement records left in exception
- `commission_detail_report.csv`
- `apl_icm_validation/` — APL statement compared with `commission_detail_report.csv` when `APL_STATEMENT_PATH` or `APL_STATEMENT_DIR` is set
- `ledger_summary.csv`
- `agent_balances.csv`
- `forecast.log` — one line per statement row, plus the connection and summary lines
- `forecast.md` — markdown summary of the same run (counts, exception reasons, row table)
- `forecast.xlsx` — the same summary in Excel (Summary, Exception codes, Exception rows, New policies, Row results)
- `commission_split_report.csv`
- `summary.json`

The split report lists the policies this forecast processed. `file_status` is `0` when every record cleared and `1` when any record is still in exception. The published report query only returns a file after `commission_statements.status` is `0`, so those exception rows are listed in `reconciliation_exceptions.md`.

A policy that does not already exist is created only in the forecast output. The script reads match-commission, the pay-to agent's reporting hierarchy, and `GET /api/v1/agencies/hierarchy`, then attaches commission splits the same way as fileprocessing: map match-commission `level_order` to agency level UUIDs, build the pay-to → manager chain, pick role amounts by the pay-to agent's grade (or the manager's grade when pay-to is a sub-agent), and seat Agency / Sales Leader / Agent / Sub Agent from that chain. Local policy and commission slabs are stored in the forecast output only; nothing is posted to the API and nothing is written to the database. New-policy agent resolve uses NPN, then appointment number, then user code, then a unique contacts full-name match (`entity = 'agent'`). If create fails, the row stays a `POLICY_MISMATCH` with an explicit reason tag: `reason=agent`, `reason=product`, `reason=effective_date`, or `reason=commission`. When the product `commission_processing_type` is FMV (`1`), the forecast uses the same amount shapes as file processing: Regular, Renewal, prorated Renewal for new business (`NB`), and Regular minus prorated Renewal when a prior renewal true-up payment already exists. Other products still match only the slab that covers the policy month.

## Ledger validation

Compare APL payables vs ICM (agent flow) ledger CSVs by agent name. Inputs are only `--apl` and `--icm` (alias `--latest`). Joins by normalized agent name (collapse spaces, case-insensitive). No Agent ID join. For each APL row, uses `Earned Commission` when non-zero, otherwise `Earned Credit`. Policy numbers strip `'` and flag `OVR` suffixes. An agent is matched when both files have the name and totals agree within $0.01.

```bash
python3 validate_ledger_apl_vs_agent_flow.py \
  --apl "/path/PR_9_APL_PAYABLES.CSV" \
  --icm "/path/PR_9_Latest.csv" \
  --out "forecast-out/ledger_validation"
```

Writes `forecast-out/ledger_validation/ledger_validation.xlsx` (Summary, Comparison with `Matched`, Failures) and `ledger_validation_failures.md`.
