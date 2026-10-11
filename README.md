# Weheat CSV export (PowerShell)

Export data from your [Weheat](https://www.weheat.nl) heat pump to CSV with a single
PowerShell script. Nothing to install: it runs on the PowerShell that comes with Windows
(5.1) and on PowerShell 7 (Windows, macOS, Linux).

It talks to the same API as the official Python client,
[wefabricate/wh-python](https://github.com/wefabricate/wh-python): it logs in with your
Weheat portal account and calls the `third_party` REST API.

## What you need

- A **client ID and client secret** for the Weheat API. You get these from Weheat
  (the same ones wh-python needs).
- The **username (e-mail) and password** you use for the Weheat online portal.

## 1. Download

**Option A: ZIP (no Git needed)**

1. On this page, click the green **Code** button, then **Download ZIP**.
2. Unzip it, for example to `C:\weheat`.
3. Windows marks downloaded scripts as "from the internet". Unblock them once: open
   PowerShell in that folder and run
   ```powershell
   Get-ChildItem *.ps1 | Unblock-File
   ```

**Option B: only the script, from PowerShell**

```powershell
mkdir C:\weheat; cd C:\weheat
$base = 'https://raw.githubusercontent.com/BoukeStoelinga/weheat-csv-export/main'
Invoke-WebRequest "$base/weheat-export.ps1" -OutFile weheat-export.ps1
Invoke-WebRequest "$base/weheat-credentials.example.txt" -OutFile weheat-credentials.txt
```

**Option C: Git**

```powershell
git clone https://github.com/BoukeStoelinga/weheat-csv-export.git
cd weheat-csv-export
```

## 2. Add your credentials

Copy `weheat-credentials.example.txt` to `weheat-credentials.txt` (Option B already did
this) in the same folder as the script, and fill it in:

```text
ClientId     = your-client-id
ClientSecret = your-client-secret
Username     = you@example.com
Password     = your-portal-password
HeatPumpId   =
```

- Leave `Password` empty to be asked for it every run.
- Leave `HeatPumpId` empty to use the first active heat pump on your account
  (run with `-Data HeatPumps` to see all ids).
- The file is plain text: keep it private. `.gitignore` keeps it out of Git.

Instead of the file you can also pass `-ClientId`, `-ClientSecret`, `-Username`,
`-Password`, `-HeatPumpId`, or set the environment variables `WEHEAT_CLIENT_ID`,
`WEHEAT_CLIENT_SECRET`, `WEHEAT_USERNAME`, `WEHEAT_PASSWORD`, `WEHEAT_HEAT_PUMP_ID`.
Command-line parameters win over the file, the file wins over environment variables.
Use `-CredentialFile <path>` to keep the file somewhere else.

## 3. Run

Open PowerShell in the script's folder (in Explorer: type `powershell` in the address
bar and press Enter) and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\weheat-export.ps1 -Data Logs -Interval Hour -OutFile logs.csv
```

`-ExecutionPolicy Bypass` only applies to this one run; it's needed because Windows
blocks unsigned scripts by default. If you've allowed scripts already
(`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`), `.\weheat-export.ps1 ...`
is enough.

### Examples

```powershell
# Your heat pumps (ids, names, models)
.\weheat-export.ps1 -Data HeatPumps

# Latest status
.\weheat-export.ps1 -Data Latest -OutFile latest.csv

# 15-minute logs for the last 7 days
.\weheat-export.ps1 -Data Logs -Interval FifteenMinute -Start (Get-Date).AddDays(-7) -OutFile logs.csv

# Daily energy for 2025, with ; as separator for Dutch/European Excel
.\weheat-export.ps1 -Data Energy -Interval Day -Start 2025-01-01 -End 2026-01-01 -Delimiter ';' -OutFile energy-2025.csv

# Every raw log record of yesterday
.\weheat-export.ps1 -Data RawLogs -Start (Get-Date).Date.AddDays(-1) -End (Get-Date).Date -OutFile raw.csv
```

Without `-OutFile` the CSV is printed to the console.

## Options

| Parameter | Default | Meaning |
|---|---|---|
| `-Data` | `Logs` | `Logs` (aggregated per interval), `RawLogs` (every record), `Latest`, `Energy` (per interval), `EnergyTotal` (lifetime totals), `HeatPumps` |
| `-Interval` | `Hour` | `Minute`, `FiveMinute`, `FifteenMinute`, `Hour`, `Day`, `Week`, `Month`, `Year` (for `Logs` and `Energy`) |
| `-Start`, `-End` | last 24 hours | Time range, in local time. Dates like `2025-01-01` work |
| `-OutFile` | console | CSV file to write (UTF-8) |
| `-Delimiter` | `,` | CSV separator; use `';'` for Excel with Dutch/European settings |
| `-CredentialFile` | `weheat-credentials.txt` next to the script | Credentials file |
| `-HeatPumpId` | first active heat pump | Which heat pump |

Long time ranges are split into several requests automatically, because the API limits
how much one request may cover (for example 2 days at `Minute`, a month at `Hour`).
Nested fields become `parent.child` columns; timestamps are UTC in ISO 8601.

Add `-Verbose` to see every API request.

## Troubleshooting

- **"running scripts is disabled on this system"**: use the
  `powershell -ExecutionPolicy Bypass -File ...` form above.
- **`invalid_client`**: client ID or secret is wrong.
- **`invalid_grant` / "Invalid user credentials"**: username or password is wrong.
- **403 / 404 on a heat pump**: the heat pump id isn't on your account; check with
  `-Data HeatPumps`.
- **Excel puts everything in one column**: rerun with `-Delimiter ';'`.

## Disclaimer

Not affiliated with Weheat. Uses the API as documented by Weheat's own
[wh-python](https://github.com/wefabricate/wh-python) client. MIT licensed.
