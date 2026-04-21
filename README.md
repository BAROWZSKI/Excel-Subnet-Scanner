
###### <p align="center"> *This is official repository maintained by me*</center> </p>

# Excel Subnet Scanner

A VBA macro that scans subnets defined in an Excel sheet, checks IP availability via ping, resolves hostnames via DNS, and writes results back — all unattended. Designed to run silently on a Windows server via Task Scheduler.

![ExampleExcel](/img.png "example")

---

## How It Works

- Reads subnet list from a sheet named **Ozet**
- For each subnet, creates (or overwrites) a dedicated sheet
- Pings every IP in the range and marks it as `Used` or `Free`
- Resolves hostnames via `nslookup` and `ping -a` fallback
- Preserves manually entered data (responsible person, notes, etc.) across runs
- Calculates occupancy rate and writes it back to the Ozet sheet
- Saves a timestamped backup before each run
- Logs every step to a `.txt` file

---

## File Structure

After first run, the following folders are created automatically next to the `.xlsm` file:

```
YourFile.xlsm
├── Backups\
│   └── 12_04_2025_23-59-00.xlsm
└── Logbook\
    └── 12_04_2025.txt
```

---

## Setup

### 1. Prepare the Excel File

- Enable macros when opening the file
- Import `SubnetScanner.bas` into a standard VBA Module:
  - Press `Alt + F11` to open the VBA editor
  - Go to **Insert → Module**
  - Copy-paste the contents of `SubnetScanner.bas` into the module

### 2. Create the Ozet Sheet

The macro reads from a sheet named exactly **Ozet**. The sheet must have the following columns:

| Column | Content | Example |
|--------|---------|---------|
| A | Subnet name | `Office-LAN` |
| B | CIDR range | `192.168.1.0/24` |
| C | Occupancy (auto-filled) | *(leave empty)* |
| D | Gateway IP | `192.168.1.1` |
| E | Firewall name | `FW-01` |

Row 1 is the header row. Data starts from row 2.

Example:

```
Name          Range               Occupancy   Gateway         Firewall
Office-LAN    192.168.1.0/24                  192.168.1.1     FW-Core
Server-VLAN   10.10.0.0/28                    10.10.0.1       FW-Core
```

### 3. Run Manually (for testing)

- Set `DEBUG_MODE = True` at the top of the module
- Open the Excel file
- Press `Alt + F8`, select `CreateAndScanSubnets`, click **Run**
- A popup will appear when done

### 4. Deploy to Server (unattended)

- Set `DEBUG_MODE = False` in the macro before deploying
- Edit `Trigger.ps1` and update these two lines:

```powershell
$excelFilePath = "C:\Path\To\Your\File.xlsm"
$logDir        = "C:\Path\To\Your\Logbook"
```

---

## Task Scheduler Setup

1. Open **Task Scheduler** as Administrator
2. Click **Create Basic Task**
3. Give it a name (e.g. `SubnetScanner-Weekly`)
4. Trigger: **Weekly → Saturday → 00:00**
5. Action: **Start a program**

Fill in as follows:

```
Program:   powershell.exe
Arguments: -ExecutionPolicy Bypass -NonInteractive -File "C:\Path\To\Trigger.ps1"
Start in:  C:\Path\To\   ← folder containing Trigger.ps1, no quotes
```

6. After creating, **double-click the task** → go to the **General** tab:
   - Select **"Run whether user is logged on or not"**
   - Check **"Run with highest privileges"**
   - Enter the server account credentials when prompted

7. To test: right-click the task → **Run**. Check the Logbook folder to verify it worked.

---

## Notes

- The macro preserves manually entered data in columns D, E, F, G, and I across weekly runs
- Column H (Firewall) is always sourced from the Ozet sheet — manual edits to it will be overwritten
- The scan duration depends on subnet size and network latency. 50+ mixed subnets can take 4-6 hours
- `DEBUG_MODE = True` shows popups on errors — useful for local testing, must be `False` on server
