# Extraction Patterns

How to reliably parse `ServerInfo_*.txt` files produced by `Export-ServerInfo.ps1`.

---

## File Structure

```
============================================================
AsInstalled Server Export - HOSTNAME
Generated: YYYY-MM-DD HH:MM:SS
============================================================

[Section Name]
FieldName: Value
FieldName: Value

[Next Section]
...
============================================================
  Export complete. ...
============================================================
```

- **Section headers:** `^\[(.+)\]$`
- **Field lines:** `^(\w[\w &\/\(\)\-]+): (.+)$` — note spaces allowed in field names
- **Queue separators:** `  ============================================`
- **Queue subsection headers:** `  -- (.+) --`

---

## Section Index

Extract section blocks with:
```
\[Section Name\]([\s\S]*?)(?=\[|\Z)
```

Known sections produced by current script version:

| Section header | Content |
|----------------|---------|
| `[System Identity]` | hostname, VM, BIOS, manufacturer |
| `[Operating System]` | OS name, version, arch |
| `[Hardware]` | RAM, CPU, drives |
| `[Network]` | per-NIC blocks |
| `[Domain]` | AD info |
| `[Installed Software]` | app detection results |
| `[Roles and Features]` | Windows features |
| `[Database]` | SQL/PostgreSQL config |
| `[Print Server]` | spooler + all queues + driver + port list |
| `[Device Counts]` | FujiFilm/FX counts |

---

## Print Queue Blocks

Each queue in `[Print Server]` starts after `  ============================================` and ends at the next one. The first line in the block is `  Queue Name              : <name>`.

### Extracting all queues

```python
import re

def extract_queues(text):
    # Split on queue separator lines
    sep = r'  ={40,}'
    parts = re.split(sep, text)
    queues = []
    for part in parts:
        if re.search(r'Queue Name\s*:', part):
            queues.append(part.strip())
    return queues
```

### Extracting a field from within a queue block

```python
def queue_field(queue_text, field_name):
    pattern = rf'^\s+{re.escape(field_name)}\s*:\s*(.+)$'
    m = re.search(pattern, queue_text, re.MULTILINE)
    return m.group(1).strip() if m else None
```

### Key queue fields and their patterns

| Field | Pattern | Example value |
|-------|---------|---------------|
| Queue Name | `Queue Name\s*:\s*(.+)` | `FollowYou` |
| Driver | `Driver\s*:\s*(.+)` | `FF Multi-model Print Driver 2` |
| Port | `Port\s*:\s*(.+)` | `192.168.60.90` |
| Status | `Status\s*:\s*(.+)` | `Normal` |
| Paper Size | `Paper Size\s*:\s*(.+)` | `A4` |
| Duplex | `2-sided Print.*:\s*(.+)` | `Two-sided Long Edge (Portrait / Flip on Long)` |
| Output Color | `Output Color\s*:\s*(.+)` | `Color` |
| Staple (job default) | `Staple\s*:\s*(.+)` | `Unknown` |
| Use Application Color | `Use Application Color\s*:\s*(.+)` | `On  (passes application dmColor through; no driver override)` |
| ICM Method | `ICM Method\s*:\s*(.+)` | `ICM Disabled (use dmColor from application)` |
| Finisher Installed | `Finisher Installed\s*:\s*(.+)` | `Finisher B (GB_GB4), Finisher C (GC4_GC5)` |
| Staple Capability | `Staple Capability\s*:\s*(.+)` | `Yes (FreeStaple)` |
| Punch Capability | `Punch Capability\s*:\s*(.+)` | `Yes (2/3-hole, 2/4-hole)` |
| Booklet Maker | `Booklet Maker\s*:\s*(.+)` | `Booklet (BiFold, CZFold)` |
| Offset Stacking (hw) | `Offset Stacking\s*:\s*(.+)` | `Yes` |
| Shared | `Shared\s*:\s*(.+)` | `Yes (Share name: FollowYou)` |
| x64 Driver | `x64 Driver\s*:\s*(.+)` | `Installed (FF Multi-model Print Driver 2)` |
| x86 Driver | `x86 \(32-bit\) Driver\s*:\s*(.+)` | `Installed` |

---

## Field Value Interpretations

### Use Application Color

| Raw value starts with | Meaning |
|----------------------|---------|
| `On` | Toggle ON — application colour mode passes through unchanged |
| `Off (Windows system ICM` | Toggle OFF — Windows ICM active |
| `Off (driver manages` | Toggle OFF — driver ICM active |
| `Off (device hardware` | Toggle OFF — hardware ICM active |
| `Unknown` | Not in default PrintTicket (driver doesn't advertise it) |

### ICM Method

| Raw value | dmICMMethod value | Meaning |
|-----------|-------------------|---------|
| `ICM Disabled (use dmColor from application)` | 1 | No colour management — app controls |
| `ICM handled by OS` | 2 | Windows colour system manages conversion |
| `ICM handled by device` | 3 | Printer hardware manages ICC profiles |
| `ICM handled by host` | 4 | Host system (print server) manages |
| `Not available` | — | Registry read failed or DEVMODE < 192 bytes |

### Duplex

| Raw value | Meaning |
|-----------|---------|
| `One-sided (Simplex)` | Single-sided |
| `Two-sided Long Edge (Portrait / Flip on Long)` | Standard duplex / portrait |
| `Two-sided Short Edge (Landscape / Flip on Short)` | Landscape duplex / rotated |

### Installable Options — virtual driver

When the `-- Installable Options --` section contains:
```
  N/A (software/virtual driver - no hardware options)
```
This means `Get-PrinterProperty` returned no properties — the queue uses a virtual/software driver (e.g., FF Multi-model Print Driver 2 for follow-you printing) and has no physical finisher hardware directly attached. Do not interpret this as "no finisher capability" — the actual device may have full finishing hardware; the driver just doesn't expose it via this API.

---

## Installed Drivers Section

```
Installed Printer Drivers (6 x64, 2 x86):
  --- x64 (64-bit) ---
  Driver  : FF Apeos C7071 PCL 6
  Version : 3
  Print Processor : winprint
  ---
  ...
  --- x86 (32-bit) ---
  Driver  : FF Multi-model Print Driver 2
  Version : 3
  Print Processor : winprint
  ---
```

Pattern for x86 driver names:
```python
x86_block = re.search(r'--- x86 \(32-bit\) ---([\s\S]*?)(?=Configured|$)', text)
x86_drivers = re.findall(r'Driver\s*:\s*(.+)', x86_block.group(1)) if x86_block else []
```

---

## Device Counts Section

```
[Device Counts]
FujiFilm Devices: 2
FF Devices: 2
Fuji Xerox: 0
FX Devices: 0
```

Note: `FujiFilm Devices` and `FF Devices` are the same count (all FF/FujiFilm/Apeos queues). `Fuji Xerox` and `FX Devices` are a separate count for legacy Fuji Xerox branded queues.

---

## Network Section

Each NIC is a sub-block within `[Network]`:
```
  Adapter: Ethernet0
  MAC: 00-50-56-A3-1B-2C
  IP: 192.168.60.150
  Subnet: 255.255.255.0
  Gateway: 192.168.60.1
  DNS1: 192.168.60.1
  DHCP: Static
  Link Speed: 10000 Mbps
```

NIC blocks start at `Adapter:` and end at the next `Adapter:` or section boundary.

---

## Null / Missing Values

| Displayed as | Meaning |
|-------------|---------|
| `Unknown` | Query succeeded but value not in standard field (PrintTicket, registry) |
| `Not available` | Query failed or data missing |
| `Not installed` | Feature/driver checked and confirmed absent |
| `None detected` | Collection returned empty result |
| `N/A` | Not applicable for this queue/driver type |
| *(field absent)* | Section not reached (script error or OS not supported) |

---

## Validation Checklist

When ingesting a new export, verify:

1. File header line 1 contains `AsInstalled Server Export - ` + hostname
2. `[Print Server]` section exists and contains at least one queue
3. `[Device Counts]` section exists with numeric values
4. Each queue block has a `Driver` field — if missing, the queue entry is malformed
5. Queues with physical drivers should have `-- Installable Options --` section (not `N/A`)
6. `Spool Folder:` field appears in `[Print Server]` before the queue list
