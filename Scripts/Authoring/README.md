# 🧙 Statblock Authoring System — Developer Notes

> This document summarizes the current Statblock authoring pipeline for the BRP creature toolkit.  
> It covers creature creation, cult/magic allocation, module structure, and data dependencies.

---

## 🗺️ Overview

The system converts a textual BRP creature entry into a complete statblock, decorates it with
cult data, and adds Spirit and Rune magic automatically.  
It is modular, data-driven, and designed for use inside the **`Y:\Stat_blocks\Scripts`** tree.

---

## ⚙️ Module Structure

| Path | Purpose |
|------|----------|
| `Authoring/New-Creature.ps1` | Entry script for creating new creature statblocks interactively. |
| `Authoring/New-StatDiceRow.ps1` | Parses individual stat lines from the CSV source. |
| `Authoring/Import-CreatureWeaponsFromText.ps1` | Converts weapon text tables into structured weapon objects. |
| `Authoring/New-HitLocationSheet.ps1` | Assigns the correct hit-location table (`Humanoid`, `Quadruped`, etc.). |
| `Statblock-tools.psm1` | Core engine: initializes contexts, builds statblocks, calculates HP/SR/DB/Move. |
| `Authoring/Add-CultInfoToStatblock.psm1` | Adds cult information and metadata to a statblock. |
| `Authoring/Get-CultData.psm1` | Cult data helpers: `Get-CultNames`, `Get-CultMagic`, `Get-CultRoles`, etc. |
| `Authoring/SpiritMagicRandomizer.psm1` | **Combined Spirit + Rune Magic allocator** module. |
| `_devLoader.ps1` | Development bootstrapper: imports all modules, helpers, and verifies data. |
| `Statblock.ps1` | Primary driver: builds, decorates, and prints a complete statblock. |

---

## 🧩 Core Flow

### 1️⃣ Creature Creation

```powershell
New-Creature -Creature 'Broo'
Parses a pasted BRP text entry.

Builds $sb (statblock object) using Statblock-tools.psm1.

Adds hit locations and weapons.

The resulting $sb is a baseline creature without magic or cult decoration.

2️⃣ Cult Decoration
Performed via Add-CultInfoToStatblock.psm1
(usually invoked automatically inside Statblock.ps1).

3️⃣ Magic Allocation
Both Spirit and Rune magic are applied after the base creature is created.

🔮 Spirit Magic System
Implemented in: Authoring/SpiritMagicRandomizer.psm1
Data: Y:\Stat_blocks\Data\spirit_magic_catalog.csv

Key Functions
Function	Purpose
Import-SpiritMagicCatalog	Loads the CSV catalog into memory.
Get-SpiritBudgetByRole	Returns Spirit Magic point budget by Role (CHA-capped).
New-RandomSpiritMagicLoadout	Randomly selects spells and intensities within budget.
Set-StatblockSpiritMagic	Writes chosen spells safely to $sb.Magic['Spirit'].

Allocation Rules
Only creatures with INT can learn spirit magic.

Total spirit points ≤ CHA.

Role defines typical budget:

Lay member: 2–4 points

Initiate: 5–10 points

Rune Lord/Priest: 10+ (future balance).

Variable spells (e.g., Bladesharp) roll random intensity within Role limits.

Data source drives available spell list and per-role maxima.

⚡ Rune Magic System
Implemented in: same module (SpiritMagicRandomizer.psm1)
Data: Y:\Stat_blocks\Data\Cults.xlsx

Key Functions
Function	Purpose
Get-RoleRunePointRange	Returns min/max Rune Points per Role.
New-RunePoints	Rolls total Rune Points (requires INT > 0).
Get-CultRuneSpellCatalog	Reads Rune spells (Common vs Special) from the cult sheet.
New-RuneSpellLoadout	Picks 1 Special spell per Rune Point. Common spells are listed automatically.
Set-StatblockRuneMagic	Writes Rune Magic data to $sb.Magic['RunePoints'], ['RuneSpecial'], ['RuneCommon'].

Allocation Rules
Role	Rune Points (min + roll)
Initiate	3 + 1–3
Rune Lord / Rune Lady	5 + 1–5
Rune Priest	6 + 1–8
Others	0

Each Rune Point grants one special Rune spell, regardless of that spell’s cost.

All Common cult spells are automatically available.

Associated cults’ Rune spells are included if linked in *_Associations.

📜 Printing Output (Statblock.ps1)
After creature generation, the script prints:

Core stats and derived values.

Chaos features (if any).

Cult + Role line (added recently).

Spirit Magic summary table (if present).

Rune Magic summary:

yaml
Copy code
Rune Magic: 7 Rune Points
  Special:
    Bladesharp
    Heal Body
  Common (always available):
    Extension
    Dispel Magic
    Reflection
Passions, Languages, Skills, Weapons, Hit Locations.

Run examples:

powershell
Copy code
.\Statblock.ps1 -Creature 'Broo'
.\Statblock.ps1 -Creature 'Broo' -Cult 'Thed' -Role 'Initiate' -Seed 77
.\Statblock.ps1 -Creature 'Broo' -Cult 'Thed' -Role 'RuneLord' -Seed 123
🧰 Dev Loader
_devLoader.ps1 ensures the working session is ready:

powershell
Copy code
# Loads core + authoring modules
Import-Module 'Y:\Stat_blocks\Scripts\Statblock-tools.psm1' -Force
Import-Module 'Y:\Stat_blocks\Scripts\Authoring\SpiritMagicRandomizer.psm1' -Force
Import-Module 'Y:\Stat_blocks\Scripts\Authoring\Get-CultData.psm1' -Force
Import-Module 'Y:\Stat_blocks\Scripts\Authoring\Add-CultInfoToStatblock.psm1' -Force

# Loads helpers
. "$root\New-StatDiceRow.ps1"
. "$root\Import-CreatureWeaponsFromText.ps1"
. "$root\New-HitLocationSheet.ps1"
. "$root\New-Creature.ps1"
It also imports ImportExcel, confirms the presence of
spirit_magic_catalog.csv and Cults.xlsx, and prints an “Available:” list of loaded functions.

📁 Data Files
File	Description
spirit_magic_catalog.csv	Core list of Spirit Magic spells; includes Min/Max/RoleMax columns.
Cults.xlsx	Contains multiple sheets per cult:
CultName_Magic, CultName_Associations, CultName_Roles.
(optional) StatDice.xlsx	Master stat source for creatures (if used by Initialize-StatblockContext).

🧱 Current Safety Guards
All property adds use Add-Member -Force–safe checks to avoid collisions.

Get-RoleRunePointRange normalizes role names ("RuneLord", "Rune Lord", etc.).

Fallback defaults prevent nulls or crashes when data files are missing.

Magic modules tolerate empty datasets (graceful “no spells” output).

🚀 Planned / Optional Enhancements
Add [switch]$NoSpiritMagic / [switch]$NoRuneMagic flags to Statblock.ps1.

Refine associated-cult Rune spell inclusion (use “Provides” column).

Auto-apply Spirit Magic within Add-CultInfoToStatblock for cult-specific tuning.

Split SpiritMagicRandomizer.psm1 → MagicAllocator.psm1 (future rename).

Add debug toggle to reduce verbose console noise.

🧭 Quick Reference
powershell
Copy code
# Create new creature statblock
$sb = New-Creature -Creature 'Broo'

# Decorate with cult + magic
$sb = Add-CultInfoToStatblock $sb -Cult 'Thed' -Role 'Initiate'

# Randomize spirit magic separately (for testing)
$cat  = Import-SpiritMagicCatalog -CsvPath "Y:\Stat_blocks\Data\spirit_magic_catalog.csv"
$rand = New-RandomSpiritMagicLoadout -PointsBudget 7 -CHA 12 -Role 'Initiate' -Catalog $cat
$sb   = Set-StatblockSpiritMagic -Statblock $sb -Spells $rand

# Generate full statblock with magic
.\Statblock.ps1 -Creature 'Broo' -Cult 'Thed' -Role 'RuneLord' -Seed 99
🧾 Changelog
Date	Change Summary
Oct 2025	Added combined Spirit + Rune magic allocator (SpiritMagicRandomizer.psm1).
Oct 2025	Integrated magic generation into Statblock.ps1.
Oct 2025	Expanded Add-CultInfoToStatblock and Get-CultData modules.
Oct 2025	Enhanced _devLoader.ps1 for data checks and module auto-imports.
Oct 2025	Added Cult + Role display to statblock printout.

Document version: October 2025
Maintainer: Paul Sturge