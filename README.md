# PC Resale Check

A Windows tool that scans your PC, runs optional hands-on tests, and builds **one ready-to-paste prompt** for any AI (Claude, ChatGPT, Gemini...) so you get an estimated **resale value** with price adjustments, negotiation points and a confidence level.

No extra Python packages needed: it uses only the standard library (including `tkinter` for the popup windows).

## Features

**Automatic scan**
- Hardware: brand/model, serial, CPU, RAM, GPU, storage, display resolution, OS, BIOS
- Battery wear (design vs. full-charge capacity)
- Disk health and optional SMART data (power-on hours, SSD wear)
- Last N days of system errors: WHEA hardware errors, disk errors, unexpected shutdowns, BSODs

**Optional automatic checks**
- Wi-Fi adapter, signal, supported standards, internet and DNS test
- Bluetooth presence and errors
- Sound, camera, microphone, fingerprint and touchscreen devices
- Device Manager errors
- Windows / Office license status
- TPM version, Secure Boot, domain-joined (company managed) detection
- Display refresh rate and disk write speed (256 MB test)

**Optional hands-on tests**
- Sound test (left, right, both speakers)
- Full-screen color and motion test (dead pixels, flicker, backlight bleed)
- Keyboard, mouse and touchpad test
- Camera and microphone test
- Ports and extras checklist (USB, HDMI, jack, card reader, hinges, fan, BIOS password, account locks)

**Seller context and repair history**
- A questions popup collects country, currency, purchase year, original price, warranty, cosmetic condition, accessories and selling goal
- A `repair_log.csv` records repairs, replacements and unresolved faults with costs

Tests you skip are marked as **unverified** in the prompt, so the AI accounts for the uncertainty.

## Requirements

- Windows 10 or 11
- Python 3.8+ for Windows (with `tkinter`, included in the standard installer)
- Recommended: run from an **Administrator** terminal for full battery, event log, SMART and Secure Boot access
- Optional: [smartmontools](https://www.smartmontools.org/) (`smartctl` on PATH) for deeper disk checks

## Usage

```bash
python pc_checker.py
```

Then follow the windows:

1. Read the instructions (and optionally open `repair_log.csv`)
2. Answer the questions
3. Choose which tests to run
4. Wait for the scan (1-2 minutes) and do the hands-on tests
5. The prompt is copied to your clipboard. Paste it into your AI and press Enter

### Command line options

```bash
# Pre-fill the questions popup
python pc_checker.py --country Qatar --currency QAR --purchase-year 2021

# No popups (console only, automatic checks only)
python pc_checker.py --no-gui --country Qatar --currency QAR
```

| Option | Description |
| --- | --- |
| `--country` | Country where you will sell |
| `--currency` | Currency for the price (e.g. QAR, USD, EUR) |
| `--purchase-year` | Year you bought the PC |
| `--cosmetic` | Cosmetic condition |
| `--accessories` | Included accessories |
| `--days` | Event log window in days (default 30) |
| `--no-gui` | Console only, no popups (hands-on tests are skipped) |

## Repair log

On first run, `repair_log.csv` is created next to the script:

```csv
part,issue,status,date,cost,note
battery,Swollen battery,replaced,2025-03-10,180,Original replaced with new
screen,Dead pixel cluster,not_working,2026-01-05,450,Estimated panel cost
```

- `status` is one of `repaired`, `replaced`, `not_working`, `partial`
- `cost` is what you paid (repaired/replaced) or the **estimated** cost to fix (not_working/partial)
- Rows with status `example` are ignored

## Output

Each run saves the prompt to a `reports/` folder next to the script, named `<serial>_<date>_<time>_prompt.txt`.

The prompt asks the AI to:
1. Look up current used/sold prices for your exact model (if it can browse)
2. State the market price of a healthy unit
3. Itemize each price adjustment (battery, disk, errors, cosmetics, license, Windows 11 eligibility, repairs, unresolved faults...)
4. Give quick-sale, fair, optimistic and minimum prices in your currency
5. Add a confidence level and the top 3 things a buyer will use to negotiate

## Privacy

Everything runs locally. The script sends nothing anywhere. The generated prompt contains your PC serial number and hardware details, so review it before pasting it into an online AI and remove anything you don't want to share.

## Notes

- Windows only; the script exits on other systems.
- Some parsing (`netsh` output) assumes an English-language Windows.
- Reports and your personal `repair_log.csv` are better kept out of version control (see below).

### Suggested `.gitignore`

```
repair_log.csv
reports/
__pycache__/
```

## License

Add the license of your choice (for example MIT) as a `LICENSE` file.
