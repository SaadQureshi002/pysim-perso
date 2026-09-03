import json
from pathlib import Path

import pandas as pd

from gsm_data_generator import DataGenerationScript, json_loader


REPO_ROOT = Path(__file__).resolve().parent
SETTINGS_FILE = REPO_ROOT / "settings.json"


def get_next_ranges(config):
    """
    Issuance ledger se next available ICCID/IMSI suggest karta hai.
    Agar ledger nahi hai to settings.json wali values use hoti hain.
    """

    output_dir = REPO_ROOT / config.PATHS.OUTPUT_FILES_DIR
    ledger_path = output_dir / ".issuance_ledger.json"

    default_iccid = config.DISP.iccid
    default_imsi = config.DISP.imsi

    if ledger_path.exists():
        try:
            with open(ledger_path, "r", encoding="utf-8") as file:
                ledger = json.load(file)

            if ledger:
                last_batch = ledger[-1]

                default_iccid = str(int(last_batch["iccid_end"]) + 1)
                default_imsi = str(int(last_batch["imsi_end"]) + 1)

        except Exception:
            pass

    return default_iccid, default_imsi


def ask_value(message, default):
    value = input(f"{message} [{default}]: ").strip()
    return value if value else str(default)


def main():
    print("=" * 60)
    print("        GSM / SIM DATA GENERATOR - EXCEL EXPORT")
    print("=" * 60)

    # Load base configuration
    config = json_loader(str(SETTINGS_FILE))

    # Automatically calculate next unused ranges
    default_iccid, default_imsi = get_next_ranges(config)

    print("\nEnter batch details.")
    print("Press ENTER to use the value shown in brackets.\n")

    start_iccid = ask_value("Starting ICCID", default_iccid)
    start_imsi = ask_value("Starting IMSI", default_imsi)
    batch_size = ask_value("Number of SIMs", config.DISP.size)
    file_name = ask_value("Output file name", config.PATHS.FILE_NAME)

    # Optional fixed SIM values
    pin1 = ask_value("PIN1", config.DISP.pin1)
    puk1 = ask_value("PUK1", config.DISP.puk1)
    pin2 = ask_value("PIN2", config.DISP.pin2)
    puk2 = ask_value("PUK2", config.DISP.puk2)

    # Apply user inputs
    config.DISP.iccid = start_iccid
    config.DISP.imsi = start_imsi
    config.DISP.size = int(batch_size)

    config.DISP.pin1 = pin1
    config.DISP.puk1 = puk1
    config.DISP.pin2 = pin2
    config.DISP.puk2 = puk2

    config.PATHS.FILE_NAME = file_name

    print("\nGenerating SIM data...")

    script = DataGenerationScript(config)
    script.json_to_global_params()

    # Generate data
    result_dfs, keys = script.generate_all_data()

    print(f"\nBatch size       : {script.params.DATA_SIZE}")
    print(f"Starting ICCID   : {start_iccid}")
    print(f"Starting IMSI    : {start_imsi}")
    print(f"Generated frames : {', '.join(sorted(result_dfs))}")

    # Original TXT outputs + issuance protection
    written = script.write_outputs(result_dfs)

    print("\nTXT files:")
    for output_type, path in sorted(written.items()):
        print(f"{output_type:<7} -> {path}")

    # Excel output
    output_dir = REPO_ROOT / config.PATHS.OUTPUT_FILES_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    excel_path = output_dir / f"{file_name}.xlsx"

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for sheet_name, dataframe in result_dfs.items():
            dataframe.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    print("\n" + "=" * 60)
    print("SUCCESS")
    print("=" * 60)
    print(f"Excel file created:\n{excel_path}")


if __name__ == "__main__":
    main()