import os.path
import time

import cleanup
import config
import dataframe
import dialog_setup
import dialogs
import pywintypes
import spreadsheet
import statistics


def main():
    program_start_time = time.time()
    cfg = config.load_config()
    log = config.initialize_logging(cfg)

    log.info("Starting Customer Data Cleaner")
    log.info(f"Configuration loaded from config.toml")

    (
        input_file_path,
        output_file_path,
        output_worksheet,
    ) = config.get_spreadsheet_paths_and_worksheet(cfg)

    log.info(f"Input file: {input_file_path}")
    log.info(f"Output file: {output_file_path}")
    log.info(f"Output worksheet: {output_worksheet}")

    log.info("Loading input dataframe")
    df = dataframe.get_dataframe(input_file_path)
    original_df = df.copy()
    log.info(f"Loaded {len(df)} records from input file")

    log.info("Starting data cleaning")
    cleanup.clean_data(df, cfg)
    finished_cleaning_time = time.time()
    log.info(f"Data cleaning completed in {finished_cleaning_time - program_start_time:.2f} seconds")

    log.info("Calculating statistics")
    stats = statistics.get_statistics(
        program_start_time,
        finished_cleaning_time,
        original_df,
        df,
        cfg["columns"]["email"],
    )
    log.info(f"Statistics: {stats}")

    log.info("Launching Excel application")
    excel = spreadsheet.launch_spreadsheet_app()

    try:
        if not os.path.exists(output_file_path):
            log.info(f"Creating new workbook: {output_file_path}")
            wb, ws = spreadsheet.create_new_workbook(
                excel,
                output_file_path,
                output_worksheet,
            )
        else:
            log.info(f"Opening existing workbook: {output_file_path}")
            wb, ws = spreadsheet.open_spreadsheet(
                excel,
                output_file_path,
                output_worksheet,
            )

        log.info("Removing existing filters")
        spreadsheet.remove_filters(ws)

        log.info("Writing cleaned data to spreadsheet")
        spreadsheet.input_values_to_spreadsheet(df, ws)

        log.info("Auto-fitting columns and rows")
        spreadsheet.auto_fit_columns_and_rows(ws)

        log.info("Left-aligning text")
        spreadsheet.left_align_text(ws)

        log.info("Saving workbook")
        wb.Save()

        log.info("Setting up and running main dialog")
        dialog_setup.setup_and_run_dialog(
            excel,
            wb,
            output_file_path,
            ws,
            cfg,
            stats,
        )

    finally:
        log.info("Closing Excel application")
        try:
            excel.Quit()
        except pywintypes.com_error:
            log.warning("Failed to close Excel application")

    log.info("Customer Data Cleaner completed successfully")

if __name__ == "__main__":
    main()