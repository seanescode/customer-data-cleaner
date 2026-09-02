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

    (
        input_file_path,
        output_file_path,
        output_worksheet,
    ) = config.get_spreadsheet_paths_and_worksheet(cfg)

    df = dataframe.get_dataframe(input_file_path)
    original_df = df.copy()

    cleanup.clean_data(df, cfg)

    finished_cleaning_time = time.time()

    stats = statistics.get_statistics(
        program_start_time,
        finished_cleaning_time,
        original_df,
        df,
        cfg["columns"]["email"],
    )
    excel = spreadsheet.launch_spreadsheet_app()

    try:
        if not os.path.exists(output_file_path):
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = output_worksheet
            wb.SaveAs(output_file_path)
        else:
            wb = excel.Workbooks.Open(output_file_path)
            ws = wb.Worksheets(output_worksheet)

        spreadsheet.input_values_to_spreadsheet(df, ws)
        spreadsheet.auto_fit_columns_and_rows(ws)
        wb.Save()
        dialogs.show_statistics_dialog(stats)

        dialog_setup.setup_and_run_dialog(
            excel,
            wb,
            output_file_path,
            ws,
            cfg,
        )

    finally:
        try:
            excel.Quit()
        except pywintypes.com_error:
            pass


if __name__ == "__main__":
    main()