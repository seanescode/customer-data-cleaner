import cleanup
import config
import dataframe
import dialog_setup
import spreadsheet
import pywintypes
import time
import statistics

def main():
    program_start_time = time.time()
    cfg = config.load_config()

    (input_file_path,
     output_file_path,
     output_worksheet) = config.get_spreadsheet_paths_and_worksheet(cfg)

    df = dataframe.get_dataframe(input_file_path)
    cleanup.clean_data(df, cfg)
    finished_cleaning_time = time.time()


    stats = statistics.get_statistics(
        program_start_time,
        finished_cleaning_time,
        df,
        cfg["columns"]["email"])

    print(stats)
    excel = spreadsheet.launch_spreadsheet_app()


    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = output_worksheet
        spreadsheet.input_values_to_spreadsheet(df, ws)
        spreadsheet.auto_fit_columns_and_rows(ws)
        wb.SaveAs(output_file_path)
        dialog_setup.setup_and_run_dialog(
            excel,
            wb,
            output_file_path,
            ws,
            cfg
        )

    finally:
        try:
            excel.Quit()
        except pywintypes.com_error:
            pass

if __name__ == '__main__':
    main()