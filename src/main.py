import cleanup
import config
import data
import dialog_setup
import spreadsheet


def main():
    cfg = config.load_config()
    file_loc, worksheet_name = config.get_spreadsheet(cfg)
    df = data.get_dataframe(file_loc)
    cleanup.clean_data(df, cfg)
    excel = spreadsheet.launch_spreadsheet_app()
    wb, ws = spreadsheet.open_spreadsheet(excel, file_loc, worksheet_name)
    data.apply_dataframe_to_spreadsheet(df, ws)
    dialog_setup.setup_and_run_dialog(excel, wb, file_loc, ws, cfg)


if __name__ == '__main__':
    main()
