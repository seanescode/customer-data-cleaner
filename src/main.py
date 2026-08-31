import cleanup
import config
import dialog_setup
import spreadsheet


def main():
    cfg = config.load_config()
    file_path, worksheet_name = config.get_spreadsheet_path_and_worksheet(cfg)
    df = spreadsheet.get_dataframe(file_path)
    cleanup.clean_data(df, cfg)
    excel = spreadsheet.launch_spreadsheet_app()
    wb, ws = spreadsheet.open_spreadsheet(excel, file_path, worksheet_name)
    spreadsheet.update_spreadsheet_from_dataframe(df, ws)
    dialog_setup.setup_and_run_dialog(excel, wb, file_path, ws, cfg)


if __name__ == '__main__':
    main()
