class CustomerDataCleanerError(Exception):
    """Base exception for all customer data cleaner errors"""
    pass


class ConfigNotFoundError(CustomerDataCleanerError):
    """Raised when config file is not found"""
    def __init__(self, config_path):
        self.config_path = config_path
        super().__init__(f"Config file not found: {config_path}")


class ConfigParseError(CustomerDataCleanerError):
    """Raised when config file cannot be parsed"""
    def __init__(self, error_message):
        super().__init__(f"Invalid TOML in config file: {error_message}")


class ConfigKeyError(CustomerDataCleanerError):
    """Raised when required config key is missing"""
    def __init__(self, key):
        super().__init__(f"Missing required config key: {key}")


class SpreadsheetNotFoundError(CustomerDataCleanerError):
    """Raised when spreadsheet file is not found"""
    def __init__(self, file_path):
        self.file_path = file_path
        super().__init__(f"Spreadsheet not found: {file_path}")


class SpreadsheetAccessError(CustomerDataCleanerError):
    """Raised when spreadsheet cannot be accessed (permissions, etc.)"""
    def __init__(self, error_message):
        super().__init__(f"Error reading Excel file: {error_message}")


class DataCleaningError(CustomerDataCleanerError):
    """Raised when data cleaning fails"""
    def __init__(self, error_message):
        super().__init__(f"Data cleaning error: {error_message}")


class ExcelConnectionError(CustomerDataCleanerError):
    """Raised when Excel connection fails"""
    def __init__(self, error_message=None):
        if error_message:
            super().__init__(f"Excel connection error: {error_message}")
        else:
            super().__init__("Excel connection error. Please ensure Excel is still open.")


class DuplicateDetectionError(CustomerDataCleanerError):
    """Raised when duplicate detection fails"""
    def __init__(self, error_message):
        super().__init__(f"Duplicate detection error: {error_message}")


class EmailValidationError(CustomerDataCleanerError):
    """Raised when email validation fails"""
    def __init__(self, error_message):
        super().__init__(f"Email validation error: {error_message}")


class ColumnNotFoundError(CustomerDataCleanerError):
    """Raised when required column is not found in data"""
    def __init__(self, column_name):
        super().__init__(f"Column '{column_name}' not found in spreadsheet")


class DialogSetupError(CustomerDataCleanerError):
    """Raised when dialog setup fails"""
    def __init__(self, error_message):
        super().__init__(f"Dialog setup error: {error_message}")
