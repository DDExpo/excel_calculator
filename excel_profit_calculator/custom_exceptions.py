

class NoneInColumnName(Exception):
    """Exception raised when None is found in a column name."""

    def __str__(self):
        return f"None value found in column name: {self.column_name}"
