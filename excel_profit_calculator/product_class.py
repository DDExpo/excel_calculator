

class Product():

    def __init__(self, name: str = '',
                 main_kaspi_name: str = '',
                 connected_kaspi_names: dict[str: tuple[str, str]] = None,
                 price: str = '0') -> None:

        self.name = name
        self.price = price
        self.main_kname = main_kaspi_name
        self.connected_kaspi_names = connected_kaspi_names or {}

    def __str__(self) -> str:
        return self.name
