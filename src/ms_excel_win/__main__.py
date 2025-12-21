from . import __version__


def main() -> None:
    """
    Temporary entry point to confirm the package wiring.
    Replace with Excel-specific startup logic once xlOil hooks are added.
    """

    print(f"ms-excel-win {__version__} is ready for xlOil integration.")


if __name__ == "__main__":
    main()
