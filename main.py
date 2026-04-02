"""Entry point for the ELEMENTS EML Parser application."""

import logging
import sys
import tkinter as tk

# Configure logging before importing application modules
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(name)s — %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)

from gui import MainWindow  # noqa: E402 (import after logging setup)


def main() -> None:
    """Create and run the main application window."""
    root = tk.Tk()
    MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
