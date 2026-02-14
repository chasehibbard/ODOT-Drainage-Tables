"""
ODOT Drainage Tables Generator
================================
Converts raw flex table exports from OpenRoads Designer into
formatted ODOT drainage tables for plan sheets.

Run this file to launch the application:
    python main.py
"""

import tkinter as tk
from gui import DrainageApp


def main():
    root = tk.Tk()
    DrainageApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
