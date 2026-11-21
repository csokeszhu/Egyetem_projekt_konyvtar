"""Belépési pont:  python -m kef.main"""
from __future__ import annotations


def main() -> None:
    from kef.gui.app import App

    App().mainloop()


if __name__ == "__main__":
    main()
