"""
Simple port editor dialog
"""

import tkinter as tk
from tkinter import simpledialog

def edit_port(current_port=5000):
    """
    Minimal popup dialog to edit the port. Returns int or None.
    """
    root = tk.Tk()
    root.withdraw()
    try:
        # Keep it simple but ensure it appears on top briefly
        root.attributes('-topmost', True)
        new_port = simpledialog.askinteger(
            "Port Editor",
            f"Current port: {current_port}\n\nEnter new port (1024-65535):",
            parent=root,
            minvalue=1024,
            maxvalue=65535,
            initialvalue=current_port
        )
        return new_port
    except Exception:
        return None
    finally:
        try:
            root.destroy()
        except:
            pass


class PortEditor:
    def __init__(self, current_port):
        self.current_port = current_port
        self.new_port = edit_port(current_port)
    
    def get_port(self):
        return self.new_port
