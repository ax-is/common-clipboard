"""
Simple port editor dialog
"""

def _ask_port(current_port: int):
    import tkinter as tk
    from tkinter import simpledialog
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes('-topmost', True)
        return simpledialog.askinteger(
            "Port Editor",
            f"Current port: {current_port}\n\nEnter new port (1024-65535):",
            parent=root,
            minvalue=1024,
            maxvalue=65535,
            initialvalue=current_port
        )
    finally:
        try:
            root.destroy()
        except:
            pass

def edit_port(current_port=5000):
    """
    Minimal popup dialog to edit the port. Returns int or None.
    """
    try:
        return _ask_port(current_port)
    except Exception:
        return None


class PortEditor:
    def __init__(self, current_port):
        self.current_port = current_port
        self.new_port = edit_port(current_port)
    
    def get_port(self):
        return self.new_port
