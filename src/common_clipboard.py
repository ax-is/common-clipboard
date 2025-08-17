"""
Main file for application
"""

import requests
import time
import win32clipboard as clipboard
import sys
import os
import pickle
import re
import socket
import threading
import msvcrt
from socket import gethostbyname, gethostname, gaierror
from concurrent.futures import ThreadPoolExecutor
from threading import Thread
from multiprocessing import freeze_support
from enum import Enum
from pystray import Icon, Menu, MenuItem
from PIL import Image
from io import BytesIO
from server import run_server
from device_list import DeviceList
from port_editor import PortEditor
import winreg


class Format(Enum):
    TEXT = clipboard.CF_UNICODETEXT
    IMAGE = clipboard.RegisterClipboardFormat('PNG')

# Connection state management
connection_lock = threading.Lock()
# Guards and shared client
scan_in_progress = threading.Event()
port_dialog_open = threading.Event()
http = requests.Session()

# Global variables for single instance control
instance_lock = None


def check_single_instance():
    """Check if another instance is already running"""
    global instance_lock
    try:
        lock_file = os.path.join(os.getenv('TEMP', os.getcwd()), 'common_clipboard.lock')
        instance_lock = open(lock_file, 'w')
        msvcrt.locking(instance_lock.fileno(), msvcrt.LK_NBLCK, 1)
        return True
    except (IOError, OSError):
        return False


def register(address):
    global server_url

    server_url = f'http://{address}:{port}'
    try:
        hostname = gethostname()
        # Clean hostname to avoid issues with special characters and non-ASCII
        # Remove or replace problematic characters
        hostname = re.sub(r'[^\w\-_.]', '_', hostname)
        # Limit length and ensure it's not empty
        if not hostname or len(hostname) > 50:
            hostname = "Unknown_Device"
    except OSError:
        hostname = "Unknown_Device"
    
    http.post(server_url + '/register', json={'name': hostname})


def test_server_ip(index):
    global server_url
    global running_server
    global server_thread

    try:
        # Construct IP using our subnet and the provided index
        tested_ip = f'{split_ipaddr[0]}.{split_ipaddr[1]}.{split_ipaddr[2]}.{index}'
        tested_url = f'http://{tested_ip}:{port}'
        
        # Test timestamp endpoint (original repository method)
        try:
            response = http.get(tested_url + '/timestamp', timeout=2)
            if response.ok and float(response.text) < server_timestamp:
                # If we find a suitable remote server, stop local server thread cleanly
                try:
                    if server_thread is not None and server_thread.is_alive():
                        try:
                            http.post(f'http://127.0.0.1:{port}/shutdown', timeout=2)
                        except Exception:
                            pass
                        server_thread.join(timeout=3)
                finally:
                    server_thread = None
                    running_server = False

                server_url = tested_url
                register(tested_ip)
                systray.title = f'{APP_NAME}: Connected'
        except (ValueError, requests.exceptions.ConnectionError, requests.exceptions.Timeout, Exception):
            # Silently ignore connection errors during discovery
            pass
            
    except Exception:
        # Silently ignore connection errors during discovery
        pass


def generate_ips():
    # Bounded concurrency scan to reduce Wi‑Fi/CPU pressure
    if scan_in_progress.is_set():
        return
    scan_in_progress.set()
    try:
        max_workers = 16
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            for k in range(1, 255):  # 1..254
                if k == int(split_ipaddr[3]):
                    continue
                pool.submit(test_server_ip, k)
    finally:
        scan_in_progress.clear()


def find_server():
    global running_server
    global server_url

    # Start server without requiring internet connectivity
    start_server()
    # Wait briefly until the local server is listening to avoid race on first register()
    start_ts = time.time()
    while time.time() - start_ts < 2.0:
        try:
            resp = http.get(f'http://127.0.0.1:{port}/timestamp', timeout=0.3)
            if resp.ok:
                break
        except Exception:
            time.sleep(0.1)
    register(ipaddr)
    systray.title = f'{APP_NAME}: Server Running'

    generator_thread = Thread(target=generate_ips, daemon=True)
    generator_thread.start()


def get_copied_data():
    try:
        for fmt in list(Format):
            if clipboard.IsClipboardFormatAvailable(fmt.value):
                clipboard.OpenClipboard()
                data = clipboard.GetClipboardData(fmt.value)
                clipboard.CloseClipboard()
                return data, fmt
        else:
            raise BaseException
    except BaseException:
        try:
            return current_data, current_format
        except NameError:
            return '', Format.TEXT


def detect_local_copy():
    global current_data
    global current_format

    with connection_lock:
        if not server_url:
            return

        new_data, new_format = get_copied_data()
        if new_data != current_data:
            current_data = new_data
            current_format = new_format

            file = None
            try:
                file = BytesIO()
                file.write(current_data.encode() if current_format == Format.TEXT else current_data)
                file.seek(0)
                response = http.post(server_url + '/clipboard', data=file, headers={'Data-Type': format_to_type[current_format]}, timeout=5)
                if not response.ok:
                    print(f"Failed to send clipboard data: {response.status_code}")
            except Exception as e:
                print(f"Error sending clipboard data: {e}")
            finally:
                if file:
                    try:
                        file.close()
                    except:
                        pass


def detect_server_change():
    global current_data
    global current_format

    with connection_lock:
        if not server_url:
            return

        try:
            headers = http.head(server_url + '/clipboard', timeout=3)
            if headers.ok and headers.headers.get('Data-Attached') == 'True':
                data_request = http.get(server_url + '/clipboard', timeout=5)
                if data_request.ok:
                    data_format = type_to_format[data_request.headers['Data-Type']]
                    data = data_request.content.decode() if data_format == Format.TEXT else data_request.content

                    try:
                        clipboard.OpenClipboard()
                        clipboard.EmptyClipboard()
                        clipboard.SetClipboardData(data_format.value, data)
                        clipboard.CloseClipboard()
                        current_data, current_format = get_copied_data()
                    except Exception as e:
                        print(f"Error updating clipboard: {e}")
                        try:
                            clipboard.CloseClipboard()
                        except:
                            pass
        except Exception as e:
            print(f"Error checking server changes: {e}")


def mainloop():
    last_menu_update = 0.0
    while run_app:
        try:
            if server_url:  # Only try to detect changes if we have a server URL
                detect_server_change()
            detect_local_copy()
        except (requests.exceptions.ConnectionError, TimeoutError, OSError) as e:
            if not server_url:  # Only try to find server if we don't have one
                find_server()
        finally:
            if running_server:
                now = time.time()
                if now - last_menu_update >= 1.0:
                    systray.update_menu()
                    last_menu_update = now
            time.sleep(LISTENER_DELAY)


def start_server():
    global running_server
    global server_thread

    if server_thread is not None and server_thread.is_alive():
        try:
            http.post(f'http://127.0.0.1:{port}/shutdown', timeout=2)
        except Exception:
            pass
        server_thread.join(timeout=3)

    connected_devices.clear()
    running_server = True

    # Launch Flask in a background thread
    def _run():
        run_server(port, connected_devices, server_timestamp)

    server_thread = Thread(target=_run, daemon=True)
    server_thread.start()


# ---------------- Startup on Login (Windows) ----------------
def _startup_reg_path():
    return r"Software\\Microsoft\\Windows\\CurrentVersion\\Run"


def is_startup_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _startup_reg_path(), 0, winreg.KEY_READ) as k:
            try:
                winreg.QueryValueEx(k, APP_NAME)
                return True
            except FileNotFoundError:
                return False
    except OSError:
        return False


def toggle_startup():
    if not getattr(sys, 'frozen', False):
        print("Startup toggle is available only in the packaged app.")
        return
    exe_path = sys.executable
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _startup_reg_path(), 0, winreg.KEY_ALL_ACCESS) as k:
            if is_startup_enabled():
                try:
                    winreg.DeleteValue(k, APP_NAME)
                except FileNotFoundError:
                    pass
            else:
                winreg.SetValueEx(k, APP_NAME, 0, winreg.REG_SZ, f'"{exe_path}"')
    except OSError as e:
        print(f"Startup toggle error: {e}")


def close():
    global run_app

    run_app = False
    
    # Clean up server process
    if server_thread is not None and server_thread.is_alive():
        try:
            http.post(f'http://127.0.0.1:{port}/shutdown', timeout=2)
        except Exception:
            pass
        server_thread.join(timeout=3)

    # Save preferences
    try:
        with open(preferences_file, 'wb') as save_file:
            pickle.dump(port, save_file)
    except Exception as e:
        print(f"Error saving preferences: {e}")

    # Clean up global variables
    if instance_lock:
        try:
            instance_lock.close()
        except:
            pass

    systray.stop()
    sys.exit(0)


def toggle_server():
    """Toggle server on/off"""
    global running_server
    global server_thread
    global server_url
    
    if running_server:
        # Stop server
        if server_thread is not None and server_thread.is_alive():
            try:
                http.post(f'http://127.0.0.1:{port}/shutdown', timeout=2)
            except Exception as e:
                print(f"Error stopping server: {e}")
            server_thread.join(timeout=3)
            server_thread = None
            running_server = False
            server_url = ''
            connected_devices.clear()
            print("Server stopped")
            systray.title = f'{APP_NAME}: Stopped'
    else:
        # Start server
        find_server()


def edit_port():
    global port
    global running_server
    global server_thread

    # Prevent multiple dialogs
    if port_dialog_open.is_set():
        return
    port_dialog_open.set()
    try:
        # Stop the server first
        if server_thread is not None and server_thread.is_alive():
            try:
                http.post(f'http://127.0.0.1:{port}/shutdown', timeout=2)
            except Exception as e:
                print(f"Error stopping server: {e}")
            server_thread.join(timeout=3)
            server_thread = None
            running_server = False

        # Clear connected devices
        connected_devices.clear()

        # Show port dialog
        port_dialog = PortEditor(port)
        new_port = port_dialog.get_port()

        # Change port if user provided a new one
        if new_port is not None and new_port != port:
            port = new_port
            print(f"Port changed to: {port}")

        # Restart server with new port
        find_server()
    finally:
        port_dialog_open.clear()


def toggle_dark_icon():
    global use_dark_icon
    use_dark_icon = not use_dark_icon
    try:
        new_icon = load_icon(use_dark_icon)
        systray.icon = new_icon
        systray.update_menu()
    except Exception as e:
        print(f"Icon toggle error: {e}")


def get_menu_items():
    menu_items = (
        MenuItem('Stop Server' if running_server else 'Start Server', lambda _: toggle_server()),
        MenuItem(f'Port: {port}', Menu(MenuItem('Edit', lambda _: Thread(target=edit_port, daemon=True).start()))),
        MenuItem('Start on Login: On' if is_startup_enabled() else 'Start on Login: Off', lambda _: toggle_startup()),
        MenuItem('View Connected Devices', Menu(lambda: (
            MenuItem(f"{name} ({ip})", None) for ip, name in connected_devices.get_devices()
        ))) if running_server else None,
        MenuItem('Quit', close),
    )
    return (item for item in menu_items if item is not None)


if __name__ == '__main__':
    freeze_support()

    # Check for single instance
    if not check_single_instance():
        print("Another instance of Common Clipboard is already running.")
        sys.exit(1)

    APP_NAME = 'Common Clipboard'
    LISTENER_DELAY = 0.3

    server_url = ''
    try:
        ipaddr = gethostbyname(gethostname())
    except (gaierror, OSError):
        # Fallback: try to get IP from local network interfaces
        import socket
        try:
            # Get local IP without making external connections
            with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
                # Don't actually connect, just bind to get local address
                s.bind(('', 0))
                ipaddr = s.getsockname()[0]
                if ipaddr == '0.0.0.0':
                    # If we get 0.0.0.0, try a different approach
                    ipaddr = socket.gethostbyname('localhost')
        except OSError:
            # Final fallback: use localhost
            ipaddr = "127.0.0.1"
    split_ipaddr = [int(num) for num in ipaddr.split('.')]
    base_ipaddr = str(split_ipaddr[0])

    connected_devices = DeviceList()
    server_timestamp = time.time()

    running_server = False
    server_thread: Thread | None = None

    try:
        data_dir = os.path.join(os.getenv('LOCALAPPDATA'), APP_NAME)
    except TypeError:
        data_dir = os.getcwd()

    if not os.path.exists(data_dir):
        os.makedirs(data_dir)

    preferences_file = os.path.join(data_dir, 'preferences.pickle')
    try:
        with open(preferences_file, 'rb') as preferences:
            port = pickle.load(preferences)
    except FileNotFoundError:
        port = 5000

    current_data, current_format = get_copied_data()

    format_to_type = {Format.TEXT: 'text', Format.IMAGE: 'image'}
    type_to_format = {v: k for k, v in format_to_type.items()}

    # Handle icon path for both development and PyInstaller executable
    def load_icon():
        try:
            icon_name = 'systray_icon.ico'
            icon_path = icon_name
            if not os.path.exists(icon_path):
                if getattr(sys, 'frozen', False):
                    base_path = sys._MEIPASS
                    icon_path = os.path.join(base_path, icon_name)
                else:
                    script_dir = os.path.dirname(os.path.abspath(__file__))
                    icon_path = os.path.join(script_dir, icon_name)
            return Image.open(icon_path)
        except (FileNotFoundError, OSError):
            return Image.new('RGBA', (64, 64), (100, 100, 100, 255))

    icon = load_icon()
    
    systray = Icon(APP_NAME, icon=icon, title=APP_NAME, menu=Menu(get_menu_items))
    systray.run_detached()

    run_app = True
    # Start server immediately instead of waiting for connection error
    find_server()
    mainloop()
