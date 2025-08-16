"""
Main file for application
"""

import requests
import time
import win32clipboard as clipboard
import sys
import os
import pickle
from socket import gethostbyname, gethostname, gaierror
from threading import Thread
from multiprocessing import freeze_support, Value, Process
from multiprocessing.managers import BaseManager
from enum import Enum
from pystray import Icon, Menu, MenuItem
from PIL import Image
from io import BytesIO
from server import run_server
from device_list import DeviceList
from port_editor import PortEditor


class Format(Enum):
    TEXT = clipboard.CF_UNICODETEXT
    IMAGE = clipboard.RegisterClipboardFormat('PNG')


def register(address):
    global server_url

    server_url = f'http://{address}:{port}'
    try:
        hostname = gethostname()
        # Clean hostname to avoid issues with special characters and non-ASCII
        import re
        # Remove or replace problematic characters
        hostname = re.sub(r'[^\w\-_.]', '_', hostname)
        # Limit length and ensure it's not empty
        if not hostname or len(hostname) > 50:
            hostname = "Unknown_Device"
    except OSError:
        hostname = "Unknown_Device"
    
    requests.post(server_url + '/register', json={'name': hostname})


def test_server_ip(index):
    global server_url
    global running_server

    try:
        # Construct IP using our subnet and the provided index
        tested_ip = f'{split_ipaddr[0]}.{split_ipaddr[1]}.{split_ipaddr[2]}.{index}'
        tested_url = f'http://{tested_ip}:{port}'
        response = requests.get(tested_url + '/timestamp', timeout=2)  # Reduced timeout
        if response.ok and float(response.text) < server_timestamp.value:
            register(tested_ip)
            running_server = False
            server_process.terminate()
            systray.title = f'{APP_NAME}: Connected'
    except (requests.exceptions.ConnectionError, AssertionError):
        return


def generate_ips():
    # Only scan the local subnet (last octet) to avoid overwhelming the network
    # This is much more reasonable and won't cause WiFi issues
    for k in range(1, 255):  # Skip 0 and 255 (network and broadcast)
        if k != int(split_ipaddr[3]):  # Skip our own IP
            test_url_thread = Thread(target=test_server_ip, args=(k,), daemon=True)
            test_url_thread.start()
            # Add a small delay to prevent overwhelming the network
            time.sleep(0.01)


def find_server():
    global running_server
    global server_url

    # Start server without requiring internet connectivity
    start_server()
    register(ipaddr)
    systray.title = f'{APP_NAME}: Server Running'

    generator_thread = Thread(target=generate_ips)
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

    new_data, new_format = get_copied_data()
    if new_data != current_data:
        current_data = new_data
        current_format = new_format

        file = BytesIO()
        file.write(current_data.encode() if current_format == Format.TEXT else current_data)
        file.seek(0)
        requests.post(server_url + '/clipboard', data=file, headers={'Data-Type': format_to_type[current_format]})


def detect_server_change():
    global current_data
    global current_format

    headers = requests.head(server_url + '/clipboard', timeout=2)
    if headers.ok and headers.headers['Data-Attached'] == 'True':
        data_request = requests.get(server_url + '/clipboard')
        data_format = type_to_format[data_request.headers['Data-Type']]
        data = data_request.content.decode() if data_format == Format.TEXT else data_request.content

        try:
            clipboard.OpenClipboard()
            clipboard.EmptyClipboard()
            clipboard.SetClipboardData(data_format.value, data)
            clipboard.CloseClipboard()
            current_data, current_format = get_copied_data()
        except BaseException:
            return


def mainloop():
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
                systray.update_menu()
            time.sleep(LISTENER_DELAY)


def start_server():
    global running_server
    global server_process

    if server_process is not None:
        server_process.terminate()

    connected_devices.clear()
    running_server = True
    server_process = Process(target=run_server, args=(port, connected_devices, server_timestamp,))
    server_process.start()


def close():
    global run_app

    run_app = False
    if server_process is not None:
        server_process.terminate()

    with open(preferences_file, 'wb') as save_file:
        pickle.dump(port, save_file)

    systray.stop()
    sys.exit(0)


def edit_port():
    global port

    port_dialog = PortEditor(port)
    new_port = port_dialog.port_number.get()
    if port_dialog.applied and new_port != port:
        port = new_port
        find_server()


def get_menu_items():
    menu_items = (
        MenuItem(f'Port: {port}', Menu(MenuItem('Edit', lambda _: edit_port()))),
        MenuItem('View Connected Devices', Menu(lambda: (
            MenuItem(f"{name} ({ip})", None) for ip, name in connected_devices.get_devices()
        ))) if running_server else None,
        MenuItem('Quit', close),
    )
    return (item for item in menu_items if item is not None)


if __name__ == '__main__':
    freeze_support()

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

    BaseManager.register('DeviceList', DeviceList)
    manager = BaseManager()
    manager.start()
    connected_devices = manager.DeviceList()
    server_timestamp = Value('d', 0)

    running_server = False
    server_process: Process = None

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
    try:
        # Try to find the icon file in the current directory
        icon_path = 'systray_icon.ico'
        if not os.path.exists(icon_path):
            # If not found, try to get it from PyInstaller's temp directory
            if getattr(sys, 'frozen', False):
                # Running as PyInstaller executable
                base_path = sys._MEIPASS
                icon_path = os.path.join(base_path, 'systray_icon.ico')
            else:
                # Running as script, try relative to script location
                script_dir = os.path.dirname(os.path.abspath(__file__))
                icon_path = os.path.join(script_dir, 'systray_icon.ico')
        
        icon = Image.open(icon_path)
    except (FileNotFoundError, OSError):
        # Fallback: create a simple default icon
        icon = Image.new('RGBA', (64, 64), (100, 100, 100, 255))
    
    systray = Icon(APP_NAME, icon=icon, title=APP_NAME, menu=Menu(get_menu_items))
    systray.run_detached()

    run_app = True
    mainloop()
