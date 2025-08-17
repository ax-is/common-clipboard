"""
Class to keep track of connected devices
"""

import time
import threading


class DeviceList:
    def __init__(self, timeout=30):
        self._devices = {}
        self.timeout = timeout
        self._lock = threading.Lock()

    def get_devices(self):
        device_list = []
        with self._lock:
            for ip, device in list(self._devices.items()):
                if time.time() - device['last active'] > self.timeout:
                    del self._devices[ip]
                else:
                    device_list.append((ip, device['name']))
        return device_list

    def add_device(self, ip, name):
        with self._lock:
            self._devices.update({ip: {
                'name': name,
                'last active': time.time(),
                'received': False
            }})

    def clear(self):
        with self._lock:
            self._devices.clear()

    def update_activity(self, ip):
        with self._lock:
            if ip in self._devices:
                self._devices[ip]['last active'] = time.time()

    def get_received(self, ip):
        with self._lock:
            return self._devices[ip]['received'] if ip in self._devices else False

    def set_received(self, ip, value):
        with self._lock:
            if ip in self._devices:
                self._devices[ip]['received'] = value
