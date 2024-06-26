import os
import subprocess
import requests
import pythoncom
import win32serviceutil
import win32service
import win32event
import configparser
import time
import socket

# 請在這裡填入你的Line Notify的token和WeChat的QR login path
LINE_NOTIFY_TOKEN = 'JatEg0S1wS2tLn6I9Wok56dB1cujI0RRG2a5G9a4C93'


class PingService(win32serviceutil.ServiceFramework):
    _svc_name_ = "PingService"
    _svc_display_name_ = "Ping Service"
    _svc_description_ = "Ping multiple hosts and send Line or WeChat notification on failure"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.is_running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        self.is_running = False

    def SvcDoRun(self):
        self.main()

    def send_line_notification(self, message):
        url = 'https://notify-api.line.me/api/notify'
        headers = {
            'Authorization': f'Bearer {LINE_NOTIFY_TOKEN}'
        }
        data = {
            'message': message
        }
        requests.post(url, headers=headers, data=data)


    def ping(self, host):
        command = ['ping', '-n', '5', host]  # 在Windows下使用 '-n' 來指定ping的次數
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return result.returncode == 0

    def read_config(self):
        config = configparser.ConfigParser()
        path = os.path.dirname(os.path.abspath(__file__))
        config_file = os.path.join(path, "config.ini")
        if not os.path.exists(config_file):
            raise FileNotFoundError(f'Config file "{config_file}" not found.')

        config.read(config_file)
        if 'Settings' not in config:
            raise ValueError('Missing [Settings] section in config file.')

        if 'IPList' not in config:
            raise ValueError('Missing [IPList] section in config file.')

        ip_list = {}
        for key in config['IPList']:
            ip_address = key
            description = config['IPList'][key]
            ip_list[ip_address] = description

        return config, ip_list

    def main(self):
        try:
            config, ip_list = self.read_config()
        except (FileNotFoundError, ValueError) as e:
            print(f'Error reading config: {e}')
            return

        while self.is_running:
            for ip, description in ip_list.items():
                if not self.ping(ip):
                    source_ip = socket.gethostbyname(socket.gethostname())
                    message = f'{source_ip}無法Ping通{description} ({ip})！'
                    if config.getboolean('Settings', 'UseLine', fallback=False) and LINE_NOTIFY_TOKEN:
                        self.send_line_notification(message)

            # 等待一段時間後再進行下一次Ping檢查
            time.sleep(300)  # 等待5分鐘


if __name__ == '__main__':
    pythoncom.CoInitialize()
    win32serviceutil.HandleCommandLine(PingService)
