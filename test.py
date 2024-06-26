import subprocess
import requests
import itchat
import pythoncom
import win32serviceutil
import win32service
import win32event
import win32timezone
import configparser
import os
import time

# 請在這裡填入你的Line Notify的token和WeChat的QR login path
LINE_NOTIFY_TOKEN = 'fB9Nl80MaXPPDOpfGoHXJLiJer2RcM0fqkaDe8xEX0d'
WECHAT_QR_LOGIN_PATH = '你的WeChat QR login path'


class PingService():
    _svc_name_ = "PingService"
    _svc_display_name_ = "Ping Service"
    _svc_description_ = "Ping multiple hosts and send Line or WeChat notification on failure"

    def send_line_notification(self, message):
        url = 'https://notify-api.line.me/api/notify'
        headers = {
            'Authorization': f'Bearer {LINE_NOTIFY_TOKEN}'
        }
        data = {
            'message': message
        }
        requests.post(url, headers=headers, data=data)

    def send_wechat_notification(self, message):
        itchat.auto_login(hotReload=True, qrCallback=self.generate_qr_callback())
        itchat.send(message, toUserName='filehelper')

    def generate_qr_callback(self):
        def qr_callback(uuid, status, qrcode):
            if status == '0':
                print('Please scan the QR code to log in.')
                print(qrcode)
                with open(WECHAT_QR_LOGIN_PATH, 'wb') as f:
                    f.write(qrcode)

        return qr_callback

    def ping(self, host):
        command = ['ping', '-n', '5', host]  # 在Windows下使用 '-n' 來指定ping的次數
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return result.returncode == 0

    def read_config(self):
        config = configparser.ConfigParser()
        config_file = 'config.ini'
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

        for ip, description in ip_list.items():
            if not self.ping(ip):
                message = f'{description} ({ip}) 無法Ping通！'
                if config.getboolean('Settings', 'UseLine', fallback=False) and LINE_NOTIFY_TOKEN:
                    self.send_line_notification(message)
                if config.getboolean('Settings', 'UseWeChat', fallback=False) and WECHAT_QR_LOGIN_PATH:
                    self.send_wechat_notification(message)



ping_service = PingService()
ping_service.main()
