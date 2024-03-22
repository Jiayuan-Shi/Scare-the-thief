import time
import logging
import platform
import os
import subprocess
import psutil
import pygame
import requests
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 全局变量，记录上次请求执行的时间
last_request_time = 0
# 设置请求之间的最小时间间隔（单位：秒）
request_interval = 3600  # 1小时

def request_pushplus():
    global last_request_time
    current_time = time.time()
    if current_time - last_request_time >= request_interval:
        token = 'b****f' #在pushplus网站中可以找到
        title= '电脑断电！' #改成你要的标题内容
        content ='实验室的电脑' #改成你要的正文内容
        url = 'http://www.pushplus.plus/send?token='+token+'&title='+title+'&content='+content
        requests.get(url)
        last_request_time = current_time

def is_power_disconnected():
    try:
        if psutil.sensors_battery():
            power_plugged = psutil.sensors_battery().power_plugged
            if power_plugged:
                logging.info("电源已连接")
                return False
            else:
                logging.info("电源已断开")
                return True
        else:
            logging.warning("未检测到电源信息")
            return False
    except Exception as e:
        logging.error(f"检测电源状态时出错: {e}")
        return False

def set_max_volume_and_beep():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, 
        CLSCTX_ALL, 
        None
    )
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    try:
        mute = volume.GetMute()
        volumeRange = volume.GetVolumeRange()
        volume.SetMasterVolumeLevel(volumeRange[1], None)  # 设置音量为最大值
        if mute:
            volume.SetMute(0, None)  # 如果之前是静音状态，取消静音
    except Exception as e:
        logging.error(f"设置音量时出错: {e}")
    if platform.system() == "Windows":

        # Windows系统设置最大音量
        try:
            os.system("powershell (new-object -com wscript.shell).SendKeys([char]176)")
            # 播放蜂鸣声
            pygame.mixer.init()
            pygame.mixer.music.load("dg.mp3")
            pygame.mixer.music.set_volume(1.0)
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy():
                pygame.time.Clock().tick(10)

        except Exception as e:
            logging.error(f"设置Windows系统音量或播放蜂鸣声时出错: {e}")
    elif platform.system() == "Linux":
        # Linux系统设置最大音量
        try:
            os.system("amixer sset 'Master' 100%")
            # 播放蜂鸣声
            os.system("beep")
            logging.info("Linux系统音量已调至最大并播放蜂鸣声")
        except Exception as e:
            logging.error(f"设置Linux系统音量时出错: {e}")

if __name__ == "__main__":
    logging.info("脚本开始执行...")
    while True:
        try:
            if is_power_disconnected():
                logging.info("检测到电源被拔掉，正在处理...")
                request_pushplus()
                set_max_volume_and_beep()
        except Exception as e:
            logging.error(f"执行过程中出现错误: {e}")
        time.sleep(5)  # 每秒检测一次
