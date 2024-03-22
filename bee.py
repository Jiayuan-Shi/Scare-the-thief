import time
import logging
import platform
import os
import subprocess
import psutil
import pygame
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

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
            pygame.mixer.music.load("D:\\python\\bee\\dg.mp3")
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
                set_max_volume_and_beep()
        except Exception as e:
            logging.error(f"执行过程中出现错误: {e}")
        time.sleep(5)  # 每秒检测一次
