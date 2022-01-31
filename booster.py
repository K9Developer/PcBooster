import os
import shutil
import subprocess
import winreg
import win32con
import win32gui

# Sets vars for the active user and system disk
DISK = os.getenv("SystemDrive")
ACTIVE_USER = os.getenv('username')


def set_reg(name, value, path, root=winreg.HKEY_CURRENT_USER):

    """
    Sets a value for a key in the registry.

    :param name: The name of the key that we need to edit
    :type name: str
    :param value: The value of the key we want to edit
    :type value: int | str
    :param path: The path for the key
    :type path: str
    :param root: The root of the key, example: HKEY_CURRENT_USER
    :type root: int
    :return: None
    :rtype: None
    """

    registry_key = winreg.OpenKey(root, path, 0,
                                  winreg.KEY_WRITE)
    winreg.SetValueEx(registry_key, name, 0, winreg.REG_SZ, value)
    winreg.CloseKey(registry_key)


def clean_disk(temp_files=True, cache=True, win_temp=True, recycle_bin=True):

    """
    Runs some code to delete cache and tmp files and more.

    Paths to delete:
        - DISK/Users/ACTIVE_USER/AppData/Local/Temp - temp files
        - DISK/Users/ACTIVE_USER/AppData/Local/Microsoft/Windows/INetCache/IE - cache files
        - DISK/Windows/Temp - windows temp files
        - Recycle bin - recycle bin

    :param temp_files: If we should delete the temp files
    :type temp_files: bool
    :param cache: If we should delete the cache files
    :type cache: bool
    :param win_temp: If we should delete the windows temp files
    :type win_temp: bool
    :param recycle_bin: If we should empty the recycle bin
    :type recycle_bin: bool
    :return: None
    :rtype: None
    """


    # Deleting all permitted contents in {DISK}\Users\{ACTIVE_USER}\AppData\Local\Temp
    if temp_files:
        current_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Local\Temp'
        for file in os.listdir(current_path):
            file = os.path.join(current_path, file)
            shutil.rmtree(file, ignore_errors=True)

    # Deleting all permitted contents in DISK\Users\ACTIVE_USER\AppData\Local\Microsoft\Windows\INetCache\IE
    if cache:
        current_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Local\Microsoft\Windows\INetCache\IE'
        for file in os.listdir(current_path):
            file = os.path.join(current_path, file)
            shutil.rmtree(file, ignore_errors=True)

    # Deleting all permitted contents in DISK\Windows\Temp
    if win_temp:
        current_path = rf'{DISK}\Windows\Temp'
        for file in os.listdir(current_path):
            file = os.path.join(current_path, file)
            shutil.rmtree(file, ignore_errors=True)

    # Deleting contents of recycle bin
    if recycle_bin:
        try:
            subprocess.call(['powershell.exe', 'gci C:\`$recycle.bin -force | remove-item -recurse -force'] , shell=True)
        except:
            pass


def clean_browsers(cookies=True, cache=True, extensions=False, history=False):

    """
    Runs some code to delete browser cookies, cache, extensions and history.

    :param cookies: Whether we should delete cookie files
    :type cookies: bool
    :param cache: Whether we should delete cache files
    :type cache: bool
    :param extensions: Whether we should delete browser extensions
    :type extensions: bool
    :param history: Whether we should delete the browser history
    :type history: bool
    :return: None
    :rtype: None
    """

    # -- Vars --#
    edge = True
    opera = True
    operagx = True
    chrome = True
    firefox = True

    # ---- Sets up paths to all browsers and checks if they exist ---- #
    edge_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Local\Microsoft\Edge\User Data\Default'
    edge_cache_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Local\Microsoft\Edge\User Data\Default\Cache'
    if not os.path.exists(edge_path):
        edge = False

    opera_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Opera Software\Opera Stable'
    opera_cache_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Opera Software\Opera Stable\GPUCache'
    opera_ext_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Opera Software\Opera Stable\Extensions'
    if not os.path.exists(opera_path):
        opera = False

    operagx_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Opera Software\Opera GX Stable'
    operagx_cache_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Opera Software\Opera GX Stable\GPUCache'
    operagx_ext_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Opera Software\Opera GX Stable\Extensions'
    if not os.path.exists(operagx_path):
        operagx = False

    chrome_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Local\Google\Chrome\User Data\Default'
    chrome_cache_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Local\Google\Chrome\User Data\Default\Cache'
    chrome_ext_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Local\Google\Chrome\User Data\Default\Extensions'
    if not os.path.exists(chrome_path):
        chrome = False

    firefox_path = rf'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Mozilla\Firefox\Profiles'
    if not os.path.exists(firefox_path):
        firefox = False
    else:
        firefox_path = fr'{DISK}\Users\{ACTIVE_USER}\AppData\Roaming\Mozilla\Firefox\Profiles\{os.listdir(firefox_path)[0]}'
    firefox_ext_path = rf'{firefox_path}\extensions'

    # ------ REMOVE ALL COOKIES ------#

    if cookies:
        # Edge
        if edge:
            shutil.rmtree(os.path.join(edge_path, 'Cookies'), ignore_errors=True)
            shutil.rmtree(os.path.join(edge_path, 'Cookies-journal'), ignore_errors=True)

        # Opera
        if opera:
            shutil.rmtree(os.path.join(opera_path, 'Cookies'), ignore_errors=True)
            shutil.rmtree(os.path.join(opera_path, 'Cookies-journal'), ignore_errors=True)

        # OperaGX
        if operagx:
            shutil.rmtree(os.path.join(operagx_path, 'Cookies'), ignore_errors=True)
            shutil.rmtree(os.path.join(operagx_path, 'Cookies-journal'), ignore_errors=True)

        # Chrome
        if chrome:
            shutil.rmtree(os.path.join(chrome_path, 'Cookies'), ignore_errors=True)
            shutil.rmtree(os.path.join(chrome_path, 'Cookies-journal'), ignore_errors=True)

        # Firefox
        if firefox:
            shutil.rmtree(os.path.join(firefox_path, 'cookies.sqlite'), ignore_errors=True)

    # ------ REMOVE CACHE ------#
    if cache:
        # Edge
        if edge:
            for file in os.listdir(edge_cache_path):
                shutil.rmtree(os.path.join(edge_cache_path, file), ignore_errors=True)

        # Opera
        if opera:
            for file in os.listdir(opera_cache_path):
                shutil.rmtree(os.path.join(opera_cache_path, file), ignore_errors=True)

        # OperaGX
        if operagx:
            for file in os.listdir(operagx_cache_path):
                shutil.rmtree(os.path.join(operagx_cache_path, file), ignore_errors=True)

        # Chrome
        if chrome:
            for file in os.listdir(chrome_cache_path):
                shutil.rmtree(os.path.join(chrome_cache_path, file), ignore_errors=True)

    # ------ REMOVE EXTENSIONS ------#
    if extensions:

        # Opera
        if opera:
            for file in os.listdir(opera_ext_path):
                shutil.rmtree(os.path.join(opera_ext_path, file), ignore_errors=True)

        # OperaGX
        if operagx:
            for file in os.listdir(operagx_ext_path):
                shutil.rmtree(os.path.join(operagx_ext_path, file), ignore_errors=True)

        # Chrome
        if chrome:
            for file in os.listdir(chrome_ext_path):
                shutil.rmtree(os.path.join(chrome_ext_path, file), ignore_errors=True)

        # Firefox
        if firefox:
            for file in os.listdir(firefox_ext_path):
                shutil.rmtree(os.path.join(firefox_ext_path, file), ignore_errors=True)

    # ------ REMOVE HISTORY ------#
    if history:
        # Edge
        if edge:
            shutil.rmtree(os.path.join(edge_path, 'History'), ignore_errors=True)
            shutil.rmtree(os.path.join(edge_path, 'History-journal'), ignore_errors=True)

        # Opera
        if opera:
            shutil.rmtree(os.path.join(opera_path, 'History'), ignore_errors=True)
            shutil.rmtree(os.path.join(opera_path, 'History-journal'), ignore_errors=True)

        # OperaGX
        if operagx:
            shutil.rmtree(os.path.join(operagx_path, 'History'), ignore_errors=True)
            shutil.rmtree(os.path.join(operagx_path, 'History-journal'), ignore_errors=True)

        # Chrome
        if chrome:
            shutil.rmtree(os.path.join(chrome_path, 'History'), ignore_errors=True)
            shutil.rmtree(os.path.join(chrome_path, 'History-journal'), ignore_errors=True)

        # Firefox
        if firefox:
            shutil.rmtree(os.path.join(firefox_path, 'places.sqlite'), ignore_errors=True)


def pref_for_visual(animation=True, drop_shadow=True, transparency=True):

        """
        Turn on/off visuals for performance.

        :param animation: Whether we should disable animations
        :type animation: bool
        :param drop_shadow: Whether we should disable drop shadow
        :type drop_shadow: bool
        :param transparency: Whether we should disable transparency
        :type transparency: bool
        :return: None
        :rtype: None
        """

        try:
            if drop_shadow:
                win32gui.SystemParametersInfo(win32con.SPI_SETCURSORSHADOW, 0)
                win32gui.SystemParametersInfo(win32con.SPI_SETDROPSHADOW, 0)
            else:
                win32gui.SystemParametersInfo(win32con.SPI_SETCURSORSHADOW, 1)
                win32gui.SystemParametersInfo(win32con.SPI_SETDROPSHADOW, 1)
        except:
            pass

        try:
            if animation:
                win32gui.SystemParametersInfo(win32con.SPI_SETUIEFFECTS, 0)
                win32gui.SystemParametersInfo(win32con.SPI_SETMENUFADE, 0)
                win32gui.SystemParametersInfo(win32con.SPI_SETSELECTIONFADE, 0)
                win32gui.SystemParametersInfo(win32con.SPI_SETTOOLTIPFADE, 0)
                win32gui.SystemParametersInfo(win32con.SPI_SETLISTBOXSMOOTHSCROLLING, 0)
                win32gui.SystemParametersInfo(win32con.SPI_SETANIMATION, 0)
            else:
                win32gui.SystemParametersInfo(win32con.SPI_SETUIEFFECTS, 1)
                win32gui.SystemParametersInfo(win32con.SPI_SETMENUFADE, 1)
                win32gui.SystemParametersInfo(win32con.SPI_SETSELECTIONFADE, 1)
                win32gui.SystemParametersInfo(win32con.SPI_SETTOOLTIPFADE, 1)
                win32gui.SystemParametersInfo(win32con.SPI_SETLISTBOXSMOOTHSCROLLING, 1)
                win32gui.SystemParametersInfo(win32con.SPI_SETANIMATION, 1)
        except:
            pass

        if transparency:
            set_reg('EnableTransparency', '0', r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize')
        else:
            set_reg('EnableTransparency', '1', r'SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize')


def windows_settings(defrag=True, virus_scan=False, maintenance=True, power_plan=True, update=True, clear_ram=True):

    """
    Changes some options with window.

    :param defrag: Whether we should defrag (optimize) the disk
    :type defrag: bool
    :param virus_scan: Whether we should scan for viruses with a quick scan
    :type virus_scan: bool
    :param maintenance: Whether we should turn on windows maintenance
    :type maintenance: bool
    :param power_plan: Whether we should switch power plan to high performance
    :type power_plan: bool
    :param update: Whether we should update windows to the next version
    :type update: bool
    :param clear_ram: Whether we should clear the ram
    :type clear_ram: bool
    :return: None
    :rtype: None
    """

    try:
        if maintenance:
            os.system('MSchedExe.exe start')
        else:
            os.system('MSchedExe.exe stop')
    except:
        pass


    if defrag:
        os.system('defrag /C /H /V /D')

    try:
        if virus_scan:
            os.chdir(rf'{DISK}\ProgramData\Microsoft\Windows Defender\Platform\4.18*')
            os.system('MpCmdRun -SignatureUpdate')
            os.system('MpCmdRun -Scan -ScanType 1')
    except:
        pass

    try:
        p_list = subprocess.check_output('powercfg /list').decode("utf-8")
        p_list = p_list.split('\n')
        p_list = p_list[3:len(p_list) - 1]
        p_tmp = {}
        for i in p_list:
            p_tmp[i[i.find('('):i.find('\r')].replace('(', '').replace(')', '').replace(' *', '')] = i[i.find(
                ': ') + 2:i.find('(') - 2]
        high_pref_id = p_tmp['High performance']
        default_id = p_tmp['Balanced']

        if power_plan:
            os.system(f'powercfg /setactive {high_pref_id}')
        else:
            os.system(f'powercfg /setactive {default_id}')
    except:
        pass

    try:
        if update:
            os.system(rf'{DISK}\Windows\System32\control.exe /name Microsoft.WindowsUpdate')
    except:
        pass

    try:
        if clear_ram:
            set_reg('ClearPageFileAtShutdown', '1', r'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management', winreg.HKEY_LOCAL_MACHINE)
        else:
            set_reg('ClearPageFileAtShutdown', '0', r'SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management', winreg.HKEY_LOCAL_MACHINE)
    except:
        pass


def fix_all(settings_dict):

    """

    :param settings_dict: A dictionary of all the settings that were chosen in the GUI
    :type settings_dict: dict
    :return: The value of the returned value of the function hardware()
    :rtype: float
    """

    clear_disk_cache = settings_dict['disk_cache']
    clear_disk_temp = settings_dict['disk_tmp']
    clear_disk_win_temp = settings_dict['disk_win_tmp']
    clear_recycle_bin = settings_dict['recycle_bin']
    defrag_disk = settings_dict['disk_defrag']
    clear_browser_cookies = settings_dict['browser_cookies']
    clear_browser_cache = settings_dict['browser_cache']
    clear_browser_extensions = settings_dict['browser_extensions']
    clear_browser_history = settings_dict['browser_history']
    disable_transparency = settings_dict['disable_transparency']
    disable_dropshadow = settings_dict['disable_dropshadow']
    disable_animation = settings_dict['disable_animation']
    change_power_plan = settings_dict['power_plan']
    maintenance = settings_dict['maintenance']
    virus_scan = settings_dict['virus_scan']
    update_win = settings_dict['update_win']
    clear_ram = settings_dict['clear_ram']

    # Calls all functions with the chosen arguments
    clean_disk(temp_files=clear_disk_temp, cache=clear_disk_cache, win_temp=clear_disk_win_temp, recycle_bin=clear_recycle_bin)
    clean_browsers(cookies=clear_browser_cookies, cache=clear_browser_cache, extensions=clear_browser_extensions, history=clear_browser_history)
    pref_for_visual(animation=disable_animation, drop_shadow=disable_dropshadow, transparency=disable_transparency)
    windows_settings(defrag=defrag_disk, virus_scan=virus_scan, maintenance=maintenance, power_plan=change_power_plan, update=update_win, clear_ram=clear_ram)