import os

import PySimpleGUI as sg
import booster as su

# ---- Sets up the tabs for the GUI ---- #

disk_tab = sg.Tab(
    'Disk',
    [
        [sg.Checkbox('Remove disk cache', True, font='Curior 14', key='-DISK_CACHE-')],
        [sg.Checkbox('Remove temporary files', True, font='Curior 14', key='-DISK_TMP-')],
        [sg.Checkbox('Remove windows temporary files', False, font='Curior 14', key='-DISK_WIN_TMP-')],
        [sg.Checkbox('Empty recycle bin', True, font='Curior 14', key='-RECYCLE_BIN-')],
        [sg.Checkbox('Defrag (optimize) disk', True, font='Curior 14', key='-DISK_DEFRAG-')],
    ]
)

browser_tab = sg.Tab(
    'Browsers',
    [
        [sg.Checkbox('Remove cookies', True, font='Curior 14', key='-BROWSER_COOKIES-')],
        [sg.Checkbox('Remove browser cache', True, font='Curior 14', key='-BROWSER_CACHE-')],
        [sg.Checkbox('Remove extensions', False, font='Curior 14', key='-BROWSER_EXTENSIONS-')],
        [sg.Checkbox('Remove history', False, font='Curior 14', key='-BROWSER_HISTORY-')],
    ]
)

performance_tab = sg.Tab(
    'Performance',
    [
        [sg.Checkbox('Disable transparency', True, font='Curior 14', key='-DISABLE_TRANS-')],
        [sg.Checkbox('Disable drop shadows', True, font='Curior 14', key='-DISABLE_DS-')],
        [sg.Checkbox('Disable animations', True, font='Curior 14', key='-DISABLE_ANIM-')],
        [sg.Checkbox('Power plan to high performance', True, font='Curior 14', key='-HIGH_PERF-')],
        [sg.Checkbox('Enable windows maintenance scheduling', True, font='Curior 14', key='-ENABLE_MAINTENANCE-')],
        [sg.Checkbox('Run windows virus scan', False, font='Curior 14')],
        [sg.Checkbox('Update windows', False, font='Curior 14')],
        [sg.Checkbox('Clear RAM', True, font='Curior 14')],
    ]
)

layout = [
    [
        [sg.Button('' ,size=(10,5), pad=(150, 20), image_filename='boost.png', button_color=sg.theme_background_color(), mouseover_colors=sg.theme_background_color())],

        sg.TabGroup(expand_y=True, expand_x=True, layout=[
            [disk_tab, browser_tab, performance_tab]
            ]
        )
    ]
]

window = sg.Window('PC booster', layout=layout, icon='boost.ico', titlebar_icon='boost.ico')

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED:
        break



    if sg.popup_yes_no('Click Yes if you closed all open windows,\nif not all windows will be closed some files won\'t be deleted') == 'No':
        continue

    settings = {
        "disk_cache": values['-DISK_CACHE-'],
        "disk_tmp": values['-DISK_TMP-'],
        "disk_win_tmp": values['-DISK_WIN_TMP-'],
        "disk_defrag": values['-DISK_DEFRAG-'],
        "recycle_bin": values['-RECYCLE_BIN-'],

        "browser_cookies": values['-BROWSER_COOKIES-'],
        "browser_cache": values['-BROWSER_CACHE-'],
        "browser_extensions": values['-BROWSER_EXTENSIONS-'],
        "browser_history": values['-BROWSER_HISTORY-'],

        "disable_transparency": values['-BROWSER_COOKIES-'],
        "disable_dropshadow": values['-BROWSER_CACHE-'],
        "disable_animation": values['-BROWSER_EXTENSIONS-'],
        "power_plan": values['-BROWSER_HISTORY-'],
        "maintenance": values['-BROWSER_HISTORY-'],
        "virus_scan": values['-BROWSER_HISTORY-'],
        "update_win": values['-BROWSER_HISTORY-'],
        "clear_ram": values['-BROWSER_HISTORY-'],
    }

    # Calls the optimization program with all the options the user has chosen in the GUI
    su.fix_all(settings)

    # If the user has chosen to restart the computer it will restart it
    answer = sg.popup_yes_no('Restart computer?')
    if answer == 'Yes':
        os.system("shutdown -t 0 -r -f")

    sg.popup_ok('PC optimized successfully!')

window.close()