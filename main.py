#!/usr/bin/env python
# Module     : SysTrayIcon.py
# Synopsis   : Windows System tray icon.
# Programmer : Simon Brunning - simon@brunningonline.net - modified for Python 3
#              Matthias K. Scharrer - matthias.scharrer@gmail.com
# Date       : 17 December 2022 (11 April 2005, 13 February 2018)
# Notes      : Based on (i.e. ripped off from) Mark Hammond's
#              win32gui_taskbar.py and win32gui_menu.py demos from PyWin32

import os
import sys
import win32api
import win32con
import win32gui_struct
try:
    import winxpgui as win32gui
except ImportError:
    import win32gui
import win32ui

class SysTrayIcon(object):
    '''TODO'''
    QUIT = 'QUIT'
    SPECIAL_ACTIONS = [QUIT]

    FIRST_ID = 1023

    def __init__(self, icon, hover_text, menu_options, on_quit=None, default_menu_index=None, window_class_name=None, ):

        self.icon = icon
        self.hover_text = hover_text
        self.on_quit = on_quit
        
        menu_options = menu_options + (('Quit', 'myIcon_QUIT.ico', self.QUIT),)
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id

        self.default_menu_index = (default_menu_index or 0)
        self.window_class_name = window_class_name or "SysTrayIconPy"

        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart, win32con.WM_DESTROY: self.destroy, win32con.WM_COMMAND: self.command, win32con.WM_USER + 20: self.notify, }
        # Register the Window class.
        window_class = win32gui.WNDCLASS()
        hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = self.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map  # could also specify a wndproc.
        classAtom = win32gui.RegisterClass(window_class)
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(classAtom, self.window_class_name, style, 0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, 0, 0, hinst, None)
        win32gui.UpdateWindow(self.hwnd)
        self.notify_id = None
        self.refresh_icon()

        win32gui.PumpMessages()

    def _add_ids_to_menu_options(self, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                self.menu_actions_by_id.add((self._next_action_id, option_action))
                result.append(menu_option + (self._next_action_id,))
            elif non_string_iterable(option_action):
                result.append((option_text, option_icon, self._add_ids_to_menu_options(option_action), self._next_action_id))
            else:
                print('Unknown item', option_text, option_icon, option_action)
            self._next_action_id += 1
        return result

    def refresh_icon(self):
        # Try and find a custom icon
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(self.icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst, self.icon, win32con.IMAGE_ICON, 0, 0, icon_flags)
        else:
            print("Can't find icon file - using default.")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if self.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd, 0, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP, win32con.WM_USER + 20, hicon, self.hover_text)
        win32gui.Shell_NotifyIcon(message, self.notify_id)

    def restart(self, hwnd, msg, wparam, lparam):
        self.refresh_icon()

    def destroy(self, hwnd, msg, wparam, lparam):
        if self.on_quit: self.on_quit(self)
        nid = (self.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # Terminate the app.

    def notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:
            self.execute_menu_option(self.default_menu_index + self.FIRST_ID)
        elif lparam == win32con.WM_RBUTTONUP:
            self.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:
            pass
        return True

    def show_menu(self):
        menu = win32gui.CreatePopupMenu()
        self.create_menu(menu, self.menu_options)
        # win32gui.SetMenuDefaultItem(menu, 1000, 0)

        pos = win32gui.GetCursorPos()
        # See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/menus_0hdi.asp
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, pos[0], pos[1], 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

    def create_menu(self, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = self.prep_menu_icon(option_icon)

            if option_id in self.menu_actions_by_id:
                item, _extras = win32gui_struct.PackMENUITEMINFO(text=option_text, hbmpItem=option_icon, wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, option_action)
                item, _extras = win32gui_struct.PackMENUITEMINFO(text=option_text, hbmpItem=option_icon, hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

    def prep_menu_icon(self, icon):
        """Load icons into the tray items.
        
        Got from https://stackoverflow.com/a/45890829.
        """
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hIcon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hwndDC = win32gui.GetWindowDC(self.hwnd)
        dc = win32ui.CreateDCFromHandle(hwndDC)
        memDC = dc.CreateCompatibleDC()
        iconBitmap = win32ui.CreateBitmap()
        iconBitmap.CreateCompatibleBitmap(dc, ico_x, ico_y)
        oldBmp = memDC.SelectObject(iconBitmap)
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)

        win32gui.FillRect(memDC.GetSafeHdc(), (0, 0, ico_x, ico_y), brush)
        win32gui.DrawIconEx(memDC.GetSafeHdc(), 0, 0, hIcon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)

        memDC.SelectObject(oldBmp)
        memDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, hwndDC)

        return iconBitmap.GetHandle() 

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)
        return 0
        
    def execute_menu_option(self, id):
        menu_action = self.menu_actions_by_id[id]
        if menu_action == self.QUIT:
            win32gui.DestroyWindow(self.hwnd)
        else:
            menu_action(self)

def non_string_iterable(obj):
    try:
        iter(obj)
    except TypeError:
        return False
    else:
        return not isinstance(obj, str)

if __name__ == '__main__':
    outputfile = r".\prjtrx_events.log"
    import itertools, glob, json, os, datetime
    from functools import partial

    if os.path.exists('projects.json'):
        with open('projects.json','r') as f:
            projects = json.load(f)
    else:
        projects = {'Meta': ['StartUp', 'Break', 'DoSomething']}
        with open('projects.json','w') as f:
            json.dump(projects, f)

    icons = glob.glob('*.ico')
    fallback_icon = 'myIcon_Meta.ico'
    hover_text = "No Project selected!"

    def switchProjectCB(sysTrayIcon, newPrj, newWP, newicon = fallback_icon):
        menow = datetime.datetime.now()
        infotxt = '{} - changed <{}>:<{}>\n'.format(menow.strftime('%Y%m%d %H%M%S'), newPrj, newWP)
        with open(outputfile,'at') as f:
            f.write(infotxt)
        sysTrayIcon.hover_text = infotxt
        if newicon in icons:
            sysTrayIcon.icon = newicon
        else:
            sysTrayIcon.icon = fallback_icon
        sysTrayIcon.refresh_icon()

    menu_options = []
    for prj in projects.keys():
        my_option = prj.split("-")[-1]
        my_option_infotxt = prj.split("-")[0]
        prjIcon = "myIcon_"+my_option+".ico"
        if prjIcon not in icons:
            prjIcon = None
        if len(projects[prj])>1:
            sub_options = []
            for wp in projects[prj]:
                my_sub_option = wp.split("-")[-1]
                my_sub_option_infotxt = wp.split("-")[0]
                subPrjIcon = "myIcon_"+my_sub_option+".ico"
                if subPrjIcon not in icons:
                    subPrjIcon = prjIcon
                my_sub_option_fun = partial(switchProjectCB, newPrj=my_option_infotxt, newWP=my_sub_option_infotxt, newicon=subPrjIcon)
                sub_options += [(my_sub_option, subPrjIcon, my_sub_option_fun,)]
            sub_options = tuple(sub_options)
        elif len(projects[prj])==1:
            my_sub_option = projects[prj][0].split("-")[-1]
            my_sub_option_infotxt = projects[prj][0].split("-")[0]
            if prjIcon is None:
                prjIcon = "myIcon_"+my_sub_option+".ico"
                if prjIcon not in icons:
                    prjIcon = None
            my_sub_option_fun = partial(switchProjectCB, newPrj=my_option_infotxt, newWP=my_sub_option_infotxt, newicon=prjIcon)
            sub_options = my_sub_option_fun
        else:
            sub_options = partial(switchProjectCB, newPrj=my_option_infotxt, newWP='Default', newicon=prjIcon)
        menu_options += [(my_option, prjIcon, sub_options,)]
    menu_options = tuple(menu_options)

    bye = partial(switchProjectCB, newPrj='None', newWP='')

    def _find_default_menu_option(menu_options, startIndex):
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if (option_text == 'StartUp'):
                return startIndex, startIndex
            if callable(option_action):
                pass
            elif non_string_iterable(option_action):
                temp_index, startIndex = _find_default_menu_option(option_action, startIndex)
                if temp_index is not None:
                    return temp_index, temp_index
            else:
                print('Unknown item', option_text, option_icon, option_action)
            startIndex += 1
        return None
    default_menu_index, _dummy = _find_default_menu_option(menu_options, SysTrayIcon.FIRST_ID)

    def bye(sysTrayIcon):
        print('Bye, then.')


    SysTrayIcon(fallback_icon, hover_text, menu_options, on_quit=bye, default_menu_index=1)
