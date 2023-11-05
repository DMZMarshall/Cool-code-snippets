#System Tray
import glob
import itertools
import win32api
import win32con
import win32gui_struct
import os
import sys

try:
    import winxpgui as win32gui
except ImportError:
    import win32gui
    
#System Notifications
from win32gui import *
from uuid import uuid4  
#####################################################
# System Tray Class

class SysTrayIcon(object):
    '''TODO'''
    QUIT = 'Close Service'
    SPECIAL_ACTIONS = [QUIT]
    
    FIRST_ID = 1023
    
    def __init__(self,icon,hover_text,menu_options,on_quit=None,default_menu_index=None,window_class_name=None,):
        
        self.icon = icon
        self.hover_text = hover_text
        self.on_quit = on_quit
        
        #Guarantee a way close program from tray separate from other menu options
        menu_options = menu_options + (('Close Service', None, self.QUIT),)
        
        #Set UID's for the menu options
        self._next_action_id = self.FIRST_ID
        self.menu_actions_by_id = set()
        self.menu_options = self._add_ids_to_menu_options(list(menu_options))
        self.menu_actions_by_id = dict(self.menu_actions_by_id)
        del self._next_action_id
        
        #Default menu option highlighted
        self.default_menu_index = (default_menu_index or 0)
        
        #Internal tray app name
        self.window_class_name = window_class_name or "SysTrayIconPy"
        
        #Core tray functions
        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): self.restart,
                       win32con.WM_DESTROY: self.destroy,
                       win32con.WM_COMMAND: self.command,
                       win32con.WM_USER+20 : self.notify,}

        # Register the Window class.
        window_class = win32gui.WNDCLASS()
        hinst = window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = self.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map
        class_registered = win32gui.RegisterClass(window_class)
        
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(class_registered, self.window_class_name,style,
                                          0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT,
                                          0, 0, hinst, None)
        win32gui.UpdateWindow(self.hwnd)
        self.notify_id = None
        self.refresh_icon()
        
        #Continuously send waiting notifications until a WM_Quit notif
        win32gui.PumpMessages()

    #Setting menu options and nesting submenu options recursively
    def _add_ids_to_menu_options(self, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            
            #Populates system tray options and submenu options
            if callable(option_action) or option_action in self.SPECIAL_ACTIONS:
                
                #Force appends the close application option independant of others
                self.menu_actions_by_id.add((self._next_action_id, option_action))
                result.append(menu_option + (self._next_action_id,))
            elif non_string_iterable(option_action):
                
                #Populates the system tray options list
                result.append((option_text, option_icon,
                               self._add_ids_to_menu_options(option_action),
                               self._next_action_id))
            else:
                #Force action if an error with menu option formatting
                print('Unknown item', option_text, option_icon, option_action)
            
            #Option override prevention on the next loop
            self._next_action_id += 1
        return result
        
    def refresh_icon(self):
        # Try and find a custom icon
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(self.icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst, self.icon, win32con.IMAGE_ICON,
                                       0, 0, icon_flags)
        else:
            print("Can't find icon file - using default.")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if self.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        self.notify_id = (self.hwnd,
                          0, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,win32con.WM_USER+20,
                          hicon, self.hover_text)
        
        #Notify system of icon refresh
        win32gui.Shell_NotifyIcon(message, self.notify_id)
    
    #Reload tray icon    
    def restart(self, hwnd, msg, wparam, lparam):
        self.refresh_icon()

    def destroy(self, hwnd, msg, wparam, lparam):
        #Full termination of app
        if self.on_quit:
            self.on_quit(self)
        nid = (self.hwnd, 0)
        
        #Notify OS of app termination
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)
        
    #Denotes left or right mouse button clicks
    def notify(self, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:
            self.execute_menu_option(self.default_menu_index + self.FIRST_ID)
        elif lparam == win32con.WM_RBUTTONUP:
            self.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:
            pass
        return True
    
    #Right click function   
    def show_menu(self):
        #Menu creation
        menu = win32gui.CreatePopupMenu()
        self.create_menu(menu, self.menu_options)
        
        #Cursor tracking
        #Highlights menu options on hover
        pos = win32gui.GetCursorPos()
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN,
                                pos[0], pos[1], 0,
                                self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL,0, 0)
    
    def create_menu(self, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = self.prep_menu_icon(option_icon)
            
            if option_id in self.menu_actions_by_id:
                item, extras = win32gui_struct.PackMENUITEMINFO(text = option_text,
                                                                hbmpItem = option_icon,
                                                                wID = option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                self.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text = option_text,
                                                                hbmpItem = option_icon,
                                                                hSubMenu = submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)
    
    #System tray Icon
    def prep_menu_icon(self, icon):
        # Load the icon.
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hdcBitmap = win32gui.CreateCompatibleDC(0)
        hdcScreen = win32gui.GetDC(0)
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)

        # Fill the background.
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
        # unclear if brush needs to be feed.  Best clue I can find is:
        # "GetSysColorBrush returns a cached brush instead of allocating a new one." 
        # implies no DeleteObject

        # draw the icon
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        win32gui.SelectObject(hdcBitmap, hbmOld)
        win32gui.DeleteDC(hdcBitmap)
        
        return hbm

    def command(self, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        self.execute_menu_option(id)
        
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

#####################################################
# System Tray notification 

class WindowsBalloonTip:
    def __init__(self, title, msg):
        message_map = {win32con.WM_DESTROY: self.OnDestroy,}

        # Register the Window class.
        wc = WNDCLASS()
        hinst = wc.hInstance = GetModuleHandle(None)
        
        #Unique uuid for multiple instances of system notifications
        wc.lpszClassName = str(uuid4())
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        classNotif = RegisterClass(wc)

        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( classNotif, "Taskbar", style,
                                 0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT,
                                 0, 0, hinst, None)
        
        UpdateWindow(self.hwnd)
        iconPathName = os.path.abspath(os.path.join( sys.path[0], "balloontip.ico" ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        
        try:
            hicon = LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
            hicon = LoadIcon(0, win32con.IDI_APPLICATION)
            flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
            nid = (self.hwnd, 0, flags, hicon)
            
        #Add a handler for balloon tooltip
        Shell_NotifyIcon(NIM_ADD,nid)
        Shell_NotifyIcon(NIM_MODIFY, (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20,
                                      hicon, "Balloon  tooltip", msg, 200, title))
        
        #Delete handler for the ballon tooltip from system tray to clear memory
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        
    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0)
        return DefWindowProc(hwnd, msg, wparam, lparam)
        #Destroy balloon notif

def balloon_tip(title, msg):
    w=WindowsBalloonTip(title, msg)
    
if __name__ == '__main__':
    icons = itertools.cycle(glob.glob('*.ico'))
    hover_text = "System tray"
    def bye(self): 
    
    #Closes application and cleans up terminal
        try:
            sys.exit()
        except:
            #Catching all non-exiting exceptions, or when something goes really wrong where sys.exit() fails
            e=sys.exc_info()[0]
    #Populate System Tray
    menu_options = (('Settings', next(icons) ))
    
    SysTrayIcon(next(icons), hover_text, menu_options, on_quit=bye, default_menu_index=1)