# Turn display off by calling WindowsAPI.
 
# SendMessage(HWND_BROADCAST,WM_SYSCOMMAND, SC_MONITORPOWER, POWER_OFF)
# HWND_BROADCAST  0xffff
# WM_SYSCOMMAND   0x0112
# SC_MONITORPOWER 0xf170
# POWER_OFF       0x0002

. ..\..\Scripts\PS\Init.ps1

Count 5
Switch-DisplayOff