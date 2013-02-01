set regadd=reg add "HKEY_CURRENT_USER\Control Panel\Desktop"
%regadd% /v TileWallpaper /d 0 /f
%regadd% /v Wallpaper /d "C:\Windows\IloveA3.bmp" /f 
%regadd% /v WallpaperStyle /d 1 /f
RunDll32.exe USER32.DLL,UpdatePerUserSystemParameters
