; Soroka's additions to electron-builder's installer, which includes build/installer.nsh by default.

; Windows caches a shortcut's icon by where it points (the file and the index), not by what that file holds now. After an
; update with a new icon, the desktop and Start shortcuts, the taskbar button, Start search and "Programs and Features"
; kept the old picture of Soroka.exe,0 (0.33.3: neither ie4uinit -show nor electron-builder's own SHChangeNotify helped,
; only restarting Explorer with its icon cache deleted, which froze the desktop for a while). So each version points them
; at its own copy of the icon, a path Windows has not cached; the next update's uninstall step deletes the install folder
; and the old copy with it.

; a shortcut that exists (the user may have deleted one) is made again as electron-builder makes it, with that icon
!macro sorokaRelink LINK
  ${if} ${FileExists} "${LINK}"
    CreateShortCut "${LINK}" "$appExe" "" "$INSTDIR\icon-${VERSION}.ico" 0 "" "" "${APP_DESCRIPTION}"
    ClearErrors
    WinShell::SetLnkAUMI "${LINK}" "${APP_ID}"
  ${endIf}
!macroend

!macro customInstall
  File "/oname=$INSTDIR\icon-${VERSION}.ico" "${BUILD_RESOURCES_DIR}\icon.ico"
  !insertmacro sorokaRelink "$newStartMenuLink"
  !insertmacro sorokaRelink "$newDesktopLink"
  !insertmacro sorokaRelink "$APPDATA\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\${SHORTCUT_NAME}.lnk"
  WriteRegStr SHELL_CONTEXT "${UNINSTALL_REGISTRY_KEY}" DisplayIcon "$INSTDIR\icon-${VERSION}.ico"
  System::Call 'Shell32::SHChangeNotify(i 0x8000000, i 0, i 0, i 0)'
!macroend
