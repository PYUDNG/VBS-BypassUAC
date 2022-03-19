# VBS-BypassUAC
UAC alert bypass written in vbscript

## BypassUAC.vbs
For VBScript developers. Use `BypassUAC(Host, Hide)` Function in your code as this file provided and you'll get uac without alerts.
Arguments:
1. `Host`: Specify script host to use. 1 for wscript.exe and 2 for cscript.exe
2. `Hide`: Whether run the script with windows hidden. True or False

## BypassUAC.source.vbs
For Other scripts. Just run `BypassUAC.source.vbs [-h|/h] Command1[ Command2 Command3...]` to get your command executed with UAC privilege without alerts.

## BypassUAC.min.vbs
The Same as `BypassUAC.source.vbs`, but the code is minified.
