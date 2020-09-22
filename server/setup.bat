@echo Longbow Server
@echo ==============
@echo Copying DLLS..
copy ImageUtils.dll c:\windows\system\ImageUtils.dll
copy Base64.dll c:\windows\system\base64.dll
copy Base64.dep c:\windows\system\base64.dep
@echo Registering ActiveX DLL..
c:\windows\system\regsvr32 c:\windows\system\base64.dll
@echo Complete.