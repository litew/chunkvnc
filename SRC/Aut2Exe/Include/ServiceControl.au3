Global $STANDARD_RIGHTS_REQUIRED = 0x000F0000

; Service Control Manager access types
Global $SC_MANAGER_CONNECT = 0x0001
Global $SC_MANAGER_CREATE_SERVICE = 0x0002
Global $SC_MANAGER_ENUMERATE_SERVICE = 0x0004
Global $SC_MANAGER_LOCK = 0x0008
Global $SC_MANAGER_QUERY_LOCK_STATUS = 0x0010
Global $SC_MANAGER_MODIFY_BOOT_CONFIG = 0x0020

Global $SC_MANAGER_ALL_ACCESS = BitOR($STANDARD_RIGHTS_REQUIRED, _
                                      $SC_MANAGER_CONNECT, _
                                      $SC_MANAGER_CREATE_SERVICE, _
                                      $SC_MANAGER_ENUMERATE_SERVICE, _
                                      $SC_MANAGER_LOCK, _
                                      $SC_MANAGER_QUERY_LOCK_STATUS, _
                                      $SC_MANAGER_MODIFY_BOOT_CONFIG)

; Service access types
Global $SERVICE_QUERY_CONFIG = 0x0001
Global $SERVICE_CHANGE_CONFIG = 0x0002
Global $SERVICE_QUERY_STATUS = 0x0004
Global $SERVICE_ENUMERATE_DEPENDENTS = 0x0008
Global $SERVICE_START = 0x0010
Global $SERVICE_STOP = 0x0020
Global $SERVICE_PAUSE_CONTINUE = 0x0040
Global $SERVICE_INTERROGATE = 0x0080
Global $SERVICE_USER_DEFINED_CONTROL = 0x0100

Global $SERVICE_ALL_ACCESS = BitOR($STANDARD_RIGHTS_REQUIRED, _
                                   $SERVICE_QUERY_CONFIG, _
                                   $SERVICE_CHANGE_CONFIG, _
                                   $SERVICE_QUERY_STATUS, _
                                   $SERVICE_ENUMERATE_DEPENDENTS, _
                                   $SERVICE_START, _
                                   $SERVICE_STOP, _
                                   $SERVICE_PAUSE_CONTINUE, _
                                   $SERVICE_INTERROGATE, _
                                   $SERVICE_USER_DEFINED_CONTROL)

; Service controls
Global $SERVICE_CONTROL_STOP = 0x00000001
Global $SERVICE_CONTROL_PAUSE = 0x00000002
Global $SERVICE_CONTROL_CONTINUE = 0x00000003
Global $SERVICE_CONTROL_INTERROGATE = 0x00000004
Global $SERVICE_CONTROL_SHUTDOWN = 0x00000005
Global $SERVICE_CONTROL_PARAMCHANGE = 0x00000006
Global $SERVICE_CONTROL_NETBINDADD = 0x00000007
Global $SERVICE_CONTROL_NETBINDREMOVE = 0x00000008
Global $SERVICE_CONTROL_NETBINDENABLE = 0x00000009
Global $SERVICE_CONTROL_NETBINDDISABLE = 0x0000000A
Global $SERVICE_CONTROL_DEVICEEVENT = 0x0000000B
Global $SERVICE_CONTROL_HARDWAREPROFILECHANGE = 0x0000000C
Global $SERVICE_CONTROL_POWEREVENT = 0x0000000D
Global $SERVICE_CONTROL_SESSIONCHANGE = 0x0000000E

; Service types
Global $SERVICE_KERNEL_DRIVER = 0x00000001
Global $SERVICE_FILE_SYSTEM_DRIVER = 0x00000002
Global $SERVICE_ADAPTER = 0x00000004
Global $SERVICE_RECOGNIZER_DRIVER = 0x00000008
Global $SERVICE_DRIVER = BitOR($SERVICE_KERNEL_DRIVER, _
                               $SERVICE_FILE_SYSTEM_DRIVER, _
                               $SERVICE_RECOGNIZER_DRIVER)
Global $SERVICE_WIN32_OWN_PROCESS = 0x00000010
Global $SERVICE_WIN32_SHARE_PROCESS = 0x00000020
Global $SERVICE_WIN32 = BitOR($SERVICE_WIN32_OWN_PROCESS, _
                              $SERVICE_WIN32_SHARE_PROCESS)
Global $SERVICE_INTERACTIVE_PROCESS = 0x00000100
Global $SERVICE_TYPE_ALL = BitOR($SERVICE_WIN32, _
                                 $SERVICE_ADAPTER, _
                                 $SERVICE_DRIVER, _
                                 $SERVICE_INTERACTIVE_PROCESS)

; Service start types
Global $SERVICE_BOOT_START = 0x00000000
Global $SERVICE_SYSTEM_START = 0x00000001
Global $SERVICE_AUTO_START = 0x00000002
Global $SERVICE_DEMAND_START = 0x00000003
Global $SERVICE_DISABLED = 0x00000004

; Service error control
Global $SERVICE_ERROR_IGNORE = 0x00000000
Global $SERVICE_ERROR_NORMAL = 0x00000001
Global $SERVICE_ERROR_SEVERE = 0x00000002
Global $SERVICE_ERROR_CRITICAL = 0x00000003

;===============================================================================
; Description:   Starts a service on a computer
; Parameters:    $sComputerName - name of the target computer. If empty, the local computer name is used
;                $sServiceName - name of the service to start
; Requirements:  None
; Return Values: On Success - 1
;                On Failure - 0 and @error is set to extended Windows error code
; Note:          This function does not check to see if the service has started successfully
;===============================================================================
Func _StartService($sComputerName, $sServiceName)
   Local $hAdvapi32
   Local $hKernel32
   Local $arRet
   Local $hSC
   Local $hService
   Local $lError = -1

   $hAdvapi32 = DllOpen("advapi32.dll")
   If $hAdvapi32 = -1 Then Return 0
   $hKernel32 = DllOpen("kernel32.dll")
   If $hKernel32 = -1 Then Return 0
   $arRet = DllCall($hAdvapi32, "long", "OpenSCManager", _
                    "str", $sComputerName, _
                    "str", "ServicesActive", _ 
                    "long", $SC_MANAGER_CONNECT)
   If $arRet[0] = 0 Then
      $arRet = DllCall($hKernel32, "long", "GetLastError")
      $lError = $arRet[0]
   Else
      $hSC = $arRet[0]
      $arRet = DllCall($hAdvapi32, "long", "OpenService", _
                       "long", $hSC, _
                       "str", $sServiceName, _
                       "long", $SERVICE_START)
      If $arRet[0] = 0 Then
         $arRet = DllCall($hKernel32, "long", "GetLastError")
         $lError = $arRet[0]
      Else
         $hService = $arRet[0]
         $arRet = DllCall($hAdvapi32, "int", "StartService", _
                          "long", $hService, _
                          "long", 0, _
                          "str", "")
         If $arRet[0] = 0 Then
            $arRet = DllCall($hKernel32, "long", "GetLastError")
            $lError = $arRet[0]
         EndIf
         DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hService)         
      EndIf
      DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hSC)
   EndIf
   DllClose($hAdvapi32)
   DllClose($hKernel32)
   If $lError <> -1 Then 
      SetError($lError)
      Return 0
   EndIf
   Return 1
EndFunc

;===============================================================================
; Description:   Stops a service on a computer
; Parameters:    $sComputerName - name of the target computer. If empty, the local computer name is used
;                $sServiceName - name of the service to stop
; Requirements:  None
; Return Values: On Success - 1
;                On Failure - 0 and @error is set to extended Windows error code
; Note:          This function does not check to see if the service has stopped successfully
;===============================================================================
Func _StopService($sComputerName, $sServiceName)
   Local $hAdvapi32
   Local $hKernel32
   Local $arRet
   Local $hSC
   Local $hService
   Local $lError = -1

   $hAdvapi32 = DllOpen("advapi32.dll")
   If $hAdvapi32 = -1 Then Return 0
   $hKernel32 = DllOpen("kernel32.dll")
   If $hKernel32 = -1 Then Return 0
   $arRet = DllCall($hAdvapi32, "long", "OpenSCManager", _
                    "str", $sComputerName, _
                    "str", "ServicesActive", _
                    "long", $SC_MANAGER_CONNECT)
   If $arRet[0] = 0 Then
      $arRet = DllCall($hKernel32, "long", "GetLastError")
      $lError = $arRet[0]
   Else
      $hSC = $arRet[0]
      $arRet = DllCall($hAdvapi32, "long", "OpenService", _
                       "long", $hSC, _
                       "str", $sServiceName, _
                       "long", $SERVICE_STOP)
      If $arRet[0] = 0 Then
         $arRet = DllCall($hKernel32, "long", "GetLastError")
         $lError = $arRet[0]
      Else
         $hService = $arRet[0]
         $arRet = DllCall($hAdvapi32, "int", "ControlService", _
                          "long", $hService, _
                          "long", $SERVICE_CONTROL_STOP, _
                          "str", "")
         If $arRet[0] = 0 Then
            $arRet = DllCall($hKernel32, "long", "GetLastError")
            $lError = $arRet[0]
         EndIf
         DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hService)         
      EndIf
      DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hSC)
   EndIf
   DllClose($hAdvapi32)
   DllClose($hKernel32)   
   If $lError <> -1 Then 
      SetError($lError)
      Return 0
   EndIf
   Return 1
EndFunc

;===============================================================================
; Description:   Checks if a service exists on a computer
; Parameters:    $sComputerName - name of the target computer. If empty, the local computer name is used
;                $sServiceName - name of the service to check
; Requirements:  None
; Return Values: On Success - 1
;                On Failure - 0
;===============================================================================
Func _ServiceExists($sComputerName, $sServiceName)
   Local $hAdvapi32
   Local $arRet
   Local $hSC
   Local $bExist = 0

   $hAdvapi32 = DllOpen("advapi32.dll")
   If $hAdvapi32 = -1 Then Return 0
   $arRet = DllCall($hAdvapi32, "long", "OpenSCManager", _
                    "str", $sComputerName, _
                    "str", "ServicesActive", _
                    "long", $SC_MANAGER_CONNECT)
   If $arRet[0] <> 0 Then
      $hSC = $arRet[0]
      $arRet = DllCall($hAdvapi32, "long", "OpenService", _
                       "long", $hSC, _
                       "str", $sServiceName, _
                       "long", $SERVICE_INTERROGATE)
      If $arRet[0] <> 0 Then
         $bExist = 1
         DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $arRet[0])
      EndIf
      DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hSC)
   EndIf
   DllClose($hAdvapi32)
   Return $bExist
EndFunc

;===============================================================================
; Description:   Checks if a service is running on a computer
; Parameters:    $sComputerName - name of the target computer. If empty, the local computer name is used
;                $sServiceName - name of the service to check
; Requirements:  None
; Return Values: On Success - 1
;                On Failure - 0
; Note:          This function relies on the fact that only a running service responds
;                to a SERVICE_CONTROL_INTERROGATE control code. Check the ControlService
;                page on MSDN for limitations with using this method.
;===============================================================================
Func _ServiceRunning($sComputerName, $sServiceName)
   Local $hAdvapi32
   Local $arRet
   Local $hSC
   Local $hService   
   Local $bRunning = 0

   $hAdvapi32 = DllOpen("advapi32.dll")
   If $hAdvapi32 = -1 Then Return 0
   $arRet = DllCall($hAdvapi32, "long", "OpenSCManager", _
                    "str", $sComputerName, _
                    "str", "ServicesActive", _
                    "long", $SC_MANAGER_CONNECT)
   If $arRet[0] <> 0 Then
      $hSC = $arRet[0]
      $arRet = DllCall($hAdvapi32, "long", "OpenService", _
                       "long", $hSC, _
                       "str", $sServiceName, _
                       "long", $SERVICE_INTERROGATE)
      If $arRet[0] <> 0 Then
         $hService = $arRet[0]
         $arRet = DllCall($hAdvapi32, "int", "ControlService", _
                          "long", $hService, _
                          "long", $SERVICE_CONTROL_INTERROGATE, _
                          "str", "")
         $bRunning = $arRet[0]
         DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hService)
      EndIf
      DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hSC)
   EndIf
   DllClose($hAdvapi32)   
   Return $bRunning
EndFunc

;===============================================================================
; Description:   Creates a service on a computer
; Parameters:    $sComputerName - name of the target computer. If empty, the local computer name is used
;                $sServiceName - name of the service to create
;                $sDisplayName - display name of the service
;                $sBinaryPath - fully qualified path to the service binary file
;                               The path can also include arguments for an auto-start service
;                $sServiceUser - [optional] default is LocalSystem
;                                name of the account under which the service should run
;                $sPassword - [optional] default is empty
;                             password to the account name specified by $sServiceUser
;                             Specify an empty string if the account has no password or if the service 
;                             runs in the LocalService, NetworkService, or LocalSystem account
;                 $nServiceType - [optional] default is $SERVICE_WIN32_OWN_PROCESS
;                 $nStartType - [optional] default is $SERVICE_AUTO_START
;                 $nErrorType - [optional] default is $SERVICE_ERROR_NORMAL
;                 $nDesiredAccess - [optional] default is $SERVICE_ALL_ACCESS
;                 $sLoadOrderGroup - [optional] default is empty
;                                    names the load ordering group of which this service is a member
; Requirements:  Administrative rights on the computer
; Return Values: On Success - 1
;                On Failure - 0 and @error is set to extended Windows error code
; Note:          Dependencies cannot be specified using this function
;                Refer to the CreateService page on MSDN for more information
;===============================================================================
Func _CreateService($sComputerName, _
                    $sServiceName, _
                    $sDisplayName, _
                    $sBinaryPath, _
                    $sServiceUser = "LocalSystem", _
                    $sPassword = "", _
                    $nServiceType = 0x00000010, _
                    $nStartType = 0x00000002, _
                    $nErrorType = 0x00000001, _
                    $nDesiredAccess = 0x000f01ff, _
                    $sLoadOrderGroup = "")
   Local $hAdvapi32
   Local $hKernel32
   Local $arRet
   Local $hSC
   Local $lError = -1   

   $hAdvapi32 = DllOpen("advapi32.dll")
   If $hAdvapi32 = -1 Then Return 0
   $hKernel32 = DllOpen("kernel32.dll")
   If $hKernel32 = -1 Then Return 0
   $arRet = DllCall($hAdvapi32, "long", "OpenSCManager", _
                    "str", $sComputerName, _
                    "str", "ServicesActive", _
                    "long", $SC_MANAGER_ALL_ACCESS)
   If $arRet[0] = 0 Then
      $arRet = DllCall($hKernel32, "long", "GetLastError")
      $lError = $arRet[0]
   Else
      $hSC = $arRet[0]
      $arRet = DllCall($hAdvapi32, "long", "OpenService", _
                       "long", $hSC, _
                       "str", $sServiceName, _
                       "long", $SERVICE_INTERROGATE)
      If $arRet[0] = 0 Then
         $arRet = DllCall($hAdvapi32, "long", "CreateService", _
                          "long", $hSC, _
                          "str", $sServiceName, _
                          "str", $sDisplayName, _
                          "long", $nDesiredAccess, _
                          "long", $nServiceType, _
                          "long", $nStartType, _
                          "long", $nErrorType, _
                          "str", $sBinaryPath, _
                          "str", $sLoadOrderGroup, _
                          "ptr", 0, _
                          "str", "", _
                          "str", $sServiceUser, _
                          "str", $sPassword)
         If $arRet[0] = 0 Then            
            $arRet = DllCall($hKernel32, "long", "GetLastError")
            $lError = $arRet[0]
         Else
            DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $arRet[0])
         EndIf
      Else
         DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $arRet[0])
      EndIf      
      DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hSC)
   EndIf
   DllClose($hAdvapi32)
   DllClose($hKernel32)   
   If $lError <> -1 Then 
      SetError($lError)
      Return 0
   EndIf
   Return 1
EndFunc

;===============================================================================
; Description:   Deletes a service on a computer
; Parameters:    $sComputerName - name of the target computer. If empty, the local computer name is used
;                $sServiceName - name of the service to delete
; Requirements:  Administrative rights on the computer
; Return Values: On Success - 1
;                On Failure - 0 and @error is set to extended Windows error code
;===============================================================================
Func _DeleteService($sComputerName, $sServiceName)
   Local $hAdvapi32
   Local $hKernel32
   Local $arRet
   Local $hSC
   Local $hService
   Local $lError = -1   

   $hAdvapi32 = DllOpen("advapi32.dll")
   If $hAdvapi32 = -1 Then Return 0
   $hKernel32 = DllOpen("kernel32.dll")
   If $hKernel32 = -1 Then Return 0
   $arRet = DllCall($hAdvapi32, "long", "OpenSCManager", _
                    "str", $sComputerName, _
                    "str", "ServicesActive", _
                    "long", $SC_MANAGER_ALL_ACCESS)
   If $arRet[0] = 0 Then
      $arRet = DllCall($hKernel32, "long", "GetLastError")
      $lError = $arRet[0]
   Else
      $hSC = $arRet[0]
      $arRet = DllCall($hAdvapi32, "long", "OpenService", _
                       "long", $hSC, _
                       "str", $sServiceName, _
                       "long", $SERVICE_ALL_ACCESS)
      If $arRet[0] = 0 Then
         $arRet = DllCall($hKernel32, "long", "GetLastError")
         $lError = $arRet[0]
      Else
         $hService = $arRet[0]
         $arRet = DllCall($hAdvapi32, "int", "DeleteService", _
                          "long", $hService)
         If $arRet[0] = 0 Then
            $arRet = DllCall($hKernel32, "long", "GetLastError")
            $lError = $arRet[0]
         EndIf
         DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hService)
      EndIf
      DllCall($hAdvapi32, "int", "CloseServiceHandle", "long", $hSC)
   EndIf
   DllClose($hAdvapi32)
   DllClose($hKernel32)   
   If $lError <> -1 Then 
      SetError($lError)
      Return 0
   EndIf
   Return 1
EndFunc
