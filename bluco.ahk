deviceName := "Galaxy Buds2 Pro"

DllCall("LoadLibrary", "str", "Bthprops.cpl", "ptr")
VarSetCapacity(BLUETOOTH_DEVICE_SEARCH_PARAMS, 24+A_PtrSize*2, 0)
NumPut(24+A_PtrSize*2, BLUETOOTH_DEVICE_SEARCH_PARAMS, 0, "uint")
NumPut(1, BLUETOOTH_DEVICE_SEARCH_PARAMS, 4, "uint")   ; fReturnAuthenticated
VarSetCapacity(BLUETOOTH_DEVICE_INFO, 560, 0)
NumPut(560, BLUETOOTH_DEVICE_INFO, 0, "uint")
loop
{
   If (A_Index = 1)
   {
      foundedDevice := DllCall("Bthprops.cpl\BluetoothFindFirstDevice", "ptr", &BLUETOOTH_DEVICE_SEARCH_PARAMS, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr")
      if !foundedDevice
      {
         msgbox no bluetooth devices
         return
      }
   }
   else
   {
      if !DllCall("Bthprops.cpl\BluetoothFindNextDevice", "ptr", foundedDevice, "ptr", &BLUETOOTH_DEVICE_INFO)
      {
         msgbox no found maybw
         break
      }
   }
   dev := StrGet(&BLUETOOTH_DEVICE_INFO+64)
;   msgbox dev
   if (StrGet(&BLUETOOTH_DEVICE_INFO+64) = deviceName)
   {
      VarSetCapacity(Handsfree, 16)
      DllCall("ole32\CLSIDFromString", "wstr", "{0000111e-0000-1000-8000-00805f9b34fb}", "ptr", &Handsfree)   ; https://www.bluetooth.com/specifications/assigned-numbers/service-discovery/
      VarSetCapacity(AudioSink, 16)
      DllCall("ole32\CLSIDFromString", "wstr", "{0000110b-0000-1000-8000-00805f9b34fb}", "ptr", &AudioSink)
   
      hr1 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &Handsfree, "int", 0)   ; voice
      hr2 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &AudioSink, "int", 0)   ; music
     
      if (hr1 = 0) and (hr2 = 0)
         break
      else
      {
         ;DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &Handsfree, "int", 1)   ; voice
         ;DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &AudioSink, "int", 1)   ; music
     
         ;DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &Handsfree, "int", 0)   ; voice
         ;DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &AudioSink, "int", 0)   ; music


         hr1 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &Handsfree, "int", 1) ; voice
         hr2 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &AudioSink, "int", 1) ; music
         ;hr3 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &GenAudServ, "int", 0) ; music
         ;hr4 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &HdstServ, "int", 1) ; music
         ;hr5 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &AVRCTarget, "int", 1) ; music
         ;hr6 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &AVRC, "int", 1) ; music
         ;hr7 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &AVRCController, "int", 0) ; music
         ;hr8 := DllCall("Bthprops.cpl\BluetoothSetServiceState", "ptr", 0, "ptr", &BLUETOOTH_DEVICE_INFO, "ptr", &PnP, "int", 0) ; music
      
         break
      }
   }
}

DllCall("Bthprops.cpl\BluetoothFindDeviceClose", "ptr", foundedDevice)
msgbox done
ExitApp