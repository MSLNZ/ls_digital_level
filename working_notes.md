{2020-03-18}
 Trying to get LS15_bluetooth.py running on my HP Elitebook running Windows 10.
 This software does currently work on Lucy's Lenovo Thinkpad running Win 7, but I remember we couldn't get it running on just any laptop. It did work on my Lenovo Thinkpad running Win7, but this has recently been rebuilt as a win 10 machine

 `serial.tools.list_ports.comports()` is returning an empty list whether bluetooth is on or off.

 Pyserial should work with win10 see this 
 https://github.com/pyserial/pyserial/issues/451

 which is about the slow speed of this query for bluetooth ports in pyserial version 3.4 vs earlier versions. I've got 3.4 installed in  my root anaconda environment. The issue specifically references the command on Lenovo ThinkPad


 The instructions about pairing in win 7 aren't valid for win 10

This looks promissing

https://superuser.com/questions/1237268/making-a-bluetooth-device-recognized-as-a-com-port

I managed to get my headphones looking  like they're on COM3 both in pythn and in 

Control Panel > Bluetooth Settings > COM Ports Add
```
>>> import serial.tools.list_ports
>>> serial.tools.list_ports.comports()
[<serial.tools.list_ports_common.ListPortInfo object at 0x000002677D5FF708>]
>>> ports = _
>>> ports[0][0]
'COM3'
>>>
```

Next try with LS15

