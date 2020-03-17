# ls_digital_level

Read a Leica LS15 digital level via bluetooth and enter in an excel spreadsheet



## Initial Pairing


The LS15 has to
be paired on Bluetooth with the Laptop. This sets up which COM port the LS15
appears on. This only needs to be done once and should not need to be done
again.

1. Turn the laptop and LS15 on. Do the compass calibration on the LS15 if asked. This can be skipped by using the enter key on the LS15, but its worth doing when actually on site to ensure the  angle coordinates are correct.

2. Click on the Bluetooth icon on the laptop status bar and check that bluetooth is on.

3. Click on the Bluetooth icon on the laptop status bar and select Add a device

4. The Level should turn up as LS701020 where 701020 is the serial number of the LS15.

5. Double-click on the LS701020 icon on the laptop to pair with the LS15. This will involve entering a pin (0000) on the laptop and confirming the connection on the LS15 instrument panel.

6. Right-click on the icon in the Bluetooth devices window, Properties -> Services to find which COM port the level has been assigned to.

## 

## Running LS15 Reader software

1. Turn the laptop and LS15 on. Do the compass calibration on the LS15 if asked

2. Start the LS15 Reader software from the icon on the PC desktop or by running 
   
   `python LS15_bluetooth.py`in a terminal.

3. Select the COM port from the dropdown list in the application GUI. If you don’t know it see step 6 above. If the right COM port is not on the drop down restart the LS15 Reader software.

4. Click on Open Serial Port in the LS15 Reader window. A window should appear above the Bluetooth icon on the status bar on the laptop, click on this and enter the pin (0000). Confirm the connection on the LS15 by selecting 'OK' on the Level

5. The software GUI on the laptop should show “Bluetooth connected” and the Bluetooth icon on the LS15 should not have an exclamation mark

6. Enter the Q-Level program on the LS15. This is required, the software does not work with the basic level program. Pressing the red trigger button on th right
   hand side of the level should take a reading and transfer it to the laptop. The
   Height, Distance and bearing should show up in the software window. Ignore what is happening with the Level program.

7. To transfer the data to Excel, save and close the relevant spreadsheet (if open).
   Open it again from the button on the LS15 software. Now when the red trigger
   button on the LS15 is pressed the data should be transferred to both to the
   software window and the active cell in Excel.

8. All strings received from the LS15 are recorded in a file LS15_YYYYMMDD_HHSS.log
   (on the desktop) and all valid measurements are stored in csv format

Timestamp, height, distance, X, Y, Z 

in the file LS15_YYYYMMDD_HHMM.csv (in the desktop folder of the laptop)



## Settings on the LS15

The Level should have the following settings, these are persistent when the level is turned off and on.

* Interface -> Bluetooth

* Data -> Output -> interface

* Data -> Output -> GSI16

* Work -> Trigger Key -> AF+DIST+REC
  

If required 

* Mode -> Mean
