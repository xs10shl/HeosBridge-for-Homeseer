# HeosBridge-for-Homeseer

This is a script I use to connect Heos and my home audio system (a legacy NuVo Grand Concerto) together via Homeseer HS3. It's comprised of three scripts, a config file, and a folder with some default images.  I use these scripts daily, but please use at your own risk, as I can't guarantee functionality on your system.  Also, since the scripts will alter your HS database, please make a backup of your system prior to installing and using.

As of today, this is a single-Heos-player monitor system which knows how to connect to a Denon or Marantz system only.   Also, it does not implement the entire complement of Heos commands.  I only implemented the ones I needed for NuVo menuing to work the way I wanted it to. 

To use: 
1) Copy the HeosBridge folder into your HS3/html/Images directory, so the final path to the images is hs3/html/images/HeosBridge/
2) Copy the 3 .vb scripts into your HS3/scripts directory.
3) Copy the HeosBridge.ini file into the HS3/Config directory.
4) Edit HeosBridge.ini with the correct IP address of your Heos system.
5) Run the initHeosBridge.vb script by creating an event to manually run a script.  This will create the devices needed for the HeosBridge to work.
6) Add the following line into your startup.vb file : hs.runscriptfunc("HeosButtons.vb", "HeosInitialize", "STARTUP", False, False)
7) Restart Homeseer, and check the log file for and startup errors in the script.
8) Press the "connect" button on the HeosBridge Root device to test functionality.  

This script works as follows.  After initHeosBridge.vb runs, it is no longer needed until you want to make changes to the HeosBridge devices by editing the file and re-running it.  Running it again will update any relevant devices based on your script updates, and keep the same dvRefs, in case you are referencing your devices using the dvRef method.  The scripts all usee the 'AddressExists' method of lookups, so you can alter the device names and dvRefs and everything should still work. The Heos Devices are linked with the script HeosButtons.vb on creation, which acts as the function code behind each buttonpress. HeosButtons also contains the code for sending requests to Heos, and returning results from Browse requests.  The third script HeosMonitor.vb is a looping script which constantly monitors a separate telnet session to Heos to retrieve status.  It's unfortunately a sycnronous script, but I tried to make it play as nice with system resources as I could, given this limitation.  I personally use 3 of these looping scripts on one homeseer installation with no noticable loss of processing power, so I'm not planning on re-writing it as an async solution at this point. 

Of note, I believe I've stripped out all the specific code that I use to interface with my current NuVo system, but I have not tried to install it on a virgin system, so there may be two or three unresolved globals or references that I missed.  Let me know if you find anything like that, and I'll make changes to the code to eliminate them.
