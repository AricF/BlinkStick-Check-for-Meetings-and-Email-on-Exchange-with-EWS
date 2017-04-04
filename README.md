# BlinkStick-Check-for-Meetings-and-Email-on-Exchange-with-EWS

BlinkStick notifications on email and meeting alerts from Exchange Server using EWS.

###### Description
This is a small console app that uses Exchange Web Services (EWS) to check for upcoming meetings and new email messages.

### Configuration
Most of the configuration is done in the app.config file.  If you don't specify a username or password in that file it will ask you for them on launch.  You _**will need**_ to specify the url for the EWS for your installation of Exchange Server in the app.config.  For security reasons you may not want to save your password in the app.config file.

###Dependencies (already included in this repository)
This is dependent on 3 other projects I checked in so you can start running immediately.  They are:
1.  LibUsbDotNet
2.  HidSharp
3.  BlinkStickDotNet

I did this to follow arvydas' example in this project: https://github.com/arvydas/BlinkStickDotNet/tree/master/BlinkStickDotNet. By doing this it makes it easier for someone using Mono to simply grab the project and run.

###Usage
When the console app is running, pressing R (see comment below) will reset the new email count.  Pressing Esc will exit the app.

###Hardware Tested On
This was tested against a BlinkStick Nano in a Belkin powered USB hub. https://www.blinkstick.com/products/blinkstick-nano. (Note that I had trouble getting the Nano to show up as a USB device if I plugged it directly into the computer and had to use a hub to get it to work correctly.  This appears to be a common issue.)

###Exchange Server Compatibility
This is compatible with Exchange 2007 SP1 and higher.  I purposely did not use newer calls for compatibility with older Exchange servers.


###Known Issues:
I have worked to make sure this works in Mono and not just native .NET on Windows.  However at the end of running it doesn't seem to nicely close down all the threads created in some of the supporting USB libraries and you may have to manually kill the console session.  If you know how to get that to work in Mono, I'd love to hear it.  

###Comment about R Command:
There is one implementation change I'd like to make.  In determining if you have new unread email, the app counts the number of unread messages on startup.  Then whatever frequency you set the app to check again (default = 5 minutes) it looks to see if that count has increased.  If it has, it assumes you have new email.  This logic starts falling apart when you skip an unimportant message and leave it as unread even if you've read newer messages and therefore don't really have new email.  I'd prefer to do a check where it simply looks to see if the most recent email is marked as unread or not.  However, I haven't come up with the XML to pull that message out and examine it's unread status yet.  If you are an EWS expert I'd love your help to accomplish this change in logic.

