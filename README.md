# BlinkStick-Check-for-Meetings-and-Email-on-Exchange-with-EWS
BlinkStick notifications on email and meeting alerts from Exchange Server using EWS.

This is a small console app that uses Exchange Web Services (EWS) to check for upcoming meetings and new email messages.

Most of the configuration is done in the app.config file.  If you don't specify a username or password in that file it will ask you for them on launch.  You **will need** to specify the url for the EWS for your installation of Exchange Server in the app.config.  For security reasons you may not want to save your password in the app.config file.

This is dependent on 3 other projects I checked in so you can start running immediately.  They are:
1.  LibUsbDotNet
2.  HidSharp
3.  BlinkStickDotNet

When the console app is running, pressing R (see comment below) will reset the new email count.  Pressing Esc will exit the app.

This is compatible with Exchange 2007 SP1 and higher.  I purposely did not use newer calls for compatibility with older Exchange servers.

I have worked to make sure this works in Mono and not just native .NET on Windows.  

Comment about R Command:
There is one issue I'd still like to address.  In determining if you have new unread email, the app counts the number of unread messages.  Then whatever frequency you set the app to check again (default = 5 minutes) it looks to see if that count has increased.  If it has, it assumes you have new email.  This starts falling apart when you skip an unimportant message and leave it as unread even if you've read newer messages and therefore don't really have new email.  I'd prefer to do a check where it simply looks to see if the most recent email is marked as unread or not.  However, I haven't come up with the XML to pull that message out and examine it's unread status yet.

