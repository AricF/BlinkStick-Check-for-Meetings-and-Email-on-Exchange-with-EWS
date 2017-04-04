using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Configuration;
using BlinkStickDotNet;

namespace CheckForMeetingReminders
{

    class Program
    {

        //I know, look at all these globals.
        //I started out not realizing how large this was going to get.  I should've broken this up into
        //seperate classes and passed parameters but it started as a two method console app with no real upfront design...

        //logic check...
        //we care about
        // 1. being 15 minutes away from a meeting (blink medium priority)
        // 2. being 5 minutes away from a meeting (blink highest priority
        // 3. being in a meeting  (blink low priority)
        // 4. could've gotten a new email (blink lowest priority)

        static bool meetingLessThan5MinutesAway = false;   //highest priority
        static bool meetingLessThan15MinutesAway = false;  //medium priority
        static bool inAMeeting = false;                    //low priority
        static int newUnreadEmailSinceStartingTheApp = 0; //lowest priority


        //the first time we run we record the number of unread emails.
        //we continue to use that value to compare future unread emails which to us indicates that there is new mail.
        //this logic can start to fall apart as the user reads/deletes previously unread emails.
        static bool firstTimeCheck = true;

		//setup threads used to animate the Blinkstick
		static System.Threading.Thread threadForBlinking = new System.Threading.Thread(new System.Threading.ThreadStart(AnimationForAll));




        //timers, like fezzes, are cool
        public static System.Timers.Timer timerForCheckingExchangeForEmailsAndReminders;
        public static System.Timers.Timer timerForUpdatingTheDisplay;

        public static bool tracingToScreen = false;
        public static bool tracingToLogFile = false;
        

        static int NumberOfUnreadItemsSinceProgramStarted = 0;
        static int NumberOfTotalEmailsSinceProgramStarted = 0;


        static string Office365WebServicesURL = "";
        static int frequencyOfCheckingAppointmentsInMinutes = 5;
        static int frequencyOfUpdatingDisplayInSeconds = 3;

        static NetworkCredential userCredentials;

        static DateTime lastTimeChecked;
        static DateTime nextTimeToCheck;
        static TimeSpan minutesLeftBeforeNextCheck;

        static BlinkStick device;



        static void Main(string[] args)
        {

            bool shouldExitApp = false;

            tracingToLogFile = Convert.ToBoolean(ConfigurationManager.AppSettings["tracingToLogFile"]);
            tracingToScreen = Convert.ToBoolean(ConfigurationManager.AppSettings["tracingToScreen"]);
            
            string logFileName = ConfigurationManager.AppSettings["logFileName"];
            if (string.IsNullOrEmpty(logFileName))
            {
                logFileName = "CheckForMeetingReminders.log";
            }

            Office365WebServicesURL = ConfigurationManager.AppSettings["Office365WebServicesURL"];

            if (string.IsNullOrEmpty(Office365WebServicesURL))
            {
                Console.WriteLine("Exchange URL not setup in app.config.  It should look like https://<servername>/ews/Exchange.asmx");
            }

            frequencyOfCheckingAppointmentsInMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["frequencyOfCheckingAppointmentsInMinutes"]);

            if (frequencyOfCheckingAppointmentsInMinutes == 0)
            {
                frequencyOfCheckingAppointmentsInMinutes = 5;
                Console.WriteLine("frequencyOfCheckingAppointments not setup in app.config.  Defaulting to 5 minutes");
            }

            //Trust all certificates - sledgehammer approach, you should setup a more intelligent callback.
            //System.Net.ServicePointManager.ServerCertificateValidationCallback =
            //    ((sender, certificate, chain, sslPolicyErrors) => true);

            //this is a little safer than the above line and allows for self signed certs and verifies the
            //cert is on the correct machine by machine name
            ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallBack;

            // Start tracing to console and a log file.
            Tracing.OpenLog("./" + logFileName);
            Tracing.WriteLine(false, "Blink on Appointments and new Mail started:" + DateTime.Now);
            Console.WriteLine("Blink on Appointments and new Mail started:" + DateTime.Now);

            //Like the name of this variable? Several posts and some sample code for EWS said this should be the email address.
            //That's simply not true.  It should be the login id (which could possibly be the the same as the email address
            //if your admin chose to set things up that way but I've never seen that situation occur.)
            var loginIDNotEmailAddress = ConfigurationManager.AppSettings["loginID"];

            if (string.IsNullOrEmpty(loginIDNotEmailAddress))
            {

                Console.Write("Enter your login ID (you can also set this in the app.config file to avoid having to type it in): ");
                loginIDNotEmailAddress = Console.ReadLine();
            }


            SecureString password = GetPasswordFromConsoleOrConfig();
            if (password.Length == 0)
            {

                Tracing.WriteLine(false, "Password empty, closing program.");
                Console.WriteLine("Password empty, closing program.");
                goto Finish;  //a goto!  The world is going to end!!!
            }



            var domain = ConfigurationManager.AppSettings["domain"];

            if (string.IsNullOrEmpty(domain))
            {

                Console.Write("Enter your AD Domain.  Just press return to leave it blank if you don't have one. (you can also set this in the app.config file to avoid having to type it in): ");
                loginIDNotEmailAddress = Console.ReadLine();
            }

            if (string.IsNullOrEmpty(domain))
            {

                userCredentials = new NetworkCredential(loginIDNotEmailAddress, password);
            }
            else
            {
                userCredentials = new NetworkCredential(loginIDNotEmailAddress, password, domain);
            }

            //setup hardware
            device = BlinkStick.FindFirst();
            if (device != null && device.OpenDevice())
            {
                //turn off lights at start
                device.TurnOff();

            }
            else
            {
                Tracing.WriteLine(true, "Could not find a Blinkstick device.");
                Console.WriteLine("\nCould not find a Blinkstick device.\n");
                goto Finish;  //a goto!  The world is going to end!!!
            }



            timerForCheckingExchangeForEmailsAndReminders = new System.Timers.Timer();
            timerForCheckingExchangeForEmailsAndReminders.Interval = TimeSpan.FromMinutes(frequencyOfCheckingAppointmentsInMinutes).TotalMilliseconds;
            timerForCheckingExchangeForEmailsAndReminders.Elapsed += new System.Timers.ElapsedEventHandler(CheckFrequencyTimerElapsed);
            timerForCheckingExchangeForEmailsAndReminders.Enabled = true;

            timerForUpdatingTheDisplay = new System.Timers.Timer();
            timerForUpdatingTheDisplay.Interval = TimeSpan.FromSeconds(frequencyOfUpdatingDisplayInSeconds).TotalMilliseconds;
            timerForUpdatingTheDisplay.Elapsed += new System.Timers.ElapsedEventHandler(DisplayTimerElapsed);
            timerForUpdatingTheDisplay.Enabled = true;

            

            threadForBlinking.IsBackground = true;

            threadForBlinking.Start();

            DrawMenu();

            //do one initial check, the others are triggered by the timerForCheckingExchangeForEmailsAndReminders timer.
            CheckForAppointmentsAndNewEmails();


            while (!shouldExitApp)
            {
                ConsoleKeyInfo info = Console.ReadKey();
                if (info.Key == ConsoleKey.Escape)
                {
                    shouldExitApp = true;
                }
                if (info.Key == ConsoleKey.R)
                {
                    firstTimeCheck = true;
                    CheckForNewUnreadMail(userCredentials);
                    AlterBlinkStatusAsAppropriate(newUnreadEmailSinceStartingTheApp, meetingLessThan5MinutesAway, inAMeeting, meetingLessThan15MinutesAway);
                }
                DrawMenu();
            }

            Tracing.WriteLine(true, "Exiting...");
            Console.WriteLine("\nExiting...");

        Finish:

            Tracing.CloseLog();
            if (!shouldExitApp)  //if this isn't true we got here due to an error condition to halt the exit...
            {
                Console.WriteLine("Press any key to exit: ");
                Console.ReadKey();
            }

			//Doesn't exit properly in Mono.  The HidSharp thread keeps running
			//I've tried the lines below but they don't work.
			//threadForBlinking.Abort();
			//System.Environment.Exit(0);

        }
        

        static void AnimationForAll()
        {

            var color = new RgbColor();
            string[] colorString;

            int repeatAmount = 0;
            int blinkTime = 500; //500 = 1/2 second
            while (true)
            {
                if (meetingLessThan5MinutesAway)
                {
                    //we dynamically load the color settings in case the user changed the app.config file what the app was running
                    //and didn't want to restart the app for the changes to take effect.
                    try
                    {
                        colorString = (ConfigurationManager.AppSettings["meetingLessThan5MinutesAwayRGB"]).Split(',');

                        color.R = Convert.ToByte(colorString[0]);
                        color.G = Convert.ToByte(colorString[1]);
                        color.B = Convert.ToByte(colorString[2]);

                        blinkTime = Convert.ToInt32(ConfigurationManager.AppSettings["meetingLessThan5MinutesAwayBlinkRateInMilliseconds"]);
                        repeatAmount = Convert.ToInt32(ConfigurationManager.AppSettings["meetingLessThan5MinutesAwayBlinkRepeatAmount"]);
                    }
                    catch (Exception ex)
                    {
                        //why?  if the user misconfigured the app.config file
                        Tracing.WriteLine(true, "exception trying to set LED colors. Check the app.config file for errors or missing parameters.");
                        Tracing.WriteLine(false, ex.ToString());
                        color.R = 255;
                        color.G = 0;
                        color.B = 0;
                        repeatAmount = 5;
                        blinkTime = 100;
                    }
                    device.Blink(0, 0, color, repeatAmount, blinkTime);
                    System.Threading.Thread.Sleep(20); //it's a good idea to give the Blinkstick time to breathe between calls
                }
                else if (meetingLessThan15MinutesAway)
                {
                    try
                    {
                        colorString = (ConfigurationManager.AppSettings["meetingLessThan15MinutesAwayRGB"]).Split(',');

                        color.R = Convert.ToByte(colorString[0]);
                        color.G = Convert.ToByte(colorString[1]);
                        color.B = Convert.ToByte(colorString[2]);

                        blinkTime = Convert.ToInt32(ConfigurationManager.AppSettings["meetingLessThan15MinutesAwayBlinkRateInMilliseconds"]);
                        repeatAmount = Convert.ToInt32(ConfigurationManager.AppSettings["meetingLessThan15MinutesAwayBlinkRepeatAmount"]);
                    }
                    catch (Exception ex)
                    {
                        //why?  if the user misconfigured the app.config file
                        Tracing.WriteLine(true, "exception trying to set LED colors. Check the app.config file for errors or missing parameters.");
                        Tracing.WriteLine(false, ex.ToString());
                        color.R = 255;
                        color.G = 255;
                        color.B = 0;
                        repeatAmount = 5;
                        blinkTime = 300;
                    }
                    device.Blink(0, 0, color, repeatAmount, blinkTime);
                    System.Threading.Thread.Sleep(20); //it's a good idea to give the Blinkstick time to breathe between calls

                }
                else if (inAMeeting)
                {
                    try
                    {
                        colorString = (ConfigurationManager.AppSettings["inAMeetingRGB"]).Split(',');

                        color.R = Convert.ToByte(colorString[0]);
                        color.G = Convert.ToByte(colorString[1]);
                        color.B = Convert.ToByte(colorString[2]);

                        blinkTime = Convert.ToInt32(ConfigurationManager.AppSettings["inAMeetingBlinkRateInMilliseconds"]);
                        repeatAmount = Convert.ToInt32(ConfigurationManager.AppSettings["inAMeetingBlinkRepeatAmount"]);
                    }
                    catch (Exception ex)
                    {
                        //why?  if the user misconfigured the app.config file
                        Tracing.WriteLine(true, "exception trying to set LED colors. Check the app.config file for errors or missing parameters.");
                        Tracing.WriteLine(false, ex.ToString());
                        color.R = 0;
                        color.G = 0;
                        color.B = 255;
                        repeatAmount = 5;
                        blinkTime = 500;
                    }
                    device.Blink(0, 0, color, repeatAmount, blinkTime);


                    System.Threading.Thread.Sleep(20); //it's a good idea to give the Blinkstick time to breathe between calls

                }

                if (newUnreadEmailSinceStartingTheApp > 0)
                {
                    try
                    {
                        colorString = (ConfigurationManager.AppSettings["newUnreadEmailSinceStartingTheAppRGB"]).Split(',');

                        color.R = Convert.ToByte(colorString[0]);
                        color.G = Convert.ToByte(colorString[1]);
                        color.B = Convert.ToByte(colorString[2]);

                        blinkTime = Convert.ToInt32(ConfigurationManager.AppSettings["newUnreadEmailSinceStartingTheAppBlinkRateInMilliseconds"]);
                        repeatAmount = Convert.ToInt32(ConfigurationManager.AppSettings["newUnreadEmailSinceStartingTheAppBlinkRepeatAmount"]);
                    }
                    catch (Exception ex)
                    {
                        //why?  if the user misconfigured the app.config file
                        Tracing.WriteLine(true, "exception trying to set LED colors. Check the app.config file for errors or missing parameters.");
                        Tracing.WriteLine(false, ex.ToString());
                        color.R = 0;
                        color.G = 255;
                        color.B = 0;
                        repeatAmount = 1;
                        blinkTime = 500;
                    }


                    device.Pulse(0, 0, color, repeatAmount, blinkTime, 50);
                    System.Threading.Thread.Sleep(20); //it's a good idea to give the Blinkstick time to breathe between calls

                }


                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));

            }

        }

        static void ResetTimerForDisplay()
        {
            lastTimeChecked = DateTime.Now;
            nextTimeToCheck = lastTimeChecked.AddMinutes(frequencyOfCheckingAppointmentsInMinutes);
        }

        static void CheckFrequencyTimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            DrawMenu();
            CheckForAppointmentsAndNewEmails();
        }

        static void CheckForAppointmentsAndNewEmails()
        {
            ResetTimerForDisplay();

            CheckForNewUnreadMail(userCredentials);
            CheckForUpcomingAppointments(userCredentials);

            AlterBlinkStatusAsAppropriate(newUnreadEmailSinceStartingTheApp, meetingLessThan5MinutesAway, inAMeeting, meetingLessThan15MinutesAway);


        }

        static void AlterBlinkStatusAsAppropriate(int howManyNewEmails, bool meetingLessThan5MinutesAway, bool inAMeeting, bool meetingLessThan15MinutesAway)
        {
            // meetingLessThan5MinutesAway          //highest priority
            // meetingLessThan15MinutesAway         //medium priority
            // inAMeeting                           //low priority
            // newUnreadEmailSinceStartingTheApp    //lowest priority


            //then determine if meetingLessThan5MinutesAway, inAMeeting, or meetingLessThan15MinutesAway are already true

            //Then prioritize like this:
            // 1. being 15 minutes away from a meeting (blink medium priority)
            // 2. being 5 minutes away from a meeting (blink highest priority
            // 3. being in a meeting  (blink low priority)

            //if we are 3 and not 1 or 2, then low priority
            //if we are 2 and not 1 then medium priority
            //if we are 1 high 

            if (newUnreadEmailSinceStartingTheApp > 0)
            {
                //just light up
                Console.WriteLine("\nthere is new unread mail\n");
            }

            if (meetingLessThan5MinutesAway)
            {
                //blink like the dickens!
                Console.WriteLine("\nmeeting is less than 5 minutes away\n");
            }
            else if (meetingLessThan15MinutesAway)
            {
                //blink a lot
                Console.WriteLine("\nmeeting is less than 15 minutes away\n");
            }
            else if (inAMeeting)
            {
                //blink some
                Console.WriteLine("\nyou are in a meeting\n");
            }
            else if (newUnreadEmailSinceStartingTheApp <= 0)  //note that we may want to find a way to indicate this alongside meeting notifications as they aren't exclusive
            {
                //turn off the light
                Console.WriteLine("\nnothing to see here... move along\n");
            }

        }

        static void DisplayTimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            minutesLeftBeforeNextCheck = nextTimeToCheck.Subtract(DateTime.Now);
            Console.CursorLeft = 0;
            string humanReadableTimespan = string.Format("{0:hh\\:mm\\:ss}", minutesLeftBeforeNextCheck);
            Console.Write(humanReadableTimespan + " until next check");
        }

        private static SecureString GetPasswordFromConsoleOrConfig()
        {
            SecureString password = new SecureString();
            bool readingPassword = true;

            var insecurePassword = ConfigurationManager.AppSettings["password"];



            if (!string.IsNullOrEmpty(insecurePassword))
            {
                foreach (char c in insecurePassword)
                    password.AppendChar(c);

                password.MakeReadOnly();
                Console.WriteLine("Password loaded from app.config. Warning: for security it is safer to manually enter the password by removing from app.config.");
            }
            else
            {

                Console.Write("Enter password (you can add password to the app.config although that's not very secure): ");

                while (readingPassword)
                {
                    ConsoleKeyInfo userInput = Console.ReadKey(true);

                    switch (userInput.Key)
                    {
                        case (ConsoleKey.Enter):
                            readingPassword = false;
                            break;
                        case (ConsoleKey.Escape):
                            password.Clear();
                            readingPassword = false;
                            break;
                        case (ConsoleKey.Backspace):
                            if (password.Length > 0)
                            {
                                password.RemoveAt(password.Length - 1);
                                Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                                Console.Write(" ");
                                Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            }
                            break;
                        default:
                            if (userInput.KeyChar != 0)
                            {
                                password.AppendChar(userInput.KeyChar);
                                Console.Write("*");
                            }
                            break;
                    }
                }
                password.MakeReadOnly();
                Console.WriteLine();
            }


            return password;
        }

        private static void CheckForNewUnreadMail(NetworkCredential userCredentials)
        {
            /// This is the XML request that is sent to the Exchange server.
            var getFolderSOAPRequest =
                  "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
                  "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n" +
                  "   xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n" +
                  "<soap:Header>\n" +
                  "    <t:RequestServerVersion Version=\"Exchange2007_SP1\" />\n" +
                  "  </soap:Header>\n" +
                  "  <soap:Body>\n" +
                  "    <GetFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"\n" +
                  "               xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n" +
                  "      <FolderShape>\n" +
                  "        <t:BaseShape>Default</t:BaseShape>\n" +
                  "      </FolderShape>\n" +
                  "      <FolderIds>\n" +
                  "        <t:DistinguishedFolderId Id=\"inbox\"/>\n" +
                  "      </FolderIds>\n" +
                  "    </GetFolder>\n" +
                  "  </soap:Body>\n" +
                  "</soap:Envelope>\n";

            // Write the get folder operation request to the console and log file.
            Tracing.WriteLine(true, "Get folder operation request:");
            Tracing.WriteLine(true, getFolderSOAPRequest);
            //  string Office365WebServicesURL = "https://dv_casarray.apsc.com/ews/Exchange.asmx";
            var getWebRequest = WebRequest.CreateHttp(Office365WebServicesURL);
            getWebRequest.AllowAutoRedirect = false;
            getWebRequest.Credentials = userCredentials;
            getWebRequest.Method = "POST";
            getWebRequest.ContentType = "text/xml";

            try
            {

                var requestWriter = new StreamWriter(getWebRequest.GetRequestStream());
                requestWriter.Write(getFolderSOAPRequest);
                requestWriter.Close();


                var getFolderResponse = (HttpWebResponse)(getWebRequest.GetResponse());
                if (getFolderResponse.StatusCode == HttpStatusCode.OK)
                {
                    var responseStream = getFolderResponse.GetResponseStream();
                    XElement responseEnvelope = XElement.Load(responseStream);
                    if (responseEnvelope != null)
                    {
                        // Write the response to the console and log file.
                        Tracing.WriteLine(true, "Response:");
                        StringBuilder stringBuilder = new StringBuilder();
                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.Indent = true;
                        XmlWriter writer = XmlWriter.Create(stringBuilder, settings);
                        responseEnvelope.Save(writer);
                        writer.Close();
                        Tracing.WriteLine(true, stringBuilder.ToString());

                        // Check the response for error codes. If there is an error, throw an application exception.
                        IEnumerable<XElement> errorCodes = from errorCode in responseEnvelope.Descendants
                                                           ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                                                           select errorCode;
                        foreach (var errorCode in errorCodes)
                        {
                            if (errorCode.Value != "NoError")
                            {
                                switch (errorCode.Parent.Name.LocalName.ToString())
                                {
                                    case "Response":
                                        string responseError = "Response-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(responseError);

                                    case "UserResponse":
                                        string userError = "User-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(userError);
                                }
                            }
                        }

                        // Process the response.
                        IEnumerable<XElement> folders = from folderElement in
                                                            responseEnvelope.Descendants
                                                            ("{http://schemas.microsoft.com/exchange/services/2006/messages}Folders")
                                                        select folderElement;

                        foreach (var folder in folders)
                        {
                            Tracing.Write(true, "Folder name:     ");
                            var folderName = from folderElement in
                                                 folder.Descendants
                                                 ("{http://schemas.microsoft.com/exchange/services/2006/types}DisplayName")
                                             select folderElement.Value;
                            Tracing.WriteLine(true, folderName.ElementAt(0));

                            Tracing.Write(true, "Total messages:  ");
                            var totalCount = from folderElement in
                                                 folder.Descendants
                                                   ("{http://schemas.microsoft.com/exchange/services/2006/types}TotalCount")
                                             select folderElement.Value;
                            Tracing.WriteLine(true, totalCount.ElementAt(0));

                            Tracing.Write(true, "Unread messages: ");
                            var unreadCount = from folderElement in
                                                  folder.Descendants
                                                    ("{http://schemas.microsoft.com/exchange/services/2006/types}UnreadCount")
                                              select folderElement.Value;
                            Tracing.WriteLine(true, unreadCount.ElementAt(0));

                            var totalEmails = Convert.ToInt32(totalCount.First());
                            var currentUnreadEmails = Convert.ToInt32(unreadCount.First());
                            if (firstTimeCheck)
                            {
                                NumberOfUnreadItemsSinceProgramStarted = currentUnreadEmails;
                                NumberOfTotalEmailsSinceProgramStarted = totalEmails;
                                newUnreadEmailSinceStartingTheApp = 0;
                                firstTimeCheck = false;
                            }
                            else
                            {

                                newUnreadEmailSinceStartingTheApp = currentUnreadEmails - NumberOfUnreadItemsSinceProgramStarted;

                                //this case only happens if the user read or deleted some unread emails that existed 
                                //prior to starting this app but we may as well account for that.
                                if (newUnreadEmailSinceStartingTheApp < 0)
                                {
                                    NumberOfUnreadItemsSinceProgramStarted = currentUnreadEmails;
                                    newUnreadEmailSinceStartingTheApp = 0;
                                }
                                //we could probably make this smarter by adding some logic for NumberOfTotalEmailsSinceProgramStarted
                                //to deal with the user having deleted a combination of already read, and previously unread emails.
                            }
                        }
                    }//if not null
                }
                else //http response not ok
                {
                    //should probably throw an exception here...
                    Tracing.WriteLine(false, "HTTP response was not ok: " + getFolderResponse.StatusCode.ToString());
                    Console.WriteLine("HTTP response was not ok: " + getFolderResponse.StatusCode.ToString());
                }

            }
            catch (WebException ex)
            {
                Tracing.WriteLine(false, "Caught Web exception:");
                Tracing.WriteLine(false, ex.Message);

                Console.WriteLine("Caught Web exception:");
                Console.WriteLine(ex.Message);
            }
            catch (ApplicationException ex)
            {
                Tracing.WriteLine(false, "Caught application exception:");
                Tracing.WriteLine(false, ex.Message);

                Console.WriteLine("Caught application exception:");
                Console.WriteLine(ex.Message);
            }

        }

        private static void CheckForUpcomingAppointments(NetworkCredential userCredentials)
        {

            DateTime startOfPeriod = DateTime.Now;



            DateTime endOfPeriod = startOfPeriod.AddMinutes(15);

            Tracing.WriteLine(true, "startOfPeriod: " + String.Format("{0:s}", startOfPeriod));
            Tracing.WriteLine(true, "endOfPeriod:   " + String.Format("{0:s}", endOfPeriod));


            string ourRequest =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
                "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"\n" +
                "               xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\"\n" +
                "               xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n" +
                "               xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n" +
                "  <soap:Body>\n" +
                "    <GetUserAvailabilityRequest xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"\n" +
                "                xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n" +
                "     <t:TimeZone xmlns=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n" +
                "        <Bias>" + 60 * 7 + "</Bias>\n" +  //minutes from UTC, so 480 = 8 hours so that's PST, 472 = MST 
                "        <StandardTime>\n" +
                "          <Bias>0</Bias>\n" +
                "          <Time>02:00:00</Time>\n" +
                "          <DayOrder>5</DayOrder>\n" +
                "          <Month>10</Month>\n" +
                "          <DayOfWeek>Sunday</DayOfWeek>\n" +
                "        </StandardTime>\n" +
                "        <DaylightTime>\n" +
                "          <Bias>0</Bias>\n" +  //normally -60 but in Arizona 0 because we don't do Daylight Savings Time
                "          <Time>02:00:00</Time>\n" +
                "          <DayOrder>1</DayOrder>\n" +
                "          <Month>4</Month>\n" +
                "          <DayOfWeek>Sunday</DayOfWeek>\n" +
                "        </DaylightTime>\n" +
                "      </t:TimeZone>\n" +
                "      <MailboxDataArray>\n" +
                "        <t:MailboxData>\n" +
                "          <t:Email>\n" +
                "            <t:Address>aric.friesen@aps.com</t:Address>\n" +
                "          </t:Email>\n" +
                "          <t:AttendeeType>Required</t:AttendeeType>\n" +
                "          <t:ExcludeConflicts>false</t:ExcludeConflicts>\n" +
                "        </t:MailboxData>\n" +
                "      </MailboxDataArray>\n" +
                "      <t:FreeBusyViewOptions>\n" +
                "        <t:TimeWindow>\n" +
                "          <t:StartTime>" + String.Format("{0:s}", startOfPeriod) + "</t:StartTime>\n" + //example: 2017-03-02T00:00:00
                "          <t:EndTime>" + String.Format("{0:s}", endOfPeriod) + "</t:EndTime>\n" +  //example: 2017-03-02T23:59:59
                "        </t:TimeWindow>\n" +
                "        <t:MergedFreeBusyIntervalInMinutes>5</t:MergedFreeBusyIntervalInMinutes>\n" + //example value = 60
                "        <t:RequestedView>DetailedMerged</t:RequestedView>\n" +
                "      </t:FreeBusyViewOptions>\n" +
                "    </GetUserAvailabilityRequest>\n" +
                "  </soap:Body>\n" +
                "</soap:Envelope> \n";





            // Write the get folder operation request to the console and log file.
            Tracing.WriteLine(true, "Get operation request:");
            Tracing.WriteLine(true, ourRequest);

            var getWebRequest = WebRequest.CreateHttp(Office365WebServicesURL);
            getWebRequest.AllowAutoRedirect = false;
            getWebRequest.Credentials = userCredentials;
            getWebRequest.Method = "POST";
            getWebRequest.ContentType = "text/xml";

            try
            {

                var requestWriter = new StreamWriter(getWebRequest.GetRequestStream());
                requestWriter.Write(ourRequest);
                requestWriter.Close();


                var getFolderResponse = (HttpWebResponse)(getWebRequest.GetResponse());
                if (getFolderResponse.StatusCode == HttpStatusCode.OK)
                {
                    var responseStream = getFolderResponse.GetResponseStream();
                    XElement responseEnvelope = XElement.Load(responseStream);
                    if (responseEnvelope != null)
                    {
                        // Write the response to the console and log file.
                        Tracing.WriteLine(true, "Response:");
                        StringBuilder stringBuilder = new StringBuilder();
                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.Indent = true;
                        XmlWriter writer = XmlWriter.Create(stringBuilder, settings);
                        responseEnvelope.Save(writer);
                        writer.Close();
                        Tracing.WriteLine(true, stringBuilder.ToString());

                        // Check the response for error codes. If there is an error, throw an application exception.
                        IEnumerable<XElement> errorCodes = from errorCode in responseEnvelope.Descendants
                                                           ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                                                           select errorCode;
                        foreach (var errorCode in errorCodes)
                        {
                            if (errorCode.Value != "NoError")
                            {
                                switch (errorCode.Parent.Name.LocalName.ToString())
                                {
                                    case "Response":
                                        string responseError = "Response-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(responseError);

                                    case "UserResponse":
                                        string userError = "User-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(userError);
                                }
                            }
                        }

                        // Process the response.

                        IEnumerable<XElement> folders = from folderElement in

                                                            responseEnvelope.Descendants("{http://schemas.microsoft.com/exchange/services/2006/types}CalendarEvent")
                                                        select folderElement;


                        //logic check...
                        //we care about
                        // 1. being 15 minutes away from a meeting (blink medium priority)
                        // 2. being 5 minutes away from a meeting (blink highest priority
                        // 3. being in a meeting  (blink low priority)
                        // 4. could've gotten a new email (blink lowest priority)
                        // - and note, we may not have any items in the following enumeration in which case we have no meetings pending or happening now.

                        DateTime startTime;
                        DateTime endTime;

                        //reset status alternately could check for folders containing 0 calendarEvents and set 
                        meetingLessThan5MinutesAway = false;
                        meetingLessThan15MinutesAway = false;
                        inAMeeting = false;


                        foreach (var calendarEvent in folders)
                        {
                            Tracing.Write(true, "Start Time:     ");
                            var folderName = from folderElement in
                                                 calendarEvent.Descendants("{http://schemas.microsoft.com/exchange/services/2006/types}StartTime")
                                             select folderElement.Value;
                            Tracing.WriteLine(true, folderName.ElementAt(0));

                            //example format: Start Time:     2017-03-02T09:30:00
                            startTime = Convert.ToDateTime(folderName.ElementAt(0));

                            Tracing.Write(true, "End Time:  ");
                            var totalCount = from folderElement in
                                                 calendarEvent.Descendants("{http://schemas.microsoft.com/exchange/services/2006/types}EndTime")
                                             select folderElement.Value;
                            Tracing.WriteLine(true, totalCount.ElementAt(0));

                            //example format: End Time:  2017-03-02T10:00:00
                            endTime = Convert.ToDateTime(totalCount.ElementAt(0));

                            DoDateLogicForAppointments(startTime, endTime);
                        }



                    }// if responseEnvolope not null
                }
                else //http response not ok
                {
                    //should probably throw an exception here...
                    Tracing.WriteLine(false, "HTTP response was not ok: " + getFolderResponse.StatusCode.ToString());
                    Console.WriteLine("HTTP response was not ok: " + getFolderResponse.StatusCode.ToString());
                }

            }
            catch (WebException ex)
            {
                Tracing.WriteLine(false, "Caught Web exception:");
                Tracing.WriteLine(false, ex.Message);

                Console.WriteLine("Caught Web exception:");
                Console.WriteLine(ex.Message);
            }
            catch (ApplicationException ex)
            {
                Tracing.WriteLine(false, "Caught application exception:");
                Tracing.WriteLine(false, ex.Message);

                Console.WriteLine("Caught application exception:");
                Console.WriteLine(ex.Message);
            }

        }

        static void DoDateLogicForAppointments(DateTime startTime, DateTime endTime)
        {
            /*Our global bools:
                newUnreadEmailSinceStartingTheApp
                meetingLessThan5MinutesAway
                inAMeeting
                meetingLessThan15MinutesAway */

            //first determine if startTime is within 5 minutes of now
            //second determine if startTime is within 15 minutes of now
            //third determine if startTime is past now and if endtime is after now (then we are in a meeting)

            var rightNow = DateTime.Now;


            //we'll include the frequencyOfCheckingAppointmentsInMinutes in addition because in theory if your frequency was 5 minutes.
            //and we didn't subtract it here you could miss the 5 minute warning entirely if it happend to be checking on the 5 minute mark
            //of each cycle.
            if (startTime.AddMinutes(-5 - frequencyOfCheckingAppointmentsInMinutes) <= rightNow && !(startTime < rightNow))
            {
                meetingLessThan5MinutesAway = true;
                meetingLessThan15MinutesAway = false;
                inAMeeting = false;
            }
            else if (startTime.AddMinutes(-15 - frequencyOfCheckingAppointmentsInMinutes) <= rightNow && !(startTime < rightNow))
            {
                meetingLessThan5MinutesAway = false;
                meetingLessThan15MinutesAway = true;
                inAMeeting = false;
            }
            else if (startTime < rightNow && endTime >= rightNow)
            {
                meetingLessThan5MinutesAway = false;
                meetingLessThan15MinutesAway = false;
                inAMeeting = true;
            }


        }

        // callback used to validate the certificate in an SSL conversation
        private static bool CertificateValidationCallBack(object sender, System.Security.Cryptography.X509Certificates.X509Certificate cert, System.Security.Cryptography.X509Certificates.X509Chain chain, System.Net.Security.SslPolicyErrors policyErrors)
        {
            bool result = false;
            string serverName = "owa.apsc.com".ToUpper();
            if (cert.Subject.ToUpper().Contains(serverName))
            {
                result = true;
            }

            return result;
        }

        private static void DrawMenu()
        {
            Console.Clear();
            Console.WriteLine("Menu:");
            Console.WriteLine("Esc = Exit");
            Console.WriteLine("R = Reset email count to 0 new emails");

            //haven't written the logic for the next line yet...
            //Console.WriteLine("M = current meeting ended early, stop indicating I need to be in a meeting.");
        }



    }
}
