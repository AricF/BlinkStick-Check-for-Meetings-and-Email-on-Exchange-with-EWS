using System;
using System.IO;

namespace CheckForMeetingReminders
{
    /// <summary>
    /// Simple logging and screen output class.
    /// </summary>
    class Tracing
    {
        private static TextWriter logFileWriter = null;
        

       

        public static void OpenLog(string fileName)
        {
            if (Program.tracingToLogFile)
            {
                logFileWriter = new StreamWriter(fileName);
            }
        }

        /// <summary>
        /// Writes to log and screen
        /// </summary>
        /// <param name="displayToScreen">If the message is more for debugging and doesn't need to be shown to the user this should be set to true.  If it's set to false it's understood that the message is important that there is another Console.Write call in the app code to make sure the message is displayed to the user even when Tracing output to the screen is disabled in the config.</param>
        /// <param name="format"></param>
        /// <param name="args"></param>
        public static void Write(bool displayToScreen,string format, params object[] args)
        {
            if (Program.tracingToScreen)
            {
                Console.Write(format, args);
            }
            if (logFileWriter != null && Program.tracingToLogFile)
            {
                logFileWriter.Write(format, args);
            }
        }

        /// <summary>
        /// Writes to log and screen
        /// </summary>

        /// <param name="displayToScreen">If the message is more for debugging and doesn't need to be shown to the user this should be set to true.  If it's set to false it's understood that the message is important that there is another Console.Write call in the app code to make sure the message is displayed to the user even when Tracing output to the screen is disabled in the config.</param>
        /// <param name="format"></param>
        /// <param name="args"></param>
        public static void WriteLine(bool displayToScreen, string format, params object[] args)
        {
            if (Program.tracingToScreen)
            {
                Console.WriteLine(format, args);
            }
            if (logFileWriter != null && Program.tracingToLogFile)
            {
                logFileWriter.WriteLine(format, args);
            }
        }

        public static void CloseLog()
        {
            if (logFileWriter != null && Program.tracingToLogFile)
            {
                logFileWriter.Flush();
                logFileWriter.Close();
            }
        }
    }
}