using System;
using System.IO;

namespace APE.Communication
{
    /// <summary>
    /// Contains methods to help in debugging APE.Communication
    /// </summary>
    static class DebugLogging
    {
        private static object theLock = new object();

        /// <summary>
        /// Used to help debug issues in APE.Communication
        /// All lines of this method should be commented out in normal operation as it has a large impact on performance
        /// </summary>
        /// <param name="Line"></param>
        internal static void WriteLog(string Line)
        {
            /*
            lock (theLock)    // Needed as WriteLog is used in callbacks which are async
            {
                TextWriter log = File.AppendText(Environment.GetEnvironmentVariable("TEMP") + @"\APE_Debug.log");
                log.WriteLine(DateTime.Now.ToString() + "\t" + Line);
                log.Close();
            }
            */
        }
    }
}
