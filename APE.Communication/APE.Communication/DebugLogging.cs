//
//Copyright 2016 David Beales
//
//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.
//
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
