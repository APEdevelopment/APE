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
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.IO.Compression;
using NM = APE.Native.NativeMethods;

namespace APE.Capture
{
    /// <summary>
    /// Provides methods to capture debug string output
    /// </summary>
    public sealed class DebugString
    {
        /// <summary>
        /// Private constructor so no one can create a instance of this static class
        /// </summary>
        private DebugString()
        {
        }

        /// <summary>
        /// String to store the file name
        /// </summary>
        private static string m_Filename;

        /// <summary>
        /// Process ID to filter the output for
        /// </summary>
        private static int m_PID;

        /// <summary>
        /// TextWriter to handle output to the file
        /// </summary>
        private static TextWriter m_tw;

        /// <summary>
        /// Event handle for slot 'DBWIN_BUFFER_READY'
        /// </summary>
        private static IntPtr m_AckEvent = IntPtr.Zero;

        /// <summary>
        /// Event handle for slot 'DBWIN_DATA_READY'
        /// </summary>
        private static IntPtr m_ReadyEvent = IntPtr.Zero;

        /// <summary>
        /// Handle for our shared file
        /// </summary>
        private static IntPtr m_SharedFile = IntPtr.Zero;

        /// <summary>
        /// Handle for our shared memory
        /// </summary>
        private static IntPtr m_SharedMem = IntPtr.Zero;

        /// <summary>
        /// Our capturing thread
        /// </summary>
        private static Thread m_Capturer = null;

        /// <summary>
        /// Our synchronization root
        /// </summary>
        private static object m_SyncRoot = new object();

        /// <summary>
        /// Mutex for singleton check
        /// </summary>
        private static Mutex m_Mutex = null;

        /// <summary>
        /// Starts the debug string capture
        /// <param name="Filename">The file of the zipfile to capture the debug string output in</param>
        /// </summary>
        public static void StartCapture(string Filename)
        {
            StartCapture(Filename, null);
        }

        /// <summary>
        /// Starts the debug string capture
        /// <param name="Filename">The file of the zipfile to capture the debug string output in</param>
        /// <param name="Process">The process to filter the debug string output for</param>
        /// </summary>
        public static void StartCapture(string Filename, Process process)
        {
            m_Filename = Filename;

            if (process == null)
            {
                m_PID = 0;
            }
            else
            {
                m_PID = process.Id;
            }

            lock (m_SyncRoot)
            {
                if (m_Capturer != null)
                {
                    throw new ApplicationException("DebugString capture is already started");
                }

                // Check for multiple instances	
                bool createdNew = false;
                m_Mutex = new Mutex(false, typeof(DebugString).Namespace, out createdNew);
                if (!createdNew)
                {
                    throw new ApplicationException("There is already an instance running");
                }

                //NM.SECURITY_ATTRIBUTES sa = new NM.SECURITY_ATTRIBUTES();
                //sa.lpSecurityDescriptor = IntPtr.Zero;
                IntPtr sa = IntPtr.Zero;

                // Create the event for slot 'DBWIN_BUFFER_READY'
                m_AckEvent = NM.CreateEvent(sa, false, false, "DBWIN_BUFFER_READY");
                if (m_AckEvent == IntPtr.Zero)
                {
                    throw CreateApplicationException("Failed to create event 'DBWIN_BUFFER_READY'");
                }

                // Create the event for slot 'DBWIN_DATA_READY'
                m_ReadyEvent = NM.CreateEvent(sa, false, false, "DBWIN_DATA_READY");
                if (m_ReadyEvent == IntPtr.Zero)
                {
                    throw CreateApplicationException("Failed to create event 'DBWIN_DATA_READY'");
                }

                // Get a handle to the readable shared memory at slot 'DBWIN_BUFFER'.
                m_SharedFile = NM.CreateFileMapping(new IntPtr(-1), sa, NM.PageProtection.ReadWrite, 0, 4096, "DBWIN_BUFFER");
                if (m_SharedFile == IntPtr.Zero)
                {
                    throw CreateApplicationException("Failed to create a file mapping to slot 'DBWIN_BUFFER'");
                }

                // Create a view for this file mapping so we can access it
                m_SharedMem = NM.MapViewOfFile(m_SharedFile, NM.FileMapAccess.FileMapRead, 0, 0, new IntPtr(512));
                if (m_SharedMem == IntPtr.Zero)
                {
                    throw CreateApplicationException("Failed to create a mapping view for slot 'DBWIN_BUFFER'");
                }

                //create the file
                m_tw = new StreamWriter(Path.GetDirectoryName(m_Filename) + @"\" + Path.GetFileNameWithoutExtension(m_Filename) + @".txt");

                // Start a new thread where we can capture the output
                // of OutputDebugString calls so we don't block here.
                m_Capturer = new Thread(new ThreadStart(Capture));
                m_Capturer.Start();
            }
        }

        /// <summary>
        /// Captures 
        /// </summary>
        private static void Capture()
        {
            try
            {
                // Everything after the first DWORD is our debugging text
                IntPtr pString = m_SharedMem + Marshal.SizeOf(typeof(int));

                while (true)
                {
                    NM.SetEvent(m_AckEvent);

                    NM.ThreadWaitValue ret = NM.WaitForSingleObject(m_ReadyEvent, Timeout.Infinite);

                    // if we have no capture set it means that someone
                    // called 'Stop()' and is now waiting for us to exit
                    // this endless loop.
                    if (m_Capturer == null)
                    {
                        break;
                    }

                    if (ret == NM.WAIT_OBJECT_0)
                    {
                        // The first DWORD of the shared memory buffer contains
                        // the process ID of the client that sent the debug string.
                        OnOutputDebugString(Marshal.ReadInt32(m_SharedMem), Marshal.PtrToStringAnsi(pString));
                    }
                }

            }
            catch
            {
                throw;
            }
            finally
            {
                Dispose();
            }
        }

        /// <summary>
        /// Output the debug string
        /// </summary>
        private static void OnOutputDebugString(int pid, string text)
        {
            try
            {
                if (m_PID == pid || m_PID == 0)
                {
                    //TODO write it out as html so we can get some basic formating
                    m_tw.Write(DateTime.Now + ": [" + pid + "] " + text);
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// Dispose all resources
        /// </summary>
        private static void Dispose()
        {
            // Close AckEvent
            if (m_AckEvent != IntPtr.Zero)
            {
                if (!NM.CloseHandle(m_AckEvent))
                {
                    throw CreateApplicationException("Failed to close handle for 'AckEvent'");
                }
                m_AckEvent = IntPtr.Zero;
            }

            // Close ReadyEvent
            if (m_ReadyEvent != IntPtr.Zero)
            {
                if (!NM.CloseHandle(m_ReadyEvent))
                {
                    throw CreateApplicationException("Failed to close handle for 'ReadyEvent'");
                }
                m_ReadyEvent = IntPtr.Zero;
            }

            // Close SharedFile
            if (m_SharedFile != IntPtr.Zero)
            {
                if (!NM.CloseHandle(m_SharedFile))
                {
                    throw CreateApplicationException("Failed to close handle for 'SharedFile'");
                }
                m_SharedFile = IntPtr.Zero;
            }

            // Unmap SharedMem
            if (m_SharedMem != IntPtr.Zero)
            {
                if (!NM.UnmapViewOfFile(m_SharedMem))
                {
                    throw CreateApplicationException("Failed to unmap view for slot 'DBWIN_BUFFER'");
                }
                m_SharedMem = IntPtr.Zero;
            }

            // Close our mutex
            if (m_Mutex != null)
            {
                m_Mutex.Close();
                m_Mutex = null;
            }
        }

        /// <summary>
        /// Stops this debug string capture.
        /// </summary>
        public static void StopCapture()
        {
            lock (m_SyncRoot)
            {
                if (m_Capturer == null)
                {
                    throw new ObjectDisposedException("DebugString", "This DebugString is not running.");
                }
                m_Capturer = null;
                NM.PulseEvent(m_ReadyEvent);
                while (m_AckEvent != IntPtr.Zero)
                {
                    Thread.Yield();
                }
            }

            //compress the file
            m_tw.Close();

            using (FileStream fs = new FileStream(Path.GetDirectoryName(m_Filename) + @"\" + Path.GetFileNameWithoutExtension(m_Filename) + @".zip", FileMode.Create))
            {
                using (ZipArchive zipfile = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    zipfile.CreateEntryFromFile(Path.GetDirectoryName(m_Filename) + @"\" + Path.GetFileNameWithoutExtension(m_Filename) + @".txt", Path.GetFileNameWithoutExtension(m_Filename) + @".txt", CompressionLevel.Optimal);
                }
            }

            File.Delete(Path.GetDirectoryName(m_Filename) + @"\" + Path.GetFileNameWithoutExtension(m_Filename) + @".txt");
        }

        /// <summary>
        /// Helper to create a new application exception, which has automaticly the 
        /// last win 32 error code appended.
        /// </summary>
        /// <param name="text">text of the exception</param>
        private static ApplicationException CreateApplicationException(string text)
        {
            if (text == null || text.Length < 1)
            {
                throw new ArgumentNullException("text", "'text' may not be empty or null.");
            }

            return new ApplicationException(string.Format("{0}. Last Error was {1}", text, Marshal.GetLastWin32Error()));
        }
    }
}
