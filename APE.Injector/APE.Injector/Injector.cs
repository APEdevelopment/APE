using System;
using System.Runtime.InteropServices;
using System.Diagnostics;
using APE.Syringe;
using System.Windows.Forms;

namespace APE.Injector
{
    class Injector
    {
        static int Main(string[] args)
        {
            if (args.Length != 4)
            {
                Console.WriteLine("<Target PID>, <Assembly>, <Method>, <Source PID>");
                return -1;
            }
            else
            {
                //Console.WriteLine(args[0]);
                //Console.WriteLine(args[1]);
                //Console.WriteLine(args[2]);
                //Console.WriteLine(args[3]);

                int TargetPID = int.Parse(args[0]);
                string Assembly = args[1];
                string Method = args[2];
                int SourcePID = int.Parse(args[3]);

                Process p = Process.GetProcessById(TargetPID);
                return (int)Needle.Inject(p, SourcePID, Assembly, Method);
            }
        }
    }
}
