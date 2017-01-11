using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Threading_Example_3
{
    class ThreadTest
    {
        static bool done;
        static readonly object locker = new object();
        static void Main()
        {
            new Thread (Go).Start();
            Go();
            Console.ReadLine();
        }

        // Note that Go is now an instance method
        static void Go()
        {
            lock (locker)
            {
                if (!done)
                {
                    Console.WriteLine("Done");
                    done = true;
                }
            }
        }
    }
}
