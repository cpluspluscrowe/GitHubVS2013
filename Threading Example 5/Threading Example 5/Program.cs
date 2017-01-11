using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Threading_Example_5
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread.CurrentThread.Name = "Main";
            Thread worker = new Thread(Go);
            worker.Name = "worker";
            worker.Start();
            Thread.Sleep(500);
            Go();
            Console.ReadLine();
        }
        static void Go()
        {
            Console.WriteLine("Hello from " + Thread.CurrentThread.Name);
        }
    }
}
