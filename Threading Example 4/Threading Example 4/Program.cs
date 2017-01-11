using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Threading_Example_4
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread t = new Thread(Go);
            Thread.Sleep(5000);
            t.Start();
            t.Join();
            Console.WriteLine("Thread t has ended!");
            Console.ReadLine();
        }
        static void Go()
        {
            for (int i = 0; i < 1000; i++)
            {
                Console.Write('y');
            }
        }
    }
}
