using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

namespace Threading_Example_2
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread t = new Thread(WriteY);
            t.Start();
            for (int i = 0; i < 10000; i++)
            {
                Console.Write("x");
            }
            Console.ReadLine();
        }
        static void WriteY()
        {
            for (int i = 0; i < 10000; i++)
            {
                Console.Write("y");
            }
        }
    }
}
