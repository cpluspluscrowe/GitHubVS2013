using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace random
{
    class Program
    {
        static void Main(string[] args)
        {
            Random r = new Random();
            for (int i = 0; i <= 300; i++)
            {
                Console.Write(r.Next(5, 10).ToString() + " ");
            }
            Console.ReadLine();
        }
    }
}
