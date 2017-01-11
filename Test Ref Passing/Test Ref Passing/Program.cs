using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_Ref_Passing
{
    class Example
    {
        public string S1;
        public List<string> L1;
        public Example(string s1)
        {
            this.S1 = s1;
            L1 = new List<string>();
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Example ex = new Example("Created");
            ex.L1.Add("1");
            ex.L1.Add("2");
            ex.L1.Add("3");
            Changer(ex.L1);
            foreach (string var in ex.L1)
            {
                Console.WriteLine(var);
            }
            Console.Read();
        }

        static void Changer(List<string> L1)
        {
            L1.Remove("2");
        }
    }
}
