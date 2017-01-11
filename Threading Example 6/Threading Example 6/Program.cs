using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Linq;

namespace Threading_Example_6
{
    class Program
    {
        static void Main(string[] args)
        {
            int cnt = 0;
            double[] counter = new double[10];
            for (int r = 0; r <= 0; r++)
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                for (int p = 0; p <= 100; p++)
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                }
                watch.Stop();
                var elapsedMs = watch.ElapsedMilliseconds;
                counter[cnt] = elapsedMs;
                cnt += 1;
                Console.WriteLine(elapsedMs);
            }


            counter = new double[10];
            cnt = 0;
            for (int r = 0; r <= 9; r++)
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                Thread t0 = new Thread(() => {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t0.Start();
                Thread t1 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t1.Start();
                Thread t2 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t2.Start();
                Thread t3 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t3.Start();
                Thread t4 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t4.Start();
                Thread t5 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t5.Start();
                Thread t6 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t6.Start();
                Thread t7 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t7.Start();
                Thread t8 = new Thread(() =>
                {
                    Example ex = new Example(1, 1);
                    ex.test();
                });
                t8.Start();
                
                t0.Join();
                t1.Join();
                t2.Join();
                t3.Join();
                t4.Join();
                t5.Join();
                t6.Join();
                t7.Join();
                t8.Join();


                watch.Stop();
                var elapsedMs = watch.ElapsedMilliseconds;
                counter[cnt] = elapsedMs;
                cnt += 1;
                Console.WriteLine(elapsedMs);
            }
            Console.WriteLine(counter.Average());
            Console.ReadLine();
        }
    }
    class Example
    {
        public int result;
        public Example(int i, int j){
            this.result = i * j;
        }
        public void test(){
                for (int i = 0; i <= 10000; i++)
                {
                    for (int j = 0; j <= 10000; j++)
                    {
                        
                    }
                }
    }
    }
}
