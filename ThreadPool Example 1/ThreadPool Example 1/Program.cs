using System;
using System.Threading;
using System.Collections.Generic;
public static class ThreadExtension
{
    public static void WaitAll(this IEnumerable<Thread> threads)
    {
        if (threads != null)
        {
            foreach (Thread thread in threads)
            { thread.Join(); }
        }
    }
}
public class Example
{
    public int result;
    public Example(int i, int j)
    {
        this.result = i * j;
    }
}
public class ThreadPoolExample
{
    static void ex()
    {
        for (int i = 0; i <= 10000; i++)
        {
            for (int j = 0; j <= 10000; j++)
            {
                for (int c = 0; c <= 10; c++)
                {
                    Example ex = new Example(i, j);
                }
            }
        }
    }
    static void Main()
    {
        for (int val = 0; val <= 50; val++)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            for (int p = 0; p <= val; p++)
            {
                ex();
            }
            watch.Stop();
            double elapsedMs1 = watch.ElapsedMilliseconds;



            watch = System.Diagnostics.Stopwatch.StartNew();
            List<Thread> threads = new List<Thread>();
            for (int p = 0; p <= val; p++)
            {
                Thread t = new Thread(() =>
                {
                    ex();
                });
                t.Start();
                threads.Add(t);
            }
            threads.WaitAll();
            watch.Stop();
            double elapsedMsT = watch.ElapsedMilliseconds;
            Console.WriteLine(val + ":    " + elapsedMs1/elapsedMsT);

        }
        Console.ReadLine();
    }
}