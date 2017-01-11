using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;

namespace Search_Excel
{
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
    class Program
    {
        public static Queue<string> queue = new Queue<string>();
        static void Main(string[] args)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            
            List<Thread> threads = new List<Thread>();
            string facilitiesPath = "C:\\Users\\CCrowe\\Documents\\Facilities";
            DirectoryInfo di = new DirectoryInfo(facilitiesPath);
            foreach (var file in di.GetFiles())
            {
                queue.Enqueue(file.FullName);
            }
            for (int i = 0; i <= 3; i++)
            {
                Thread thread = new Thread(() =>
                {
                    searchFiles();
                });
                thread.Start();
                threads.Add(thread);
            }
            threads.WaitAll();
            watch.Stop();
            var elapsedMs = watch.ElapsedMilliseconds;
            Console.WriteLine(elapsedMs);
            Console.ReadLine();
        }
        static void searchFiles()
        {
            Excel.Application xl = new Excel.Application();
            while (queue.Count != 0)
            {
                string fileName = queue.Dequeue();
                Excel.Workbook wb = xl.Workbooks.Open(fileName);
                if (wb != null)
                {
                    Excel.Worksheet bom = wb.Sheets["Bill of Materials"];
                    for (int i = 2; i <= bom.UsedRange.Rows.Count; i++)
                    {
                        if (bom.Range["G" + i.ToString()].Value != null)
                        {
                            Console.WriteLine(wb.Name + " " + bom.Range["G" + i.ToString()].Value);
                        }
                    }
                    wb.Close(false);
                }
            }
            xl.Quit();
        }
    }
}
