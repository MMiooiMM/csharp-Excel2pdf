using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using NLog;

namespace ExcelToPdf
{
    class Program
    {
        static ILogger logger = LogManager.GetCurrentClassLogger();
        static void Main(string[] args)
        {
            int.TryParse(System.Configuration.ConfigurationManager.AppSettings["threadcount"], out int threadcount);
            threadcount = threadcount == 0 ? 2 : threadcount;
            bool isDelete = string.IsNullOrWhiteSpace(System.Configuration.ConfigurationManager.AppSettings["IsDelete"])
                ? true : System.Configuration.ConfigurationManager.AppSettings["IsDelete"] == "Y";

            foreach (var dir in Directory.GetDirectories(Directory.GetCurrentDirectory()))
            {
                Task[] tasks = new Task[threadcount];
                var target = Directory.GetFiles(dir).Where(x => x.Contains("xlsx") && !x.Contains("$"));
                if (target.Count() == 0) continue;
                int index = 0;
                foreach (var files in target.ToSplit((int)Math.Ceiling(target.Count() / (double)threadcount)))
                {
                    tasks[index] = new Task(() => ConvertPDF(files));
                    tasks[index].Start();
                    System.Threading.Thread.Sleep(1 * 1000);
                    index++;
                }
                Task.WaitAll(tasks);
                KillThread();
                if(isDelete)
                {
                    foreach(var file in target)
                    {
                        File.Delete(file);
                    }
                }
            }
        }
        static void KillThread()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcesses().Where(x => x.ProcessName == "EXCEL"))
            {
                try
                {
                    p.Kill();
                }
                catch(Exception e)
                {
                    logger.Error(e.Message);
                }
            }

            //停頓幾秒 太快會沒關掉EXCEL就開始後續行為
            System.Threading.Thread.Sleep(2 * 1000);
        }
        static void ConvertPDF(IEnumerable<string> files)
        {
            Application excelApplication = new Application();
            foreach (var file in files)
            {
                Workbook excelWorkBook = excelApplication.Workbooks.Open(file);
                excelWorkBook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, file.Replace("xlsx", "pdf"));
                excelWorkBook.Close(false, Type.Missing, Type.Missing);
            }
            excelApplication.Quit();
        }
    }

    static class IEnumerableEx
    {
        public static IEnumerable<IEnumerable<T>> ToSplit<T>(this IEnumerable<T> list, int batchCount)
        {
            return Enumerable.Range(0, list.Count())
              .Where(x => x % batchCount == 0)
              .Select(x => {
                  return list.Skip(x).Take(batchCount);
              });
        }
        public static IEnumerable<T> Clone<T>(this IEnumerable<T> listToClone) where T : ICloneable
        {
            return listToClone.Select(item => (T)item.Clone());
        }
    }
}
