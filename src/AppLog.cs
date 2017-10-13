using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Text;

namespace ExcelConverter
{
    public class AppLog : Singleton<AppLog>
    {
        private AppLog() { }

        object QueueLock = new object();
        ConcurrentQueue<string> Queue = new ConcurrentQueue<string>();

        public int QueueLenMax { get; set; } = 1000;

        public Action<string> Write { get; private set; } = Console.WriteLine;

        public void Log(string log)
        {
            if (Queue.Count >= QueueLenMax)
            {
                string discard = null;
                while (!Queue.TryDequeue(out discard))
                    ;
            }
            Queue.Enqueue(log);
            Write(log);
        }
    }
}
