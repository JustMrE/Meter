// using Microsoft.Office.Interop.Excel;
// using Newtonsoft.Json;
// using Main = Meter.MyApplicationContext;

// namespace Meter.Forms
// {
//     public class FileWatcher
//     {
//         private FileSystemWatcher watcher;
//         private readonly string path;

//         public FileWatcher(string path)
//         {
//             this.path = path;
//             watcher = new FileSystemWatcher(path)
//             {
//                 Filter = "*.json",
//                 NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
//             };
//             watcher.Created += OnFileCreated;
//         }

//         public void Start()
//         {
//             watcher.EnableRaisingEvents = true;
//             GlobalMethods.ToLog("FileWatcher started.");

//             // Console.WriteLine("FileWatcher started.");
//         }

//         private void OnFileCreated(object sender, FileSystemEventArgs e)
//         {
//             GlobalMethods.ToLog($"File created: {e.FullPath}");
//             // Console.WriteLine($"File created: {e.FullPath}");
//             // Start a new task to process the file
//             Task.Run(() => ProcessFileIfReady(e.FullPath));
//         }

//         public async Task ProcessFileIfReady(string filePath)
//         {
//             if (IsFileReady(filePath))
//             {
//                 await ProcessFile(filePath);
//             }
//             else
//             {
//                 GlobalMethods.ToLog($"File {filePath} is not ready for processing.");
//             }
//         }

//         private bool IsFileReady(string filePath)
//         {
//             try
//             {
//                 // Attempt to open the file exclusively
//                 using (FileStream inputStream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
//                 {
//                     if (inputStream.Length > 0)
//                     {
//                         return true;
//                     }
//                 }
//             }
//             catch (IOException)
//             {
//                 // The file is unavailable because it is still being written to
//                 return false;
//             }

//             return false;
//         }

//         private async Task ProcessFile(string filePath)
//         {
//             // Read file content asynchronously
//             string content = File.ReadAllText(filePath);
//             // GlobalMethods.ToLog($"File content: {content}");
//             // Console.WriteLine($"File content: {content}");

//             // Deserialize JSON content to an object
//             var data = JsonConvert.DeserializeObject<List<PipeValue>>(content);

//             // Invoke the method on the UI thread
//             // Main.instance.menu.Invoke(new System.Action(() => Main.instance.menu.TestMethod(data)));
//             Write(data);
//             File.Delete(filePath);
//         }

//         static void Write(List<PipeValue> result)
//         {
//             Main.instance.StopAll();
//             foreach (PipeValue pv in result)
//             {
//                 GlobalMethods.ToLog("write to meter: " + pv.subjectName);
//                 if (!string.IsNullOrEmpty(pv.subjectName) && !string.IsNullOrEmpty(pv.level1Name) && !string.IsNullOrEmpty(pv.level2Name) && pv.day != null && !string.IsNullOrEmpty(pv.value))
//                 {
//                     ReferenceObject ro = null;
//                     if (Main.instance.references.references.TryGetValue(pv.subjectName, out ro))
//                     {
//                         if (ro != null)
//                         {
//                             if (ro.DB.childs.ContainsKey(pv.level1Name) && ro.DB.childs[pv.level1Name].childs.ContainsKey(pv.level2Name))
//                             {
//                                 ro.WriteToDB(pv.level1Name, pv.level2Name, int.Parse(pv.day), pv.value.Replace(",", "."));
//                             }
//                         }
//                         else
//                         {
//                             GlobalMethods.ToLog("Не найден субъект " + pv.subjectName);
//                         }
//                     }
//                     else
//                     {
//                         GlobalMethods.ToLog("Не найден субъект " + pv.subjectName);
//                     }
//                 }
//                 else if (pv != null)
//                 {
//                     if (!string.IsNullOrEmpty(pv.cod) && !string.IsNullOrEmpty(pv.day) && !string.IsNullOrEmpty(pv.value))
//                     {
//                         int cod, day;
//                         if (int.TryParse(pv.cod, out cod) && int.TryParse(pv.day, out day))
//                         {
//                             ReferenceObject ro = Main.instance.references.references.Values.AsParallel().Where(n => n.codPlan == cod).FirstOrDefault();
//                             if (ro != null)
//                             {
//                                 ro.WriteToDB("план", "утвержденный", day, pv.value.Replace(",","."));
//                             }
//                             else
//                             {
//                                 GlobalMethods.ToLogError("Не найден субъект с кодом плана " + cod);
//                                 GlobalMethods.Err("Не найден субъект с кодом плана " + cod);
//                             }
//                         }
//                     }
//                 }
//                 else
//                 {
//                     GlobalMethods.ToLog("Не достаточно данных для записи");
//                 }
//             }
//             Main.instance.ResumeAll();
//         }
        
//         public void ProcessExistingFiles()
//         {
//             string[] existingFiles = Directory.GetFiles(path, "*.json");

//             foreach (string filePath in existingFiles)
//             {
//                 Task.Run(() => ProcessFileIfReady(filePath));
//             }
//         }

        
//     }
// }

using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Main = Meter.MyApplicationContext;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;
using System.Collections.Concurrent;

namespace Meter.Forms
{
    public class FileWatcher
    {
        private static ManualResetEventSlim waitHandle = new ();
        public static bool PipeServerActive = true;
        ConcurrentQueue<List<PipeValue>> serverMessagesQueue;
        Thread meterServerWriter;

        private FileSystemWatcher watcher;
        private readonly string path;

        public FileWatcher()
        {
            this.path = "X:\\MeterWorker";
            serverMessagesQueue = new ConcurrentQueue<List<PipeValue>>();
            this.path = path;
            watcher = new FileSystemWatcher(path)
            {
                Filter = "*.json",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
            };
            watcher.Created += OnFileCreated;
        }
        public FileWatcher(string path)
        {
            serverMessagesQueue = new ConcurrentQueue<List<PipeValue>>();
            this.path = path;
            watcher = new FileSystemWatcher(path)
            {
                Filter = "*.json",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
            };
            watcher.Created += OnFileCreated;
        }

        public void Start()
        {
            GlobalMethods.ToLog("Starting FileWatcher.");
            meterServerWriter = new Thread(ProcessQueue);
            meterServerWriter.Start();
            watcher.EnableRaisingEvents = true;
            GlobalMethods.ToLog("FileWatcher started.");
        }

        public void Stop()
        {
            GlobalMethods.ToLog("Stoping FileWatcher.");
            PipeServerActive = false;
            StopProcessQueue();
            meterServerWriter.Join();
            GlobalMethods.ToLog("FileWatcher stoped.");
        }

        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            GlobalMethods.ToLog($"File created: {e.FullPath}");
            Task.Run(() => ProcessFileIfReady(e.FullPath));
        }

        public async Task ProcessFileIfReady(string filePath)
        {
            const int maxAttempts = 10;  // Максимальное количество попыток
            const int delay = 1000;  // Задержка между попытками в миллисекундах

            int attempts = 0;
            while (attempts < maxAttempts)
            {
                if (IsFileReady(filePath))
                {
                    await ProcessFile(filePath);
                    return;
                }

                attempts++;
                GlobalMethods.ToLog($"File {filePath} is not ready for processing. Attempt {attempts} of {maxAttempts}.");
                await Task.Delay(delay);  // Ждем перед повторной попыткой
            }

            GlobalMethods.ToLog($"File {filePath} was not ready after {maxAttempts} attempts.");
        }

        private bool IsFileReady(string filePath)
        {
            try
            {
                using (FileStream inputStream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    if (inputStream.Length > 0)
                    {
                        return true;
                    }
                }
            }
            catch (IOException)
            {
                return false;
            }

            return false;
        }

        private async Task ProcessFile(string filePath)
        {
            try
            {
                string content = await File.ReadAllTextAsync(filePath);
                var data = JsonConvert.DeserializeObject<List<PipeValue>>(content);
                serverMessagesQueue.Enqueue(data);
                waitHandle.Set();
                // Write(data);
                File.Delete(filePath);
            }
            catch (Exception ex)
            {
                GlobalMethods.ToLogError($"Error processing file {filePath}: {ex.Message}");
            }
        }

        static void Write(List<PipeValue> result)
        {
            Main.instance.StopAll();
            foreach (PipeValue pv in result)
            {
                GlobalMethods.ToLog("write to meter: " + pv.subjectName);
                if (!string.IsNullOrEmpty(pv.subjectName) && !string.IsNullOrEmpty(pv.level1Name) && !string.IsNullOrEmpty(pv.level2Name) && pv.day != null && !string.IsNullOrEmpty(pv.value))
                {
                    ReferenceObject ro = null;
                    if (Main.instance.references.references.TryGetValue(pv.subjectName, out ro))
                    {
                        if (ro != null)
                        {
                            if (ro.DB.childs.ContainsKey(pv.level1Name) && ro.DB.childs[pv.level1Name].childs.ContainsKey(pv.level2Name))
                            {
                                ro.WriteToDB(pv.level1Name, pv.level2Name, int.Parse(pv.day), pv.value.Replace(",", "."));
                            }
                        }
                        else
                        {
                            GlobalMethods.ToLog("Не найден субъект " + pv.subjectName);
                        }
                    }
                    else
                    {
                        GlobalMethods.ToLog("Не найден субъект " + pv.subjectName);
                    }
                }
                else if (pv != null)
                {
                    if (!string.IsNullOrEmpty(pv.cod) && !string.IsNullOrEmpty(pv.day) && !string.IsNullOrEmpty(pv.value))
                    {
                        int cod, day;
                        if (int.TryParse(pv.cod, out cod) && int.TryParse(pv.day, out day))
                        {
                            ReferenceObject ro = Main.instance.references.references.Values.AsParallel().Where(n => n.codPlan == cod).FirstOrDefault();
                            if (ro != null)
                            {
                                ro.WriteToDB("план", "утвержденный", day, pv.value.Replace(",","."));
                            }
                            else
                            {
                                GlobalMethods.ToLogError("Не найден субъект с кодом плана " + cod);
                                GlobalMethods.Err("Не найден субъект с кодом плана " + cod);
                            }
                        }
                    }
                }
                else
                {
                    GlobalMethods.ToLog("Не достаточно данных для записи");
                }
            }
            Main.instance.ResumeAll();
        }
        
        public void ProcessExistingFiles()
        {
            string[] existingFiles = Directory.GetFiles(path, "*.json");
            foreach (string filePath in existingFiles)
            {
                Task.Run(() => ProcessFileIfReady(filePath));
            }
        }
    
        void ProcessQueue()
        {
            List<PipeValue> data;

            while (PipeServerActive)
            {
                if (serverMessagesQueue != null && serverMessagesQueue.TryDequeue(out data))
                {
                    Write(data);
                    data.Clear();
                }
                else
                {
                    waitHandle.Wait();
                }
            }
        }
    
        void StopProcessQueue()
        {
            serverMessagesQueue.Clear();
            serverMessagesQueue = null;
        }
    }
}
