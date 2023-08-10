using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System;
using System.Linq;
using FluentDragDrop;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Meter.Forms;
using System.Globalization;
using System.IO.Compression;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using System.IO.Pipes;
using System.Collections.Concurrent;
using Newtonsoft.Json;
using System.Text;

namespace Meter
{
    class MyApplicationContext : ApplicationContext
    {
        [DllImport( "user32.dll" )]
        private static extern int ShowWindow( IntPtr hWnd, uint Msg );

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        private const uint SW_RESTORE = 0x09;

        public static MyApplicationContext instance;

        public static List<Form> menues;
        public static List<string> filesToDelete;
        
        public Excel.Application xlApp;
        public Excel.Workbook wb;
        public Excel.Worksheet wsCh, wsDb, wsMTEP, wsTEPm, wsTEPn;
        public bool closed, excelClosed;
        public IntPtr xlAppHwnd;
        public Rect xlAppRect;
        public NewMenuBase menu;
        public ColorsData colors;
        public double zoom;
        public RangeReferences references;
        public HeadReferences heads;
        public Formula formulas;
        public static bool dontsave;
        public static string dir;
        public static List<int> menuIndexes = new List<int>();
        string file;
        bool restarted;
        public static bool loading = false;
        private static object[,] oldValsArray;
        private static string oldVal;
        public static bool PipeServerActive = true;
        Thread meterServer;
        Thread meterServerWriter;

        static ConcurrentQueue<string> serverMessagesQueue;
        static EventWaitHandle waitHandle = new EventWaitHandle(false, EventResetMode.AutoReset);


        public Excel.WorkbookEvents_BeforeCloseEventHandler Event_BeforeClose;
        public Excel.WorkbookEvents_WindowResizeEventHandler Event_WindowResize;
        public Excel.DocEvents_BeforeRightClickEventHandler Events_BeforeRightClick;
        public Excel.DocEvents_DeactivateEventHandler Events_DeactivateSheet;
        public Excel.WorkbookEvents_SheetActivateEventHandler Events_ActivateSheet;
        public Excel.DocEvents_BeforeDoubleClickEventHandler Events_BeforeDoubleClick;
        public Excel.DocEvents_BeforeDoubleClickEventHandler Events_BeforeDoubleClick_wsMTEP;
        public Excel.DocEvents_ChangeEventHandler Events_Change;
        public Excel.DocEvents_SelectionChangeEventHandler Events_SelectionChange;
        public Excel.WorkbookEvents_BeforeSaveEventHandler Events_BeforeSave;
        public Excel.WorkbookEvents_AfterSaveEventHandler Events_AfterSave;
        public Excel.WorkbookEvents_SheetSelectionChangeEventHandler Events_SheetSelectionChange;
        public Excel.WorkbookEvents_SheetChangeEventHandler Events_SheetChange;

        private void onFormClosed(object sender, EventArgs e)
        {
            if (System.Windows.Forms.Application.OpenForms.Count == 0)
            {
                GlobalMethods.ToLog("Остановка сервера Meter...");
                StopNamedPipe();
                meterServer.Join();
                meterServerWriter.Join();
                GlobalMethods.ToLog("Cервер Meter остановлен");

                var tasks = new List<Task>();

                GlobalMethods.ToLog("Инициализация закрытия книги...");
                if (dontsave == false)
                {
                    SaveBeforeClose();
                    if (filesToDelete.Count != 0)
                    {
                        foreach (string file in filesToDelete)
                        {
                            File.Delete(file);
                        }
                    }
                }
                else
                    wb.Saved = true;

                var task = Task.Run(() => ReleaseAllComObjects());
                tasks.Add(task);

                task = Task.Run(() => ClearEvents());
                tasks.Add(task);

                Task.WaitAll(tasks.ToArray());
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wsCh);
                Marshal.ReleaseComObject(wsDb);
                Marshal.ReleaseComObject(wsMTEP);
                Marshal.ReleaseComObject(wsTEPm);
                Marshal.ReleaseComObject(wsTEPn);
                wb = null;
                wsCh = null;
                wsDb = null;
                wsMTEP = null;
                wsTEPm = null;
                wsTEPn = null;

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                GlobalMethods.ToLog("Счетчики закрыты");
                if (restarted == false) ExitThread();
            }
        }

        public MyApplicationContext()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            meterServer = new Thread(StartNamedPipe);
            meterServer.Start();

            meterServerWriter = new Thread(ProcessQueue);
            meterServerWriter.Start();

            instance = this;
            //If WinForms exposed a global event that fires whenever a new Form is created,
            //we could use that event to register for the form's `FormClosed` event.
            //Without such a global event, we have to register each Form when it is created
            //This means that any forms created outside of the ApplicationContext will not prevent the 
            //application close.
            dontsave = false;

            GlobalMethods.dpiX = Graphics.FromHwnd(IntPtr.Zero).DpiX;
            GlobalMethods.dpiY = Graphics.FromHwnd(IntPtr.Zero).DpiY;
            
            Start();

        }

        void ProcessQueue()
        {
            while (PipeServerActive == true)
            {
                string? msg = null;
                lock (serverMessagesQueue)
                {
                    if (serverMessagesQueue.Count == 0)
                    {
                        waitHandle.WaitOne();
                        continue;
                    }
                    serverMessagesQueue.TryDequeue(out msg);
                }
                if (msg != "Stop Meter Server")
                {
                    // DataWriter.Write(msg);
                    #region Old
                    try
                    {
                        PipeValue pv = JsonConvert.DeserializeObject<PipeValue>(msg);
                        if (!string.IsNullOrEmpty(pv.subjectName) && !string.IsNullOrEmpty(pv.level1Name) && !string.IsNullOrEmpty(pv.level2Name) && pv.day != null && !string.IsNullOrEmpty(pv.value))
                        {
                            ReferenceObject ro = null;
                            if (references.references.TryGetValue(pv.subjectName, out ro))
                            {
                                if (ro != null)
                                {
                                    if (ro.DB.childs.ContainsKey(pv.level1Name) && ro.DB.childs[pv.level1Name].childs.ContainsKey(pv.level2Name))
                                    {
                                        ro.WriteToDB(pv.level1Name, pv.level2Name, (int)pv.day, pv.value.Replace(",", "."));
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
                            if (pv.cod != null && pv.day != null && pv.value != null)
                            {
                                ReferenceObject ro = references.references.Values.AsParallel().Where(n => n.codPlan == pv.cod).FirstOrDefault();
                                if (ro != null)
                                {
                                    ro.WriteToDB("план", "утвержденный", (int)pv.day, pv.value.Replace(",","."));
                                }
                                else
                                {
                                    GlobalMethods.ToLog("Не найден субъект с кодом плана " + pv.cod);
                                }
                            }
                        }
                        else
                        {
                            GlobalMethods.ToLog("Не достаточно данных для записи");
                        }
                    }
                    catch
                    {
                        GlobalMethods.ToLog("Ошибка записи для полученных данных " + msg);
                    }
                    #endregion
                }
                else
                {
                    GlobalMethods.ToLog("Stopping meterServerWriter");
                    serverMessagesQueue.Clear();
                    serverMessagesQueue = null;
                }
            }
        }

        private void StartNamedPipe()
        {
            serverMessagesQueue = new ConcurrentQueue<string>();
            string? msg = null;
            while (PipeServerActive == true)
            {
                using (NamedPipeServerStream pipeServer = new NamedPipeServerStream("MeterServer"))
                {
                    pipeServer.WaitForConnection();
                    using (StreamReader sr = new StreamReader(pipeServer, Encoding.GetEncoding("windows-1251")))
                    {
                        msg = sr.ReadLine();
                        if (msg != "check")
                        {
                            serverMessagesQueue.Enqueue(msg);
                            waitHandle.Set();
                            msg = null;
                        }
                    }
                }
            }
            serverMessagesQueue.Enqueue("Stop Meter Server");
            waitHandle.Set();
            GlobalMethods.ToLog("Stopping meterServer");
        }

        private void EnqueueNewMsg(string msg)
        {
            
        }

        private void StopNamedPipe()
        {
            PipeServerActive = false;
            using (NamedPipeClientStream pipeServer = new NamedPipeClientStream("MeterServer"))
            {
                pipeServer.Connect();
                using (StreamWriter sw = new StreamWriter(pipeServer))
                {
                    sw.WriteLine("Stop Meter Server");

                }
            }
        }

        private void Start()
        {
            restarted = false;
            string file1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\db.txt";
            if (File.Exists(file1))
            {
                dir = File.ReadAllText(file1);
            }

            // dir = Process.GetCurrentProcess().MainModule.FileName ;
            // dir = System.IO.Path.GetDirectoryName(dir) + @"\DB"; 
            file = dir + @"\current\meter.xlsx";
            if (!File.Exists(file))
            {
                dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
                dir = System.IO.Path.GetDirectoryName(dir) + @"\DB"; 
                file = dir + @"\current\meter.xlsx";
            }
            
            InitExcel();
            InitForms();
            InitExcelEvents();
            menu.ClearContextMenu();
            filesToDelete = new List<string>();
        }
        
        public void OpenMonth(string thisYear, string thisMonth, string selectedYear, string selectedMonth, string file, bool silent = false)
        {
            if (silent) xlApp.Visible = false;

            if (thisMonth == DateTime.Today.ToString("MMMM", GlobalMethods.culture) && thisYear == DateTime.Today.ToString("yyyy"))
            {
                ArhivateNew(thisYear, thisMonth);
            }
            GlobalMethods.ToLog("Загрузка архива за " + selectedMonth + " " + selectedYear + " из файла " + file);
            if (menu.InvokeRequired)
            {
                menu.Invoke(new MethodInvoker(() => 
                {
                    OpenMonthMethod(thisYear, thisMonth, selectedYear, selectedMonth, file);
                }));
            }
            else
            {
                OpenMonthMethod(thisYear, thisMonth, selectedYear, selectedMonth, file);
            }
            GlobalMethods.ToLog("Открыты счетчики за " + selectedMonth + " " + selectedYear);
        }

        private void OpenMonthMethod(string thisYear, string thisMonth, string selectedYear, string selectedMonth, string file)
        {
            StopAll();
            string sourceFolder = dir + @"\TEMP";
            if (Directory.Exists(sourceFolder))
            {
                Directory.Delete(sourceFolder, true);
                Directory.CreateDirectory(sourceFolder);
            }
            System.IO.Compression.ZipFile.ExtractToDirectory(file, sourceFolder, System.Text.Encoding.UTF8, true);

            
            xlApp.EnableEvents = false;
            Excel.Application xlApp1;
            Excel.Workbook wb1;
            Excel.Worksheet wsCh1 = null, wsDb1 = null;
            Excel.Range destinationRange;
        
            wb1 = xlApp.Workbooks.Open(Filename: sourceFolder + @"\" + selectedMonth + @".xlsx");
            wb1.Activate();
            wb1.Windows[1].Visible = false;
            foreach (Excel.Worksheet ws in wb1.Worksheets)
            {
                if (ws.CodeName == "PS")
                {
                    wsCh1 = ws;
                }
                if (ws.CodeName == "DB")
                {
                    wsDb1 = ws;
                }
            }

            wsCh1.Cells.Copy(wsCh.Cells);
            wsDb1.Cells.Copy(wsDb.Cells);

            //xlApp.DisplayAlerts = false;
            //wsCh.Delete();
            //wsDb.Delete();
            //xlApp.DisplayAlerts = true;
            //wsCh = null;
            //wsDb = null;

            //wsDb1.Copy(After: wb.Worksheets[1]);
            //wsDb = wb.Worksheets[2] as Excel.Worksheet;
            //wsCh1.Copy(After: wb.Worksheets[1]);
            //wsCh = wb.Worksheets[2] as Excel.Worksheet;


            //foreach (Excel.Worksheet ws in wb.Worksheets)
            //{
            //    if (ws.CodeName == "PS")
            //    {
            //        wsCh = ws;
            //    }
            //    if (ws.CodeName == "DB")
            //    {
            //        wsDb = ws;
            //    }
            //}

            wsCh.Cells.Replace( What: "[" + selectedMonth + ".xlsx]", Replacement: "", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);
            wsCh.Activate();
            

            GlobalMethods.ReleseObject(wsCh1);
            GlobalMethods.ReleseObject(wsDb1);
            wb1.Saved = true;
            wb1.Close();
            GlobalMethods.ReleseObject(wb1);

            SaveLoader.LoadAsyncFromFolder("TEMP");
            Directory.Delete(sourceFolder, true);
            ResumeAll();
            ClearEvents();
            InitExcelEvents();
        }

        public void RunOnUiThread(System.Action<string, string, string, string> action, string thisYear, string thisMonth, string selectedMonth, string file)
        {
            if (SynchronizationContext.Current == null) return;
            SynchronizationContext.Current.Post(state => action(thisYear, thisMonth, selectedMonth, file), null);
        }

        public void RunOnUiThread(System.Action<string, string, string> action, string thisYear, string thisMonth, string file)
        {
            if (SynchronizationContext.Current == null) return;
            SynchronizationContext.Current.Post(state => action(thisYear, thisMonth, file), null);
        }

        private void InitExcel()
        {
            excelClosed = false;
            xlApp = null;
            xlApp = new Excel.ApplicationClass();
            xlApp.Visible = false;
            wb = xlApp.Workbooks.Open(file);
            wb.Activate();
            xlApp.Visible = false;
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (ws.CodeName == "PS")
                {
                    wsCh = ws;
                }
                if (ws.CodeName == "DB")
                {
                    wsDb = ws;
                }
                if (ws.CodeName == "WS_MTEP")
                {
                    wsMTEP = ws;
                }
                if (ws.CodeName == "WS_TEPM")
                {
                    wsTEPm = ws;
                }
                if (ws.CodeName == "WS_TEPN")
                {
                    wsTEPn = ws;
                }
            }

            xlAppHwnd = (IntPtr)xlApp.ActiveWindow.Hwnd;
            xlApp.WindowState = Excel.XlWindowState.xlMinimized;

            references = new RangeReferences();
            heads = new HeadReferences();
            formulas = new Formula();
            colors = new ColorsData();
            xlApp.Visible = false;
            SaveLoader.LoadAsync();
            foreach (CommandBar item in xlApp.CommandBars)
            {
                if (item.Name == "Cell")
                {
                    menuIndexes.Add(item.Index);
                }
            }
            xlApp.Visible = true;
            RestoreExcel();

            zoom = (double)xlApp.ActiveWindow.Zoom;
        }
        private void InitForms()
        {
            menues = new List<Form>();
            menues.Add(new NewMenu());
            menues.Add(new NewMenuAdmin());

            foreach (NewMenuBase form in menues)
            {
                form.FormClosed += onFormClosed;
                // form.Disposed += (object sender, EventArgs e) => 
                // {
                //     if (menues.Contains(form))
                //     {
                //         menues.Remove(form);
                //     }
                //     onFormClosed(sender, e);
                // };
            }
            menu = menues[0] as NewMenuBase;
            menu.Show();

            xlApp.Visible = true;
        }
        private void InitExcelEvents()
        {
            closed = false;

            xlAppRect = new Rect();

            Event_BeforeClose = new Excel.WorkbookEvents_BeforeCloseEventHandler(Application_BeforeClose);
            wb.BeforeClose += Event_BeforeClose;

            Event_WindowResize = new Excel.WorkbookEvents_WindowResizeEventHandler(Application_WindowResize);
            wb.WindowResize += Event_WindowResize;

            Events_BeforeRightClick = new Excel.DocEvents_BeforeRightClickEventHandler(Application_BeforeRightClick);
            wsCh.BeforeRightClick += Events_BeforeRightClick;

            Events_DeactivateSheet = new Excel.DocEvents_DeactivateEventHandler(Application_DeactivateSheet);
            wsCh.Deactivate += Events_DeactivateSheet;

            Events_ActivateSheet = new Excel.WorkbookEvents_SheetActivateEventHandler(Application_ActivateSheet);
            wb.SheetActivate += Events_ActivateSheet;

            // pasteButton = (CommandBarControl)xlApp.CommandBars[9].Controls[12];
            // pasteButton.Enabled = false;
            //CommandBarButtonEvents_ClickEvent= new _CommandBarButtonEvents_ClickEventHandler((CommandBarButton ctrl, ref bool cancelDefault) => { MessageBox.Show("Haha");});
            //pasteButton.OnAction = () => {} ;

            Events_BeforeDoubleClick = new Excel.DocEvents_BeforeDoubleClickEventHandler(Application_BeforeDoubleClick);
            wsCh.BeforeDoubleClick += Events_BeforeDoubleClick;

            Events_BeforeDoubleClick_wsMTEP = new Excel.DocEvents_BeforeDoubleClickEventHandler(Application_BeforeDoubleClick_wsMTEP);
            wsMTEP.BeforeDoubleClick += Events_BeforeDoubleClick_wsMTEP;

            Events_Change = new Excel.DocEvents_ChangeEventHandler(Application_Change);
            wsCh.Change += Events_Change;

            Events_SelectionChange = new Excel.DocEvents_SelectionChangeEventHandler(Application_SelectionChange);
            wsCh.SelectionChange += Events_SelectionChange;

            Events_SheetSelectionChange = new Excel.WorkbookEvents_SheetSelectionChangeEventHandler(Application_SelectionChange);
            wb.SheetSelectionChange += Events_SheetSelectionChange;

            Events_SheetChange = new Excel.WorkbookEvents_SheetChangeEventHandler(Application_Change);
            wb.SheetChange += Events_SheetChange;

            Events_BeforeSave = new Excel.WorkbookEvents_BeforeSaveEventHandler(Wb_BeforeSave);
            wb.BeforeSave += Events_BeforeSave;

            Events_AfterSave = new Excel.WorkbookEvents_AfterSaveEventHandler(Wb_AfterSave);
            wb.AfterSave += Events_AfterSave;

            xlApp.OnKey("^v", "");

            GlobalMethods.CalculateFormsPositions();
        }
        private void ClearEvents()
        {
            try{wb.BeforeClose -= Event_BeforeClose;}catch{}
            try{wb.WindowResize -= Event_WindowResize;}catch{}
            try{wsCh.BeforeRightClick -= Events_BeforeRightClick;}catch{}
            try{wsCh.Deactivate -= Events_DeactivateSheet;}catch{}
            try{wb.SheetActivate -= Events_ActivateSheet;}catch{}
            try{wsCh.BeforeDoubleClick -= Events_BeforeDoubleClick;}catch{}
            try{wsMTEP.BeforeDoubleClick -= Events_BeforeDoubleClick_wsMTEP;}catch{}
            try{wsCh.Change -= Events_Change;}catch{}
            try{wsCh.SelectionChange -= Events_SelectionChange;}catch{}
            try{wb.SheetSelectionChange -= Events_SheetSelectionChange;}catch{}
            try{wb.SheetChange -= Events_SheetChange;}catch{}
            try{wb.BeforeSave -= Events_BeforeSave;}catch{}
            try{wb.AfterSave -= Events_AfterSave;}catch{}
            //try{pasteButton.Click -= CommandBarButtonEvents_ClickEvent;} catch {}

            Event_BeforeClose = null;
            Event_WindowResize = null;
            Events_BeforeRightClick = null;
            Events_DeactivateSheet = null;
            Events_ActivateSheet = null;
            Events_BeforeDoubleClick = null;
            Events_Change = null;
            Events_SelectionChange = null;
            Events_BeforeSave = null;
        }
        private void Wb_BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            GlobalMethods.ToLog("Книга сохранена по инициативе пользователя (нажата кнопка сохранить)");
            SaveLoader.SaveAsync();
        }
        private void Wb_AfterSave(bool Success)
        {
            Arhivate(true);
        }
        private void SaveWB()
        {
            xlApp.EnableEvents = false;
            GlobalMethods.ToLog("Книга сохранена");
            wb.Save();
            SaveLoader.SaveAsync();
            xlApp.EnableEvents = true;
        }

        public void Application_BeforeClose(ref bool cancel)
        {
            if (!closed)
            {
                closed = true;
                foreach (NewMenuBase form in menues)
                {
                    form.FormClose();
                }
            }
        }
        private void Application_WindowResize(Excel.Window wn)
        {
            GlobalMethods.CalculateFormsPositions();
        }
        private void Application_BeforeRightClick(Excel.Range range, ref bool cancel)
        {
            menu.RightClick(range);
        }
        private void Application_DeactivateSheet()
        {
            menu.DeactivateSheet();
        }
        private void Application_ActivateSheet(object sh)
        {
            GlobalMethods.ToLog("Активирован лист '" + ((Excel.Worksheet)sh).Name + "'");
            menu.ActivateSheet(sh);
        }
        private void Application_BeforeDoubleClick(Excel.Range range, ref bool cancel)
        {

            GlobalMethods.ToLog("Двойной клик на ячейке " + range.Address + " листа " + wsCh.Name);
            menu.DblClick(range);
        }
        private void Application_BeforeDoubleClick_wsMTEP(Excel.Range range, ref bool cancel)
        {
            GlobalMethods.ToLog("Двойной клик на ячейке " + range.Address + " листа " + wsMTEP.Name);
            if (range.Column == 1)
            {
                if (ColorsData.GetRangeColor(range) == Color.GreenYellow)
                {
                    try
                    {
                        double val = (double)range.Value;
                        int? cod = null;
                        cod = Convert.ToInt32(((double)val));
                        if (cod != null)
                        {
                            ChildObject co = references.references.Values.SelectMany(n => n.PS.childs.Values).Where(m => m.codMaketTEP == cod).FirstOrDefault();
                            if (co != null)
                            {
                                co.ws.Activate();
                                co.Range.Select();
                                NewMenuBase.SetForegroundWindow(xlAppHwnd);
                            }
                        }
                    }
                    catch
                    {
                        
                    }
                }
            }
        }
        
        private void Application_SelectionChange(Excel.Range range)
        {
            GlobalMethods.ToLog("Выделены ячейки " + range.Address);
            if (range.Formula is string)
            {
                oldVal = (string)range.Formula;
            }
            else if (range.Formula is object[,])
            {
                oldValsArray = (object[,])range.Formula;
            }
            menu.SlectionChanged(range);
        }
        private void Application_SelectionChange(object sh, Excel.Range range)
        {
            if (((Excel.Worksheet)sh).CodeName != "PS")
            {
                try
                {
                    GlobalMethods.ToLog("Выделены ячейки " + range.Address);
                    if (range.Formula is string)
                    {
                        oldVal = (string)range.Formula;
                    }
                    else if (range.Formula is object[,])
                    {
                        oldValsArray = (object[,])range.Formula;
                    }
                }
                catch
                {

                }
            }
        }
        private void Application_Change(Excel.Range range)
        {
            if (range.Cells.Count > 1)
            {
                ChagedRange(wsCh, range);
            }
            else
            {
                Changed(wsCh, range);
            }
            menu.CellValueChanged(range);
        }
        private void Application_Change(object sh, Excel.Range range)
        {
            if (((Excel.Worksheet)sh).CodeName != "PS")
            {
                if (range.Cells.Count > 1)
                {
                    ChagedRange(sh, range);
                }
                else
                {
                    Changed(sh, range);
                }
            }
        }
        private void ChagedRange(object sh, Excel.Range rng)
        {
            if (rng.Cells.Count > 1000)
                return;
            object[,] newValsArray = (object[,])rng.Formula;
            for (int i = 1; i <= rng.Columns.Count; i++)
            {
                for (int j = 1; j <= rng.Rows.Count; j++)
                {
                    try
                    {
                        GlobalMethods.ToLog("Изменено значение ячейки " + ((Excel.Range)rng.Cells[j,i]).Address + " на листе " + ((Excel.Worksheet)sh).Name + " с '" + oldValsArray[j,i] + "' на '" + newValsArray[j,i] + "'");
                    }
                    catch
                    {
                        GlobalMethods.ToLog("Err");
                    }
                    
                }
            }
        }
        private void Changed(object sh, Excel.Range rng)
        {
            GlobalMethods.ToLog("Изменено значение ячейки " + rng.Address + " на листе " + ((Excel.Worksheet)sh).Name + " с '" + oldVal + "' на '" + rng.Value + "'");
        }
        private void RestoreExcel()
        {
            if (xlApp.WindowState == Excel.XlWindowState.xlMinimized)
            {
                ShowWindow(xlAppHwnd, SW_RESTORE);
            }
        }
        public void StopAll()
        {
            xlApp.Calculation = XlCalculation.xlCalculationManual;
            xlApp.ScreenUpdating = false;
            xlApp.EnableEvents = false;
            xlApp.DisplayStatusBar = false;
        }
        public void ResumeAll()
        {
            xlApp.Calculation = XlCalculation.xlCalculationAutomatic;
            xlApp.ScreenUpdating = true;
            xlApp.EnableEvents = true;
            xlApp.DisplayStatusBar = true;
        }

        protected override void OnMainFormClosed(object? sender, EventArgs e)
        {
            base.OnMainFormClosed(sender, e);

            MessageBox.Show("Closed");
        }
        public void Arhivate(bool withoutSave = false)
        {
            string month, year;
            month = menu.lblMonth.Text;
            year = menu.lblYear.Text;
            ArhivateNew(year, month, withoutSave);
            ArchivateToTemp();
        }
        private void ArhivateNew(string year, string month, bool withoutSave = false)
        {
            if (withoutSave == false) SaveWB();
            string sourceFolder = dir + @"\current";
            string tempDirectory = Path.Combine(Path.GetTempPath(), DateTime.Today.ToString("MMMM", GlobalMethods.culture));
            Directory.CreateDirectory(tempDirectory);
            foreach (string dirPath in Directory.GetDirectories(sourceFolder, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourceFolder, tempDirectory));
            }
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*", SearchOption.AllDirectories))
            {
                string newFilePath = filePath.Replace(sourceFolder, tempDirectory);
                if (newFilePath.Contains("meter"))
                {
                    newFilePath = newFilePath.Replace("meter", month);
                }
                if (newFilePath.Contains("~$"))
                {
                    continue;
                }
                File.Copy(filePath, newFilePath, true);
            }
            string archPath = dir + @"\arch";
            if (!Directory.Exists(archPath))
            {
                Directory.CreateDirectory(archPath);
            }
            archPath = archPath + @"\" + year;
            if (!Directory.Exists(archPath))
            {
                Directory.CreateDirectory(archPath);
            }
            string arhiveName = archPath + @"\" + month + @".zip";
            if (File.Exists(arhiveName))
            {
                File.Delete(arhiveName);
            }
            ZipFile.CreateFromDirectory(tempDirectory, arhiveName, CompressionLevel.Fastest, false, System.Text.Encoding.UTF8);
            Directory.Delete(tempDirectory, true);
            GlobalMethods.ToLog("Книга архивирована (" + month + " " + year + " года) в файл " + arhiveName);
        }
        private void SaveBeforeClose()
        {
            string thisYear, thisMonth, selectedYear, selectedMonth, file;

            thisMonth = menu.lblMonth.Text;
            thisYear = menu.lblYear.Text;
            selectedMonth = DateTime.Today.ToString("MMMM", GlobalMethods.culture);
            selectedYear = DateTime.Today.ToString("yyyy");
            file = dir + @"\arch\" + selectedYear + @"\" + selectedMonth + @".zip";

            SaveWB();
            string sourceFolder = dir + @"\current";
            string tempDirectory = Path.Combine(Path.GetTempPath(), GlobalMethods.username + " " + DateTime.Today.ToString("MMMM", GlobalMethods.culture));
            Directory.CreateDirectory(tempDirectory);
            foreach (string dirPath in Directory.GetDirectories(sourceFolder, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourceFolder, tempDirectory));
            }
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*", SearchOption.AllDirectories))
            {
                string newFilePath = filePath.Replace(sourceFolder, tempDirectory);
                if (newFilePath.Contains("~$"))
                {
                    continue;
                }
                File.Copy(filePath, newFilePath, true);
            }
            string archPath = dir + @"\temparch";
            if (!Directory.Exists(archPath))
            {
                Directory.CreateDirectory(archPath);
            }
            string arhiveName = archPath + @"\" + DateTime.Now.ToString("dd'.'MM'.'yyyy HH'.'mm'.'ss") + @".zip";
            ZipFile.CreateFromDirectory(tempDirectory, arhiveName, CompressionLevel.Fastest, false, System.Text.Encoding.UTF8);
            Directory.Delete(tempDirectory, true);

            DirectoryInfo directoryInfo = new DirectoryInfo(archPath);
            FileInfo[] files = directoryInfo.GetFiles();
            if (files.Length > 10)
            {
                FileInfo oldestFile = files.OrderBy(f => f.LastWriteTime).First();
                oldestFile.Delete();
                GlobalMethods.ToLog("Удален временный архив " + oldestFile);
            }

            GlobalMethods.ToLog("Книга архивирована в файл " + arhiveName);

            if (thisMonth != selectedMonth || thisYear != selectedYear) 
            {
                OpenMonth(thisYear, thisMonth, selectedYear, selectedMonth, file);
                SaveWB();
            }
        }

        private void ArchivateToTemp()
        {
            string thisYear, thisMonth, selectedYear, selectedMonth, file;
            
            thisMonth = menu.lblMonth.Text;
            thisYear = menu.lblYear.Text;
            selectedMonth = DateTime.Today.ToString("MMMM", GlobalMethods.culture);
            selectedYear = DateTime.Today.ToString("yyyy");
            file = dir + @"\arch\" + selectedYear + @"\" + selectedMonth + @".zip";

            string sourceFolder = dir + @"\current";
            string tempDirectory = Path.Combine(Path.GetTempPath(), GlobalMethods.username + " " + DateTime.Today.ToString("MMMM", GlobalMethods.culture));
            Directory.CreateDirectory(tempDirectory);
            foreach (string dirPath in Directory.GetDirectories(sourceFolder, "*", SearchOption.AllDirectories))
            {
                Directory.CreateDirectory(dirPath.Replace(sourceFolder, tempDirectory));
            }
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*", SearchOption.AllDirectories))
            {
                string newFilePath = filePath.Replace(sourceFolder, tempDirectory);
                if (newFilePath.Contains("~$"))
                {
                    continue;
                }
                File.Copy(filePath, newFilePath, true);
            }
            string archPath = dir + @"\temparch";
            if (!Directory.Exists(archPath))
            {
                Directory.CreateDirectory(archPath);
            }
            string arhiveName = archPath + @"\" + DateTime.Now.ToString("dd'.'MM'.'yyyy HH'.'mm'.'ss") + @".zip";
            ZipFile.CreateFromDirectory(tempDirectory, arhiveName, CompressionLevel.Fastest, false, System.Text.Encoding.UTF8);
            Directory.Delete(tempDirectory, true);

            DirectoryInfo directoryInfo = new DirectoryInfo(archPath);
            FileInfo[] files = directoryInfo.GetFiles();
            if (files.Length > 50)
            {
                FileInfo oldestFile = files.OrderBy(f => f.LastWriteTime).First();
                oldestFile.Delete();
                GlobalMethods.ToLog("Удален временный архив " + oldestFile);
            }

            GlobalMethods.ToLog("Книга архивирована в файл " + arhiveName);
        }
        private void ReleaseAllComObjects()
        {
            references.ReleaseAllComObjects();
            heads.ReleaseAllComObjects();
        }
    }
}
