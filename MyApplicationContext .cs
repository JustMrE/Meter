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

        // public static List<Form> forms;
        public static List<Form> menues;
        
        public Excel.Application xlApp;
        public Excel.Workbook wb;
        public Excel.Worksheet wsCh, wsDb;
        public bool closed, excelClosed;
        public IntPtr xlAppHwnd;
        public Rect xlAppRect;
        public NewMenuBase menu;
        public ColorsData colors;
        public double zoom;
        public RangeReferences references;
        public HeadReferences heads;
        public Formula formulas;
        public static bool save;
        public static string dir;
        public static List<int> menuIndexes = new List<int>();
        string file;
        bool restarted;
        public static bool loading = false;
        private static object[,] oldValsArray;
        private static string oldVal;

        public Excel.WorkbookEvents_BeforeCloseEventHandler Event_BeforeClose;
        public Excel.WorkbookEvents_WindowResizeEventHandler Event_WindowResize;
        public Excel.DocEvents_BeforeRightClickEventHandler Events_BeforeRightClick;
        public Excel.DocEvents_DeactivateEventHandler Events_DeactivateSheet;
        public Excel.WorkbookEvents_SheetActivateEventHandler Events_ActivateSheet;
        public Excel.DocEvents_BeforeDoubleClickEventHandler Events_BeforeDoubleClick;
        public Excel.DocEvents_ChangeEventHandler Events_Change;
        public Excel.DocEvents_SelectionChangeEventHandler Events_SelectionChange;
        public Excel.WorkbookEvents_BeforeSaveEventHandler Events_BeforeSave;
        public Excel.WorkbookEvents_SheetSelectionChangeEventHandler Events_SheetSelectionChange;
        public Excel.WorkbookEvents_SheetChangeEventHandler Events_SheetChange;

        private void onFormClosed(object sender, EventArgs e)
        {
            if (System.Windows.Forms.Application.OpenForms.Count == 0)
            {
                if (save == true)
                {
                    //SaveLoader.Save();
                    wb.Save();
                }
                else
                {
                    wb.Saved = true;
                }

                references.ReleaseAllComObjects();

                ClearEvents();

                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wsCh);
                Marshal.ReleaseComObject(wsDb);
                wb = null;
                wsCh = null;
                wsDb = null;

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GlobalMethods.ToLog("Счетчики закрыты");
                // Thread.Sleep(10000);
                if (restarted == false) ExitThread();
            }
        }

        public MyApplicationContext()
        {
            instance = this;
            //If WinForms exposed a global event that fires whenever a new Form is created,
            //we could use that event to register for the form's `FormClosed` event.
            //Without such a global event, we have to register each Form when it is created
            //This means that any forms created outside of the ApplicationContext will not prevent the 
            //application close.
            save = false;
            
            Start();
        }

        public void Start()
        {
            restarted = false;
            string file1 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\db.txt";
            if (File.Exists(file1))
            {
                dir = File.ReadAllText(file1);
            }

            // dir = Process.GetCurrentProcess().MainModule.FileName ;
            // dir = System.IO.Path.GetDirectoryName(dir) + @"\DB"; 
            file = dir + @"\current\meter.xlsm";
            if (!File.Exists(file))
            {
                dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
                dir = System.IO.Path.GetDirectoryName(dir) + @"\DB"; 
                file = dir + @"\current\meter.xlsm";
            }
            
            InitExcel();
            InitForms();
            InitExcelEvents();
            menu.ClearContextMenu();
        }

        public void Restart(string thisYear, string thisMonth, string file)
        {
            restarted = true;

            if (menu.InvokeRequired)
            {
                menu.Invoke(new MethodInvoker(() => 
                {
                    wb.Save();
                    foreach (NewMenuBase form in menues)
                    {
                        form.FormClose();
                    }

                    // ClearEvents();
                    // ReleaseAllComObjects();

                    // wb.Save();
                    // wb.Close();

                    // Marshal.ReleaseComObject(wb);
                    // Marshal.ReleaseComObject(wsCh);
                    // Marshal.ReleaseComObject(wsDb);

                    // xlApp.Quit();
                    // Marshal.ReleaseComObject(xlApp);
                    // GC.Collect();
                    // GC.WaitForPendingFinalizers();

                    ArhivateNew(thisYear, thisMonth);
                    string sourceFolder = dir + @"\current";
                    Directory.Delete(sourceFolder, true);
                    System.IO.Compression.ZipFile.ExtractToDirectory(file, sourceFolder, true);
                    Start();
                }));
            }
            else
            {
                wb.Save();
                foreach (NewMenuBase form in menues)
                {
                    form.Close();
                }
                // ClearEvents();
                // ReleaseAllComObjects();

                // wb.Save();
                // wb.Close();

                // Marshal.ReleaseComObject(wb);
                // Marshal.ReleaseComObject(wsCh);
                // Marshal.ReleaseComObject(wsDb);

                // xlApp.Quit();
                // Marshal.ReleaseComObject(xlApp);
                // GC.Collect();
                // GC.WaitForPendingFinalizers();

                ArhivateNew(thisYear, thisMonth);
                string sourceFolder = dir + @"\current";
                Directory.Delete(sourceFolder, true);
                System.IO.Compression.ZipFile.ExtractToDirectory(file, sourceFolder, true);
                Start();
            }
            
        }
        
        public void OpenMonth(string thisYear, string thisMonth, string selectedYear, string selectedMonth, string file)
        {
            ArhivateNew(thisYear, thisMonth);
            GlobalMethods.ToLog("Загрузка архива за " + selectedMonth + " " + selectedYear + " из файла " + file);
            if (menu.InvokeRequired)
            {
                menu.Invoke(new MethodInvoker(() => 
                {
                    string sourceFolder = dir + @"\TEMP";
                    if (Directory.Exists(sourceFolder))
                    {
                        Directory.Delete(sourceFolder, true);
                        Directory.CreateDirectory(sourceFolder);
                    }
                    System.IO.Compression.ZipFile.ExtractToDirectory(file, sourceFolder, true);

                    StopAll();
                    xlApp.EnableEvents = false;
                    Excel.Application xlApp1;
                    Excel.Workbook wb1;
                    Excel.Worksheet wsCh1 = null, wsDb1 = null;
                    Excel.Range destinationRange;
                
                    wb1 = xlApp.Workbooks.Open(Filename: sourceFolder + @"\" + selectedMonth + @".xlsm");
                    wb1.Activate();
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

                    xlApp.DisplayAlerts = false;
                    wsCh.Delete();
                    wsDb.Delete();
                    xlApp.DisplayAlerts = true;
                    wsCh = null;
                    wsDb = null;

                    wsDb1.Copy(Type.Missing, wb.Worksheets[1]);
                    wsCh1.Copy(Type.Missing, wb.Worksheets[1]);
                    
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
                    }

                    wsCh.Cells.Replace( What: "[" + selectedMonth + ".xlsm]", Replacement: "", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);
                    wsCh.Activate();
                    ResumeAll();

                    GlobalMethods.ReleseObject(wsCh1);
                    GlobalMethods.ReleseObject(wsDb1);
                    wb1.Saved = true;
                    wb1.Close();
                    GlobalMethods.ReleseObject(wb1);

                    SaveLoader.LoadAsyncFromFolder("TEMP");
                    //references.UpdateAllWSs();
                    Directory.Delete(sourceFolder, true);
                    InitExcelEvents();
                }));
            }
            else
            {
                string sourceFolder = dir + @"\TEMP";
                if (Directory.Exists(sourceFolder))
                {
                    Directory.Delete(sourceFolder, true);
                    Directory.CreateDirectory(sourceFolder);
                }
                System.IO.Compression.ZipFile.ExtractToDirectory(file, sourceFolder, true);

                StopAll();
                xlApp.EnableEvents = false;
                Excel.Application xlApp1;
                Excel.Workbook wb1;
                Excel.Worksheet wsCh1 = null, wsDb1 = null;
                Excel.Range destinationRange;
            
                wb1 = xlApp.Workbooks.Open(Filename: sourceFolder + @"\" + selectedMonth + @".xlsm");
                wb1.Activate();
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

                xlApp.DisplayAlerts = false;
                wsCh.Delete();
                wsDb.Delete();
                xlApp.DisplayAlerts = true;
                wsCh = null;
                wsDb = null;

                wsCh1.Copy(Type.Missing, wb.Worksheets[1]);
                wsDb1.Copy(Type.Missing, wb.Worksheets[1]);
                
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
                }

                wsCh.Cells.Replace( What: "[" + selectedMonth + ".xlsm]", Replacement: "", LookAt: XlLookAt.xlPart, SearchOrder: XlSearchOrder.xlByRows, MatchCase: false, SearchFormat: false, ReplaceFormat: false);
                wsCh.Activate();
                ResumeAll();

                GlobalMethods.ReleseObject(wsCh1);
                GlobalMethods.ReleseObject(wsDb1);
                wb1.Saved = true;
                wb1.Close();
                GlobalMethods.ReleseObject(wb1);

                SaveLoader.LoadAsyncFromFolder("TEMP");
                //references.UpdateAllWSs();
                Directory.Delete(sourceFolder, true);
                InitExcelEvents();
            }
            GlobalMethods.ToLog("Открыты счетчики за " + selectedMonth + " " + selectedYear);
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

            Events_BeforeDoubleClick = new Excel.DocEvents_BeforeDoubleClickEventHandler(Application_BeforeDoubleClick);
            wsCh.BeforeDoubleClick += Events_BeforeDoubleClick;

            Events_Change = new Excel.DocEvents_ChangeEventHandler(Application_Change);
            wsCh.Change += Events_Change;
            //wsDb.Change += Events_Change;

            Events_SelectionChange = new Excel.DocEvents_SelectionChangeEventHandler(Application_SelectionChange);
            wsCh.SelectionChange += Events_SelectionChange;

            Events_SheetSelectionChange = new Excel.WorkbookEvents_SheetSelectionChangeEventHandler(Application_SelectionChange);
            wb.SheetSelectionChange += Events_SheetSelectionChange;

            Events_SheetChange = new Excel.WorkbookEvents_SheetChangeEventHandler(Application_Change);
            wb.SheetChange += Events_SheetChange;

            Events_BeforeSave = new Excel.WorkbookEvents_BeforeSaveEventHandler(Wb_BeforeSave);
            wb.BeforeSave += Events_BeforeSave;

            GlobalMethods.CalculateFormsPositions();
        }
        private void ClearEvents()
        {
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
            GlobalMethods.ToLog("Книга сохранена");
            SaveLoader.SaveAsync();
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
            GlobalMethods.ToLog("Активирован лист " + ((Excel.Worksheet)sh).Name);
            menu.ActivateSheet(sh);
        }
        private void Application_BeforeDoubleClick(Excel.Range range, ref bool cancel)
        {
            GlobalMethods.ToLog("Двойной клик на ячейке " + range.Address);
            menu.DblClick(range);
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
        }
        private void Application_Change(Excel.Range range)
        {
            if (range.Cells.Count > 1)
            {
                ChagedRange(range);
            }
            else
            {
                Changed(range);
            }
            menu.CellValueChanged(range);
        }

        private void Application_Change(object sh, Excel.Range range)
        {
            if (((Excel.Worksheet)sh).CodeName != "PS")
            {
                if (range.Cells.Count > 1)
                {
                    ChagedRange(range);
                }
                else
                {
                    Changed(range);
                }
            }
        }

        private void ChagedRange(Excel.Range rng)
        {
            object[,] newValsArray = (object[,])rng.Formula;
            for (int i = 1; i <= rng.Columns.Count; i++)
            {
                for (int j = 1; j <= rng.Rows.Count; j++)
                {
                    GlobalMethods.ToLog("Изменено значение ячейки " + ((Excel.Range)rng.Cells[j,i]).Address + " с '" + oldValsArray[j,i] + "' на '" + newValsArray[j,i] + "'");
                }
            }
        }

        private void Changed(Excel.Range rng)
        {
            GlobalMethods.ToLog("Изменено значение ячейки " + rng.Address + " с '" + oldVal + "' на '" + rng.Value + "'");
        }
        
        public void RestoreExcel()
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

        public void ArhivateNew(string year, string month)
        {
            wb.Save();
            string sourceFolder = dir + @"\current";
            string tempDirectory = Path.Combine(Path.GetTempPath(), DateTime.Today.ToString("MMMM", new CultureInfo("ru-RU")));
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
    
        public void ReleaseAllComObjects()
        {
            references.ReleaseAllComObjects();
            heads.ReleaseAllComObjects();
        }
    }
}
