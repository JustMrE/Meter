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

namespace Meter
{
    class MyApplicationContext : ApplicationContext
    {
        [DllImport( "user32.dll" )]
        private static extern int ShowWindow( IntPtr hWnd, uint Msg );

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, ref Rect rectangle);

        private const uint SW_RESTORE = 0x09;

        public static MyApplicationContext instance;

        // public static List<Form> forms;
        public static List<Form> menues;

        Excel.OLEObject inkPicture; 
        
        public Excel.Application xlApp;
        public Excel.Workbook wb;
        public Excel.Worksheet wsCh, wsDb;
        public bool closed, excelClosed;
        public IntPtr xlAppHwnd;
        public Rect xlAppRect;
        // public MenuBase menu;
        public NewMenuBase menu;
        //public NewMenuBase newMenu; 
        public ColorsData colors;
        public double zoom;
        public RangeReferences references;
        public Formula formulas;
        public static bool save;
        public static string dir;
        public static List<int> menuIndexes = new List<int>();

        string file;

        public Excel.WorkbookEvents_BeforeCloseEventHandler Event_BeforeClose;
        public Excel.WorkbookEvents_WindowResizeEventHandler Event_WindowResize;
        public Excel.DocEvents_BeforeRightClickEventHandler Events_BeforeRightClick;
        public Excel.DocEvents_DeactivateEventHandler Events_DeactivateSheet;
        public Excel.WorkbookEvents_SheetActivateEventHandler Events_ActivateSheet;
        public Excel.DocEvents_BeforeDoubleClickEventHandler Events_BeforeDoubleClick;
        public Excel.DocEvents_ChangeEventHandler Events_Change;
        public Excel.DocEvents_SelectionChangeEventHandler Events_SelectionChange;
        public Excel.WorkbookEvents_BeforeSaveEventHandler Events_BeforeSave;

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
                    xlApp.DisplayAlerts = false;
                }

                references.ReleaseAllComObjects();

                ClearEvents();

                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wsCh);
                Marshal.ReleaseComObject(wsDb);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //Thread.Sleep(10000);
                ExitThread();
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
            dir = Process.GetCurrentProcess().MainModule.FileName;
            file = System.IO.Path.GetDirectoryName(dir) + @"\Счетчики.xlsm";
            if (!File.Exists(file))
            {
                dir = System.Reflection.Assembly.GetExecutingAssembly().Location;
                file = System.IO.Path.GetDirectoryName(dir) + @"\Счетчики.xlsm";
            }

            InitExcel();
            InitForms();
            InitExcelEvents();
            menu.ClearContextMenu();

            // InitMyContextMenu();
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
            
            //ColorsData.RefreshColors();

            zoom = (double)xlApp.ActiveWindow.Zoom;
        }
        private void InitForms()
        {
            // forms = new List<Form>()
            // {
            //     new Menu(),
            //     new AdminMenu()
            // };

            // foreach (var form in forms)
            // {
            //     form.FormClosed += onFormClosed;
            // }
            // forms[0].Show();

            menues = new List<Form>();
            menues.Add(new NewMenu());
            menues.Add(new NewMenuAdmin());

            foreach (NewMenuBase form in menues)
            {
                form.FormClosed += onFormClosed;
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
            // var watch = Stopwatch.StartNew();
            // if (MenuBase.test == "async")
            // {
                    SaveLoader.SaveAsync();
            // }
            // else if (MenuBase.test == "sync")
            // {
            //     SaveLoader.Save();
            // }
            // watch.Stop();
            // MessageBox.Show((watch.ElapsedMilliseconds / 1000) + "sec");
        }

        public void Application_BeforeClose(ref bool cancel)
        {
            if (!closed)
            {
                closed = true;
                // foreach (MenuBase form in MyApplicationContext.forms)
                // {
                //     form.FormClose();
                // }
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
            menu.ActivateSheet(sh);
        }
        private void Application_BeforeDoubleClick(Excel.Range range, ref bool cancel)
        {
            menu.DblClick(range);
        }
        private void Application_SelectionChange(Excel.Range range)
        {
            menu.SlectionChanged(range);
        }
        private void Application_Change(Excel.Range range)
        {
            menu.CellValueChanged(range);
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
    }
}
