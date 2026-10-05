using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using SheetPace;
class UiProbe
{
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L,T,R,B; }
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT point);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    delegate bool EnumProc(IntPtr hwnd,IntPtr unused);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc proc,IntPtr unused);
    static Form Overlay(Connect addin,string name) { return (Form)typeof(Connect).GetField(name,BindingFlags.Instance|BindingFlags.NonPublic).GetValue(addin); }
    static void Screenshot(IntPtr hwnd,string path)
    {
        RECT r; GetWindowRect(hwnd,out r); using(Bitmap b=new Bitmap(r.R-r.L,r.B-r.T))
        { using(Graphics g=Graphics.FromImage(b)) g.CopyFromScreen(r.L,r.T,0,0,b.Size); b.Save(path,ImageFormat.Png); }
    }
    [STAThread] static int Main(string[] args)
    {
        SetThreadDpiAwarenessContext(new IntPtr(-4));
        string output=Path.GetFullPath(args[0]);Directory.CreateDirectory(output);Application.EnableVisualStyles();
        dynamic app=null,book=null;Connect addin=null;int result=1;Array custom=new object[0];
        IntPtr previous=GetForegroundWindow();POINT pointer;GetCursorPos(out pointer);
        try
        {
            Console.WriteLine("STAGE create Excel");app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));app.DisplayAlerts=false;
            uint pid;GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out pid);File.WriteAllText(Path.Combine(output,"owned-excel-pid.txt"),pid.ToString());
            bool real=args.Length>1;
            book=real?app.Workbooks.Open(Path.GetFullPath(args[1]),0,true):app.Workbooks.Add();dynamic sheet=book.Worksheets[1];sheet.Activate();
            if(!real)
            {
                sheet.Range["A1:B8"].Value2=new object[,]{{"评级","编号"},{"优秀",1},{"良好",2},{"良好",3},{"中等",4},{"中等",5},{"中等",6},{null,7}};
                sheet.Range["A1:B8"].AutoFilter(1);sheet.Columns[1].ColumnWidth=20;
            }
            app.Visible=true;IntPtr hwnd=new IntPtr((int)app.Hwnd);SetForegroundWindow(hwnd);
            addin=new Connect();addin.OnConnection((object)app,ConnectMode.External,null,ref custom);
            int step=0;bool hoverOK=false,countOK=false;
            using(ApplicationContext context=new ApplicationContext())
            using(System.Windows.Forms.Timer driver=new System.Windows.Forms.Timer{Interval=500})
            {
                driver.Tick+=delegate
                {
                    try
                    {
                        if(step==0)
                        {
                            dynamic target=sheet.Range[real?"G8":"B4"];
                            int x=app.ActiveWindow.PointsToScreenPixelsX((double)target.Left+20),y=app.ActiveWindow.PointsToScreenPixelsY((double)target.Top+6);
                            SetCursorPos(x,y);Marshal.ReleaseComObject(target);driver.Interval=1100;Console.WriteLine("STAGE hover");
                        }
                        else if(step==1)
                        {
                            hoverOK=IsWindowVisible(Overlay(addin,"hover").Handle);Console.WriteLine("Hover overlay visible="+hoverOK);
                            Screenshot(hwnd,Path.Combine(output,"hover.png"));sheet.Range[real?"F5":"A1"].Select();SetForegroundWindow(hwnd);
                            if(GetForegroundWindow()!=hwnd)throw new Exception("Test Excel is not foreground");
                            app.SendKeys("%{DOWN}",false);driver.Interval=2200;Console.WriteLine("STAGE dropdown");
                        }
                        else
                        {
                            countOK=IsWindowVisible(Overlay(addin,"counts").Handle);Console.WriteLine("Native count overlay visible="+countOK);
                            Screenshot(hwnd,Path.Combine(output,"native-filter.png"));
                            object watcher=typeof(Connect).GetField("watcher",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(addin);
                            object menu=watcher.GetType().GetProperty("Latest").GetValue(watcher,null);
                            Console.WriteLine("Native menu detected="+(menu!=null));
                            if(menu!=null)foreach(object row in (System.Collections.IEnumerable)menu.GetType().GetField("Rows").GetValue(menu))Console.WriteLine("ROW "+row.GetType().GetField("Name").GetValue(row));
                            if(GetForegroundWindow()==hwnd)app.SendKeys("{ESC}",false);
                            result=hoverOK&&countOK?0:1;driver.Stop();context.ExitThread();
                        }
                        step++;
                    }
                    catch(Exception ex){Console.Error.WriteLine(ex);driver.Stop();context.ExitThread();}
                };
                driver.Start();Application.Run(context);
            }
            return result;
        }
        catch(Exception ex){Console.Error.WriteLine(ex);return 1;}
        finally
        {
            if(addin!=null)try{addin.OnDisconnection(DisconnectMode.UserClosed,ref custom);}catch{}
            try{if(book!=null)book.Close(false);}catch{}try{if(app!=null)app.Quit();}catch{}
            foreach(object o in new object[]{(object)book,(object)app})if(o!=null&&Marshal.IsComObject(o))Marshal.FinalReleaseComObject(o);
            SetCursorPos(pointer.X,pointer.Y);if(previous!=IntPtr.Zero)SetForegroundWindow(previous);GC.Collect();GC.WaitForPendingFinalizers();
        }
    }
}
