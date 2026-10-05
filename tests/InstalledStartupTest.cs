using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Microsoft.Win32;
class InstalledStartupTest
{
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L,T,R,B; }
    delegate bool EnumProc(IntPtr window,IntPtr unused);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr unused);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent,EnumProc callback,IntPtr unused);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window,out uint pid);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window,StringBuilder name,int length);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr window,StringBuilder name,int length);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd,uint msg,IntPtr wparam,IntPtr lparam);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags,uint dx,uint dy,uint data,UIntPtr extra);
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT point);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr window,out RECT rect);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr window,uint id,ref Guid iid,[MarshalAs(UnmanagedType.Interface)] out object result);
    static string ClassName(IntPtr window) { StringBuilder text=new StringBuilder(256);GetClassName(window,text,256);return text.ToString(); }
    static string Title(IntPtr window) { StringBuilder text=new StringBuilder(256);GetWindowText(window,text,256);return text.ToString(); }
    static string Hash(string path) { using(SHA256 hash=SHA256.Create())using(Stream file=File.OpenRead(path))return BitConverter.ToString(hash.ComputeHash(file)); }
    static void Check(bool value,string name) { if(!value)throw new Exception(name);Console.WriteLine("PASS "+name); }
    static IntPtr Window(uint pid,string name,bool byTitle)
    {
        IntPtr result=IntPtr.Zero;
        EnumWindows(delegate(IntPtr window,IntPtr unused) { uint owner;GetWindowThreadProcessId(window,out owner);if(owner==pid && IsWindowVisible(window) && (byTitle?Title(window):ClassName(window))==name){result=window;return false;}return true; },IntPtr.Zero);
        return result;
    }
    static POINT CellPoint(dynamic app,IntPtr main,int row,int column,bool arrow)
    {
        IntPtr grid=IntPtr.Zero;EnumChildWindows(main,delegate(IntPtr window,IntPtr unused){if(ClassName(window)=="EXCEL7"){grid=window;return false;}return true;},IntPtr.Zero);
        RECT rect;GetWindowRect(grid,out rect);dynamic view=app.ActiveWindow;
        int yTop=-1,yBottom=-1,xLeft=-1,xRight=-1;
        for(int y=rect.T+2;y<rect.B;y+=3)
        {
            dynamic cell=view.RangeFromPoint(rect.L+80,y);if(cell==null)continue;int value=(int)cell.Row;Marshal.ReleaseComObject(cell);
            if(value==row && yTop<0)yTop=y;
            if(value>row && yTop>=0){yBottom=y;break;}
        }
        if(yTop<0 || yBottom<0)throw new Exception("Cannot hit-test target row");
        int middle=(yTop+yBottom)/2;
        for(int x=rect.L+2;x<rect.R;x+=3)
        {
            dynamic cell=view.RangeFromPoint(x,middle);if(cell==null)continue;int value=(int)cell.Column;Marshal.ReleaseComObject(cell);
            if(value==column && xLeft<0)xLeft=x;
            if(value>column && xLeft>=0){xRight=x;break;}
        }
        Marshal.ReleaseComObject(view);
        if(xLeft<0 || xRight<0)throw new Exception("Cannot hit-test target column");
        return new POINT {X=arrow?xRight-7:(xLeft+xRight)/2,Y=middle};
    }
    static bool CrossAt(uint pid,POINT point)
    {
        for(int attempt=0;attempt<80;attempt++)
        {
            IntPtr row=Window(pid,"SheetPace Hover Overlay",true),column=Window(pid,"SheetPace Hover Overlay Column Above",true);RECT r,c;
            if(row!=IntPtr.Zero && column!=IntPtr.Zero && GetWindowRect(row,out r) && GetWindowRect(column,out c) && point.Y>=r.T && point.Y<r.B && point.X>=c.L && point.X<c.R)return true;
            Thread.Sleep(50);
        }
        return false;
    }
    static void ClickControl(IntPtr parent,string title)
    {
        IntPtr target=IntPtr.Zero;
        EnumChildWindows(parent,delegate(IntPtr child,IntPtr unused){if(Title(child)==title){target=child;return false;}return true;},IntPtr.Zero);
        Check(target!=IntPtr.Zero,"settings control exists: "+title);PostMessage(target,0xF5,IntPtr.Zero,IntPtr.Zero);Thread.Sleep(200);
    }
    static void ShowRibbon(IntPtr main)
    {
        var root=System.Windows.Automation.AutomationElement.FromHandle(main);
        var tab=root.FindFirst(System.Windows.Automation.TreeScope.Descendants,new System.Windows.Automation.AndCondition(
            new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.NameProperty,"SheetPace"),
            new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.ControlTypeProperty,System.Windows.Automation.ControlType.TabItem)));
        Check(tab!=null,"SheetPace Ribbon tab exists");
        object pattern;
        if(tab.TryGetCurrentPattern(System.Windows.Automation.SelectionItemPattern.Pattern,out pattern))((System.Windows.Automation.SelectionItemPattern)pattern).Select();
        else if(tab.TryGetCurrentPattern(System.Windows.Automation.InvokePattern.Pattern,out pattern))((System.Windows.Automation.InvokePattern)pattern).Invoke();
        else throw new Exception("Cannot activate SheetPace tab");
        Thread.Sleep(300);
        Check(root.FindFirst(System.Windows.Automation.TreeScope.Descendants,new System.Windows.Automation.PropertyCondition(System.Windows.Automation.AutomationElement.NameProperty,"光影设置"))!=null,"renamed settings button appears on Ribbon");
    }
    static IntPtr WaitWindow(uint pid,string title)
    {
        for(int i=0;i<80;i++){IntPtr window=Window(pid,title,true);if(window!=IntPtr.Zero)return window;Thread.Sleep(100);}
        return IntPtr.Zero;
    }
    static void Screenshot(IntPtr window,string path)
    {
        RECT rect;GetWindowRect(window,out rect);using(Bitmap image=new Bitmap(rect.R-rect.L,rect.B-rect.T))
        {using(Graphics graphics=Graphics.FromImage(image))graphics.CopyFromScreen(rect.L,rect.T,0,0,image.Size);image.Save(path,ImageFormat.Png);}
    }
    static string[] GetModules(Process process)
    {
        System.Collections.Generic.List<string> lines=new System.Collections.Generic.List<string>();
        foreach(ProcessModule module in process.Modules)lines.Add(module.ModuleName+" base="+module.BaseAddress.ToInt64().ToString("X")+" size="+module.ModuleMemorySize.ToString("X"));
        return lines.ToArray();
    }
    [STAThread] static int Main(string[] args)
    {
        Console.OutputEncoding=Encoding.UTF8;
        SetThreadDpiAwarenessContext(new IntPtr(-4));
        string source=Path.GetFullPath(args[0]),output=Path.GetFullPath(args[1]);Directory.CreateDirectory(output);
        string originalHash=Hash(source),copy=Path.Combine(output,"Startup-TestCopy.xlsx");File.Copy(source,copy,true);
        File.WriteAllText(Path.Combine(output,"started.txt"),DateTime.Now.ToString("o"));
        IntPtr previous=GetForegroundWindow();POINT pointer;GetCursorPos(out pointer);
        dynamic app=null,book=null,sheet=null,addin=null,native=null;Process process=null;bool closed=false,restoreHover=false,originalHover=false;int originalMode=0;
        try
        {
            string exe,keyPath=String.Join(((char)92).ToString(),new[]{"Software","Microsoft","Windows","CurrentVersion","App Paths","excel.exe"});
            using(RegistryKey key=Registry.LocalMachine.OpenSubKey(keyPath))exe=Convert.ToString(key.GetValue(""));
            process=Process.Start(new ProcessStartInfo(exe,"/x /r "+(char)34+copy+(char)34){UseShellExecute=false});
            uint pid=(uint)process.Id;Console.WriteLine("STAGE actual Excel file launch PID="+pid);File.WriteAllText(Path.Combine(output,"owned-excel-pid.txt"),pid.ToString());
            IntPtr main=IntPtr.Zero;
            for(int i=0;i<100;i++)
            {
                if(process.HasExited)throw new Exception("Excel exited during startup: "+process.ExitCode);
                main=Window(pid,"XLMAIN",false);
                if(main!=IntPtr.Zero)
                {
                    IntPtr grid=IntPtr.Zero;EnumChildWindows(main,delegate(IntPtr window,IntPtr unused){if(ClassName(window)=="EXCEL7"){grid=window;return false;}return true;},IntPtr.Zero);
                    if(grid!=IntPtr.Zero)
                    {
                        Guid iid=new Guid("00020400-0000-0000-C000-000000000046");object result;
                        if(AccessibleObjectFromWindow(grid,0xFFFFFFF0,ref iid,out result)==0){native=result;app=native.Application;break;}
                    }
                }
                Thread.Sleep(200);
            }
            Check(app!=null,"attach only to Excel instance launched with workbook copy");
            bool disabled=args.Length>2 && args[2]=="--disabled";
            addin=app.COMAddIns.Item("SheetPace.Connect");Check((bool)addin.Connect != disabled,disabled?"control: SheetPace disabled":"Excel loaded registered COM add-in automatically");
            book=app.ActiveWorkbook;Check(book!=null && String.Equals((string)book.FullName,copy,StringComparison.OrdinalIgnoreCase),"requested workbook copy opened");
            Check((bool)book.ReadOnly,"workbook copy opened read-only");
            sheet=book.Worksheets[1];sheet.Activate();app.Visible=true;SetForegroundWindow(main);Thread.Sleep(1200);
            if(!disabled && (args.Length<3 || args[2]!="--minimal"))
            {
            dynamic bridge=addin.Object;
            uint visualPid=(uint)(int)bridge.VisualHostProcessId;
            originalHover=(bool)bridge.GetHoverPressed(null);originalMode=(int)bridge.GetHighlightMode(null);bridge.OnToggleHover(null,true);bridge.OnSetMode(null,0);restoreHover=true;
            Marshal.ReleaseComObject(bridge);Thread.Sleep(1200);
            Check(visualPid>0 && visualPid!=pid,"visuals run in separate process");
            string diagnostics=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"SheetPace","diagnostics.log");
            Check(File.ReadAllText(diagnostics).Contains("PID="+pid+" Ribbon GetCustomUI"),"Excel requested SheetPace Ribbon XML");
            ShowRibbon(main);Screenshot(main,Path.Combine(output,"ribbon.png"));
            POINT hoverPoint=CellPoint(app,main,8,7,false);SetForegroundWindow(main);SetCursorPos(hoverPoint.X,hoverPoint.Y);
            mouse_event(2,0,0,0,UIntPtr.Zero);Thread.Sleep(40);mouse_event(4,0,0,0,UIntPtr.Zero);
            Check(CrossAt(visualPid,hoverPoint),"selected cross aligns with physically clicked cell");
            POINT away=CellPoint(app,main,12,9,false);SetCursorPos(away.X,away.Y);Thread.Sleep(300);
            Check(CrossAt(visualPid,hoverPoint),"moving mouse preserves selected cross");
            Screenshot(main,Path.Combine(output,"hover.png"));
            dynamic enabledCheck=addin.Object;enabledCheck.OnToggleHover(null,false);Thread.Sleep(500);
            Check(Window(visualPid,"SheetPace Hover Overlay",true)==IntPtr.Zero,"Ribbon switch disables cross immediately");
            enabledCheck.OnToggleHover(null,true);Marshal.ReleaseComObject(enabledCheck);Thread.Sleep(1200);
            Check(CrossAt(visualPid,hoverPoint),"Ribbon switch restores selected cross");
            if(args.Length>2 && args[2]=="--dialogs")
            {
                dynamic controller=addin.Object;controller.OnSettings(null);
                IntPtr settingsWindow=WaitWindow(visualPid,"SheetPace · 光影设置");Check(settingsWindow!=IntPtr.Zero,"Ribbon settings callback opens visual host dialog");
                Thread.Sleep(350);Screenshot(settingsWindow,Path.Combine(output,"settings.png"));
                ClickControl(settingsWindow,"鼠标跟随");ClickControl(settingsWindow,"保存设置");Thread.Sleep(1200);
                Check((int)controller.GetHighlightMode(null)==1,"settings saves mouse-follow mode");SetForegroundWindow(main);SetCursorPos(away.X,away.Y);
                Check(CrossAt(visualPid,away),"mouse-follow mode tracks pointer without clicking");Screenshot(main,Path.Combine(output,"follow.png"));
                controller.OnSettings(null);settingsWindow=WaitWindow(visualPid,"SheetPace · 光影设置");
                ClickControl(settingsWindow,"点击选中（默认）");ClickControl(settingsWindow,"保存设置");Thread.Sleep(1200);
                Check((int)controller.GetHighlightMode(null)==0,"settings saves click-selection mode");SetForegroundWindow(main);
                Check(CrossAt(visualPid,hoverPoint),"switching back restores selected cell cross");
                controller.OnSettings(null);settingsWindow=WaitWindow(visualPid,"SheetPace · 光影设置");ClickControl(settingsWindow,"鼠标跟随");
                PostMessage(settingsWindow,0x10,IntPtr.Zero,IntPtr.Zero);Thread.Sleep(500);
                Check((int)controller.GetHighlightMode(null)==0,"cancelled mode change leaves saved selection mode");
                dynamic valueCell=sheet.Range["F6"];valueCell.Select();Marshal.ReleaseComObject(valueCell);controller.OnShowFilter(null);
                IntPtr filterWindow=WaitWindow(visualPid,"SheetPace · 计数筛选");Check(filterWindow!=IntPtr.Zero,"Ribbon filter callback opens visual host dialog");
                PostMessage(filterWindow,0x10,IntPtr.Zero,IntPtr.Zero);Thread.Sleep(500);Marshal.ReleaseComObject(controller);SetForegroundWindow(main);
            }
            if(args.Length<3 || args[2]!="--hover-only")
            {
            dynamic header=sheet.Range["F5"];SetForegroundWindow(main);Thread.Sleep(300);
            Check(GetForegroundWindow()==main,"owned Excel is foreground before opening filter");
            if(args.Length>2 && args[2]=="--mouse")
            {
                dynamic other=sheet.Range["G8"];other.Select();Marshal.ReleaseComObject(other);
                POINT arrow=CellPoint(app,main,5,6,true);
                Console.WriteLine("CLICK physical="+arrow.X+","+arrow.Y);SetCursorPos(arrow.X,arrow.Y);Thread.Sleep(450);
                mouse_event(2,0,0,0,UIntPtr.Zero);Thread.Sleep(40);mouse_event(4,0,0,0,UIntPtr.Zero);
            }
            else { header.Select();app.SendKeys("%{DOWN}",false); }
            Marshal.ReleaseComObject(header);Thread.Sleep(2500);
            Screenshot(main,Path.Combine(output,"native-filter-before-check.png"));
            Check(WaitWindow(visualPid,"SheetPace Count Overlay")!=IntPtr.Zero,"native filter count overlay rendered by visual host");
            Screenshot(main,Path.Combine(output,"native-filter.png"));
            string expectedPath=Path.Combine(output,"native-expected.tsv"),diagnosticPath=Environment.GetEnvironmentVariable("SHEETPACE_NATIVE_DIAGNOSTICS");
            if(File.Exists(expectedPath))
            {
                Check(!String.IsNullOrEmpty(diagnosticPath) && File.Exists(diagnosticPath),"native menu diagnostics captured");
                string rows=File.ReadAllText(diagnosticPath);
                foreach(string line in File.ReadAllLines(expectedPath))
                {
                    string[] fields=line.Split((char)9);
                    Check(rows.Contains("ROW "+fields[0]+" found=True count="+fields[1]+" bounds="),"native count "+fields[0]+"="+fields[1]);
                }
            }
            app.SendKeys("{ESC}",false);Thread.Sleep(300);
            }
            if(args.Length<3 || args[2]!="--no-reconnect")
            {
            addin.Connect=false;Thread.Sleep(400);Check(!(bool)addin.Connect,"host COM disconnection succeeds");
            addin.Connect=true;Thread.Sleep(1000);Check((bool)addin.Connect,"host COM reconnection succeeds");dynamic modeCheck=addin.Object;Check((int)modeCheck.GetHighlightMode(null)==0,"click-selection mode persists across add-in reconnect");Marshal.ReleaseComObject(modeCheck);
            }
            }
            if(restoreHover){dynamic toggle=addin.Object;toggle.OnToggleHover(null,originalHover);toggle.OnSetMode(null,originalMode);Marshal.ReleaseComObject(toggle);Thread.Sleep(200);restoreHover=false;}
            GC.Collect();GC.WaitForPendingFinalizers();
            File.WriteAllLines(Path.Combine(output,"excel-modules-before-quit.txt"),GetModules(process));
            book.Close(false);app.Quit();closed=true;
            foreach(object value in new object[]{(object)addin,(object)sheet,(object)book,(object)native,(object)app})if(value!=null && Marshal.IsComObject(value))Marshal.FinalReleaseComObject(value);
            app=null;book=null;sheet=null;addin=null;native=null;GC.Collect();GC.WaitForPendingFinalizers();
            Check(process.WaitForExit(30000),"owned Excel exits after Quit");Check(process.ExitCode==0,"Excel process exits normally (exit="+process.ExitCode+")");
            Check(Hash(source)==originalHash,"original workbook SHA256 unchanged");
            Console.WriteLine(disabled?"PASS disabled control startup and shutdown":"PASS installed startup, Ribbon, hover, native counts, reconnect and shutdown");return 0;
        }
        catch(Exception ex){Console.Error.WriteLine(ex);File.WriteAllText(Path.Combine(output,"error.txt"),ex.ToString());return 1;}
        finally
        {
            if(restoreHover && addin!=null){try{dynamic toggle=addin.Object;toggle.OnToggleHover(null,originalHover);toggle.OnSetMode(null,originalMode);Marshal.ReleaseComObject(toggle);Thread.Sleep(200);}catch{}}
            if(!closed){try{if(book!=null)book.Close(false);}catch{}try{if(app!=null)app.Quit();}catch{}}
            SetCursorPos(pointer.X,pointer.Y);if(previous!=IntPtr.Zero)SetForegroundWindow(previous);
        }
    }
}
