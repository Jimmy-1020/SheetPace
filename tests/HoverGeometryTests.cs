using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using SheetPace;
class HoverGeometryTests
{
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr window);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr window);
    delegate bool EnumProc(IntPtr hwnd,IntPtr unused);
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L,T,R,B; }
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr unused);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint process);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd,out RECT rect);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd,System.Text.StringBuilder text,int length);
    static int checks;
    static object Field(object frame,string name) { return frame.GetType().GetField(name,BindingFlags.Instance|BindingFlags.NonPublic).GetValue(frame); }
    static IntPtr cachedOverlay;
    static IntPtr Overlay(uint pid)
    {
        if(cachedOverlay!=IntPtr.Zero)return cachedOverlay;
        IntPtr result=IntPtr.Zero;
        EnumWindows(delegate(IntPtr hwnd,IntPtr unused){uint owner;GetWindowThreadProcessId(hwnd,out owner);if(owner!=pid)return true;System.Text.StringBuilder title=new System.Text.StringBuilder(128);GetWindowText(hwnd,title,128);if(owner==pid && title.ToString()=="SheetPace Hover Overlay"){result=hwnd;return false;}return true;},IntPtr.Zero);
        cachedOverlay=result;return result;
    }
    static double Median(List<double> values) { values.Sort();return values[values.Count/2]; }
    static void Check(bool value,string name) { if(!value)throw new Exception(name);checks++; }
    static void Release(object value) { if(value!=null && Marshal.IsComObject(value))Marshal.ReleaseComObject(value); }
    static bool Matches(dynamic window,Point point,int row,int column)
    {
        dynamic cell=null,merged=null;
        try { cell=window.RangeFromPoint(point.X,point.Y);if(cell==null)return false;merged=cell.MergeArea;return (int)merged.Row==row && (int)merged.Column==column; }
        finally { Release(merged);Release(cell); }
    }
    [STAThread] static int Main(string[] args)
    {
        Console.OutputEncoding=System.Text.Encoding.UTF8;SetThreadDpiAwarenessContext(new IntPtr(-4));
        string output=Path.GetFullPath(args[0]);Directory.CreateDirectory(output);
        dynamic app=null,book=null,sheet=null,window=null,addin=null;IDisposable packet=null;Settings originalSettings=null;
        bool packetMode=args.Length>1 && args[1]=="--packet";
        try
        {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));app.DisplayAlerts=false;
            addin=app.COMAddIns.Item("SheetPace.Connect");addin.Connect=packetMode;
            book=app.Workbooks.Add();sheet=book.Worksheets[1];sheet.Range["A1:Z100"].ColumnWidth=9.5;sheet.Range["A1:Z100"].RowHeight=18;
            sheet.Range["C3:E4"].Merge();sheet.Range["C3"].Value2="Merged test";
            app.Visible=true;app.WindowState=-4137;window=app.ActiveWindow;
            IntPtr main=new IntPtr((int)app.Hwnd);SetForegroundWindow(main);Thread.Sleep(500);
            Type geometry=typeof(Connect).Assembly.GetType("SheetPace.GridGeometry");
            MethodInfo find=geometry.GetMethod("FindGrid"),slow=geometry.GetMethod("CellRectangle");
            uint visualPid=0;MethodInfo read=null;
            if(packetMode)
            {
                originalSettings=Settings.Load(Settings.DefaultPath);
                dynamic bridge=addin.Object;visualPid=(uint)(int)bridge.VisualHostProcessId;bridge.OnToggleHover(null,true);Release(bridge);Thread.Sleep(300);
                uint excelPid;GetWindowThreadProcessId(main,out excelPid);
                Type type=typeof(Connect).Assembly.GetType("SheetPace.HoverPacket");
                packet=(IDisposable)type.GetConstructor(BindingFlags.Instance|BindingFlags.NonPublic,null,new[]{typeof(uint),typeof(bool)},null).Invoke(new object[]{excelPid,false});
                read=type.GetMethod("Read",BindingFlags.Instance|BindingFlags.NonPublic);
            }
            Rectangle grid=Rectangle.Empty;
            for(int y=280;y<700 && grid.IsEmpty;y+=50)grid=(Rectangle)find.Invoke(null,new object[]{main,new Point(600,y)});
            Check(!grid.IsEmpty,"visible grid");
            List<string> report=new List<string>();
            for(int mode=0;mode<4;mode++)
            {
                window.FreezePanes=false;window.Zoom=mode==1?80:mode==2?150:100;window.ScrollRow=1;window.ScrollColumn=1;
                if(mode==3){sheet.Range["C5"].Select();window.FreezePanes=true;window.ScrollRow=12;window.ScrollColumn=6;}
                Thread.Sleep(200);
                List<double> times=new List<double>(),nativeTimes=new List<double>();int mergedChecks=0;
                for(int i=0;i<24;i++)
                {
                    Point point=i==0?new Point(grid.Left+220,grid.Top+77):new Point(grid.Left+100+(i%6)*95,grid.Top+110+(i/6)*55);
                    dynamic cell=null,merged=null;
                    try
                    {
                        cell=window.RangeFromPoint(point.X,point.Y);if(cell==null)continue;merged=cell.MergeArea;
                        int row=(int)merged.Row,column=(int)merged.Column;
                        if((int)merged.Cells.Count>1)mergedChecks++;
                        if(packetMode)SetForegroundWindow(main);
                        Stopwatch watch=Stopwatch.StartNew();Rectangle rect=Rectangle.Empty;
                        if(packetMode)
                        {
                            SetCursorPos(point.X,point.Y);object frame=null;
                            while(watch.ElapsedMilliseconds<2000)
                            {
                                object current=read.Invoke(packet,null);
                                if(current!=null && (bool)Field(current,"Visible") && (Point)Field(current,"Pointer")==point){frame=current;break;}
                                Thread.Sleep(1);
                            }
                            Check(frame!=null,"probe follows pointer");rect=(Rectangle)Field(frame,"Cell");
                            nativeTimes.Add((int)Field(frame,"ElapsedMicroseconds")/1000.0);
                            Rectangle expected=new Rectangle(grid.Left,rect.Top,grid.Width,rect.Height);
                            bool rendered=false;
                            while(watch.ElapsedMilliseconds<2000)
                            {
                                IntPtr overlay=Overlay(visualPid);RECT bounds;
                                if(overlay!=IntPtr.Zero && IsWindowVisible(overlay) && GetWindowRect(overlay,out bounds) && Rectangle.FromLTRB(bounds.L,bounds.T,bounds.R,bounds.B)==expected){rendered=true;break;}
                                Thread.Sleep(1);
                            }
                            Check(rendered,"strip overlay follows new cell");
                        }
                        else rect=(Rectangle)slow.Invoke(null,new object[]{(object)window,(object)cell,grid,point});
                        watch.Stop();times.Add(watch.Elapsed.TotalMilliseconds);
                        Check(rect.Contains(point),"rectangle contains pointer mode="+mode);
                        Check(grid.Contains(rect),"rectangle inside grid mode="+mode);
                        foreach(Point corner in new[]{new Point(rect.Left+1,rect.Top+1),new Point(rect.Right-2,rect.Bottom-2)})
                            Check(Matches(window,corner,row,column),"corner matches target mode="+mode+" point="+point+" rect="+rect+" row="+row+" col="+column+" corner="+corner);
                        // Move to an adjacent cell and return repeatedly to exercise the cache.
                    }
                    finally { Release(merged);Release(cell); }
                }
                Check(times.Count>=20,"enough geometry samples");times.Sort();
                string line="mode="+mode+" samples="+times.Count+" median_ms="+times[times.Count/2].ToString("F2")+" p95_ms="+times[(int)((times.Count-1)*0.95)].ToString("F2")+" max_ms="+times[times.Count-1].ToString("F2")+" merged="+mergedChecks;
                if(packetMode) { line+=" native_sample_median_ms="+Median(nativeTimes).ToString("F2");Check(times[times.Count/2]<200,"hover median latency below 200 ms"); }
                report.Add(line);Console.WriteLine(line);
            }
            report.Add("PASS "+checks+" geometry checks; packet="+packetMode);File.WriteAllLines(Path.Combine(output,"report.txt"),report);Console.WriteLine(report[report.Count-1]);return 0;
        }
        catch(Exception ex){Console.Error.WriteLine(ex);File.WriteAllText(Path.Combine(output,"error.txt"),ex.ToString());return 1;}
        finally
        {
            if(originalSettings!=null && addin!=null)
            {
                try { dynamic bridge=addin.Object;bridge.OnToggleHover(null,originalSettings.HoverEnabled);Release(bridge);Thread.Sleep(200); } catch { }
            }
            if(packet!=null)packet.Dispose();
            try{if(book!=null)book.Close(false);}catch{}try{if(app!=null)app.Quit();}catch{}
            foreach(object value in new object[]{(object)addin,(object)window,(object)sheet,(object)book,(object)app})Release(value);
            GC.Collect();GC.WaitForPendingFinalizers();
        }
    }
}
