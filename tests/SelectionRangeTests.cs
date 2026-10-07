using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using SheetPace;
class SelectionRangeTests
{
    delegate bool EnumProc(IntPtr hwnd, IntPtr unused);
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X,Y; }
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT point);
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L,T,R,B; }
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr unused);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint process);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd,out RECT rect);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd,System.Text.StringBuilder text,int length);
    [DllImport("user32.dll")] static extern int GetWindowRgn(IntPtr hwnd,IntPtr region);
    [DllImport("gdi32.dll")] static extern IntPtr CreateRectRgn(int l,int t,int r,int b);
    [DllImport("gdi32.dll")] static extern bool PtInRegion(IntPtr region,int x,int y);
    [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr value);
    [DllImport("user32.dll")] static extern bool GetLayeredWindowAttributes(IntPtr hwnd,out uint color,out byte alpha,out uint flags);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags,uint x,uint y,uint data,UIntPtr extra);
    [DllImport("user32.dll")] static extern void keybd_event(byte key,byte scan,uint flags,UIntPtr extra);
    static int checks;static uint visualPid;static string output;
    static void Check(bool value,string name) { checks++;if(!value)throw new Exception(name); }
    static void Release(object value) { if(value!=null && Marshal.IsComObject(value))Marshal.ReleaseComObject(value); }
    sealed class Coverage : IDisposable
    {
        internal Rectangle Bounds;internal IntPtr Region;internal byte Alpha;
        internal bool Contains(Point p) { return Bounds.Contains(p) && (Region==IntPtr.Zero || PtInRegion(Region,p.X-Bounds.Left,p.Y-Bounds.Top)); }
        public void Dispose() { if(Region!=IntPtr.Zero)DeleteObject(Region); }
    }
    static List<Coverage> Windows()
    {
        List<Coverage> result=new List<Coverage>();
        EnumWindows(delegate(IntPtr hwnd,IntPtr unused) {
            uint owner;GetWindowThreadProcessId(hwnd,out owner);if(owner!=visualPid || !IsWindowVisible(hwnd))return true;
            System.Text.StringBuilder title=new System.Text.StringBuilder(128);GetWindowText(hwnd,title,128);
            if(!title.ToString().StartsWith("SheetPace Hover Overlay") && title.ToString()!="SheetPace Multi Selection Overlay")return true;
            RECT rect;GetWindowRect(hwnd,out rect);IntPtr region=CreateRectRgn(0,0,0,0);if(GetWindowRgn(hwnd,region)==0){DeleteObject(region);region=IntPtr.Zero;}
            byte alpha;uint color,flags;GetLayeredWindowAttributes(hwnd,out color,out alpha,out flags);
            result.Add(new Coverage {Bounds=Rectangle.FromLTRB(rect.L,rect.T,rect.R,rect.B),Region=region,Alpha=alpha});return true;
        },IntPtr.Zero);return result;
    }
    static bool Includes(int value,int[] pairs) { for(int i=0;i<pairs.Length;i+=2)if(value>=pairs[i] && value<=pairs[i+1])return true;return false; }
    static bool Contains(Point p,Rectangle[] rectangles) { foreach(Rectangle rect in rectangles)if(rect.Contains(p))return true;return false; }
    static bool Validate(dynamic window,Rectangle grid,HoverFrame frame,int[] rows,int[] columns,bool assert)
    {
        List<Coverage> overlays=Windows();bool okay=true;int valid=0,lit=0;int alpha=Settings.Load(Settings.DefaultPath).Alpha;
        try {
            for(int y=grid.Top+28;y<grid.Bottom-35;y+=31)for(int x=grid.Left+40;x<grid.Right-35;x+=87) {
                dynamic cell=null;
                try {
                    cell=window.RangeFromPoint(x,y);if(cell==null)continue;
                    Point p=new Point(x,y);int row=(int)cell.Row,col=(int)cell.Column;bool expected=Includes(row,rows)||Includes(col,columns);
                    bool packet=Contains(p,frame.RowBands)||Contains(p,frame.ColumnBands);int layers=0;
                    foreach(Coverage overlay in overlays)if(overlay.Contains(p)){layers++;if(overlay.Alpha!=alpha)okay=false;}
                    valid++;if(expected)lit++;
                    if(packet!=expected || layers!=(expected?1:0)) {
                        okay=false;if(!assert)return false;if(assert)throw new Exception("coverage mismatch cell="+row+","+col+" expected="+expected+" packet="+packet+" layers="+layers+" point="+p+" bands="+frame.RowBands.Length+","+frame.ColumnBands.Length);
                    }
                } finally { Release(cell); }
            }
            return okay && valid>100 && lit>0;
        } finally { foreach(Coverage overlay in overlays)overlay.Dispose(); }
    }
    static HoverFrame Verify(dynamic window,IntPtr main,HoverPacket packet,string label,int[] rows,int[] columns)
    {
        SetForegroundWindow(main);
        if(GetForegroundWindow()!=main) { RECT rect;GetWindowRect(main,out rect);SetCursorPos(rect.R-420,rect.T+24);mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);Thread.Sleep(200); }
        Stopwatch watch=Stopwatch.StartNew();HoverFrame frame=null;bool valid=false;
        Rectangle grid=GridGeometry.FindGrids(main)[0];
        while(watch.ElapsedMilliseconds<5000) {
            frame=packet.Read();
            if(frame!=null && frame.Visible && frame.Mode==HighlightMode.ClickSelection && Validate(window,grid,frame,rows,columns,false)){valid=true;break;}
            Thread.Sleep(30);
        }
        if(frame!=null && !frame.Visible) {
            RECT bounds;GetWindowRect(main,out bounds);System.Text.StringBuilder title=new System.Text.StringBuilder(256);GetWindowText(GetForegroundWindow(),title,256);
            Console.WriteLine("Foreground title="+title+" main bounds="+bounds.L+","+bounds.T+","+bounds.R+","+bounds.B);
            using(Bitmap bitmap=new Bitmap(1920,1080)){using(Graphics graphics=Graphics.FromImage(bitmap))graphics.CopyFromScreen(0,0,0,0,bitmap.Size);bitmap.Save(Path.Combine(output,"failure.png"));}
            Console.WriteLine("Hidden frame foreground="+GetForegroundWindow()+" main="+main+" target="+frame.TargetRow+","+frame.TargetColumn+" grid="+frame.Grid+" bands="+frame.RowBands.Length+","+frame.ColumnBands.Length+" stamp="+frame.Timestamp); }
        Check(frame!=null && frame.Visible,label+" visible");
        Check(valid || Validate(window,grid,frame,rows,columns,true),label+" packet, native region, gaps and single-alpha coverage");
        Console.WriteLine("PASS "+label+" bands="+frame.RowBands.Length+","+frame.ColumnBands.Length+" sample_ms="+(frame.ElapsedMicroseconds/1000.0).ToString("F2"));return frame;
    }
    static Point CellPoint(dynamic window,Rectangle grid,int row,int column)
    {
        int foundY=-1;
        for(int y=grid.Top+25;y<grid.Bottom-30;y+=5) {
            dynamic hit=null;try{hit=window.RangeFromPoint(grid.Right-70,y);if(hit!=null && (int)hit.Row==row){foundY=y;break;}}finally{Release(hit);}
        }
        if(foundY>=0)for(int x=grid.Left+35;x<grid.Right-30;x+=5) {
            dynamic hit=null;try{hit=window.RangeFromPoint(x,foundY);if(hit!=null && (int)hit.Column==column)return new Point(x+1,foundY+1);}finally{Release(hit);}
        }
        throw new Exception("cell not visible "+row+","+column);
    }
    [STAThread] static int Main(string[] args)
    {
        Console.OutputEncoding=System.Text.Encoding.UTF8;SetThreadDpiAwarenessContext(new IntPtr(-4));
        output=Path.GetFullPath(args[0]);Directory.CreateDirectory(output);
        IntPtr previous=GetForegroundWindow();POINT originalPointer;GetCursorPos(out originalPointer);
        dynamic app=null,book=null,sheet=null,window=null,addin=null,bridge=null;HoverPacket packet=null;Settings original=Settings.Load(Settings.DefaultPath);
        try {
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));app.DisplayAlerts=false;
            book=app.Workbooks.Add();sheet=book.Worksheets[1];sheet.Range["A1:AZ150"].ColumnWidth=9.5;sheet.Range["A1:AZ150"].RowHeight=18;
            sheet.Range["C3:E4"].Merge();app.Visible=true;app.WindowState=-4137;window=app.ActiveWindow;
            IntPtr main=new IntPtr((int)app.Hwnd);uint excelPid;GetWindowThreadProcessId(main,out excelPid);
            addin=app.COMAddIns.Item("SheetPace.Connect");addin.Connect=true;bridge=addin.Object;
            visualPid=(uint)(int)bridge.VisualHostProcessId;bridge.OnToggleHover(null,true);bridge.OnSetMode(null,0);
            SetForegroundWindow(main);Thread.Sleep(1300);packet=new HoverPacket(excelPid,false);
            int scenarios=0;
            foreach(int zoom in args.Length>1 && args[1]=="--edges" ? new int[0] : new[]{100,80,150}) {
                window.FreezePanes=false;window.Zoom=zoom;window.ScrollRow=1;window.ScrollColumn=1;Thread.Sleep(250);
                sheet.Range["B7"].Select();HoverFrame first=Verify(window,main,packet,"single "+zoom,new[]{7,7},new[]{2,2});scenarios++;
                sheet.Range["B7:F7"].Select();HoverFrame expanded=Verify(window,main,packet,"horizontal unchanged active cell "+zoom,new[]{7,7},new[]{2,6});
                Check(first.ShapeRevision!=expanded.ShapeRevision,"range expansion invalidates shape");scenarios++;
                sheet.Range["B7:B12"].Select();Verify(window,main,packet,"vertical "+zoom,new[]{7,12},new[]{2,2});scenarios++;
                sheet.Range["B7:D10"].Select();Verify(window,main,packet,"rectangle "+zoom,new[]{7,10},new[]{2,4});scenarios++;
                sheet.Range["B7,D9,F11"].Select();Verify(window,main,packet,"noncontiguous cells "+zoom,new[]{7,7,9,9,11,11},new[]{2,2,4,4,6,6});scenarios++;
                sheet.Range["B7:D7,F9:F12"].Select();Verify(window,main,packet,"noncontiguous areas "+zoom,new[]{7,7,9,12},new[]{2,4,6,6});scenarios++;
                sheet.Range["C3"].Select();Verify(window,main,packet,"merged C3:E4 "+zoom,new[]{3,4},new[]{3,5});scenarios++;
            }
            window.Zoom=100;sheet.Range["C3:E4"].UnMerge();sheet.Rows[8].Hidden=true;sheet.Columns[3].Hidden=true;
            sheet.Range["B7:D10,F12"].Select();Verify(window,main,packet,"hidden row and column",new[]{7,10,12,12},new[]{2,4,6,6});scenarios++;
            sheet.Rows[8].Hidden=false;sheet.Columns[3].Hidden=false;
            sheet.Range["B7:AZ7"].Select();Verify(window,main,packet,"partly offscreen",new[]{7,7},new[]{2,52});scenarios++;
            sheet.Range["7:7,11:11"].Select();Verify(window,main,packet,"entire separated rows",new[]{7,7,11,11},new[]{1,16384});scenarios++;
            sheet.Range["B:B,F:F"].Select();Verify(window,main,packet,"entire separated columns",new[]{1,1048576},new[]{2,2,6,6});scenarios++;
            sheet.Range["C5"].Select();window.FreezePanes=true;window.ScrollRow=12;window.ScrollColumn=6;Thread.Sleep(300);
            sheet.Range["A2:B3,G15:J17"].Select();Verify(window,main,packet,"frozen and scrolled disjoint areas",new[]{2,3,15,17},new[]{1,2,7,10});scenarios++;
            sheet.Range["A2:J17"].Select();Verify(window,main,packet,"continuous range across frozen panes",new[]{2,17},new[]{1,10});scenarios++;
            window.FreezePanes=false;window.ScrollRow=1;window.ScrollColumn=1;sheet.Range["B7"].Select();Thread.Sleep(300);
            Rectangle grid=GridGeometry.FindGrids(main)[0];
            foreach(int[] target in new[]{new[]{9,4},new[]{11,6}}) {
                Point point=CellPoint(window,grid,target[0],target[1]);SetForegroundWindow(main);SetCursorPos(point.X,point.Y);
                keybd_event(0x11,0,0,UIntPtr.Zero);mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);keybd_event(0x11,0,2,UIntPtr.Zero);Thread.Sleep(200);
            }
            Verify(window,main,packet,"actual Ctrl mouse selection",new[]{7,7,9,9,11,11},new[]{2,2,4,4,6,6});scenarios++;
            sheet.Range["B7"].Select();Verify(window,main,packet,"remove extra areas",new[]{7,7},new[]{2,2});scenarios++;
            HoverFrame stable=packet.Read();SetCursorPos(grid.Right-70,grid.Bottom-70);Thread.Sleep(200);HoverFrame moved=packet.Read();
            Check(moved.Visible && moved.ShapeRevision==stable.ShapeRevision,"pointer motion retains selection geometry");
            string report="PASS "+checks+" checks; "+scenarios+" real Excel selection scenarios; all lit points have exactly one alpha layer";
            File.WriteAllText(Path.Combine(output,"report.txt"),report);Console.WriteLine(report);return 0;
        } catch(Exception ex) {
            Console.Error.WriteLine(ex);File.WriteAllText(Path.Combine(output,"error.txt"),ex.ToString());return 1;
        } finally {
            keybd_event(0x11,0,2,UIntPtr.Zero);
            try { if(bridge!=null){bridge.OnToggleHover(null,original.HoverEnabled);bridge.OnSetMode(null,(int)original.Mode);Thread.Sleep(300);} } catch{}
            if(packet!=null)packet.Dispose();try{if(book!=null)book.Close(false);}catch{}try{if(app!=null)app.Quit();}catch{}
            foreach(object value in new object[]{(object)bridge,(object)addin,(object)window,(object)sheet,(object)book,(object)app})Release(value);
            GC.Collect();GC.WaitForPendingFinalizers();
            SetCursorPos(originalPointer.X,originalPointer.Y);if(previous!=IntPtr.Zero)SetForegroundWindow(previous);
        }
    }
}
