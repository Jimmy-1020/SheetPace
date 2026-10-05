using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using SheetPace;
class UserWorkbookTest
{
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint process);
    static string Hash(string path) { using(SHA256 hash=SHA256.Create()) using(Stream stream=File.OpenRead(path)) return BitConverter.ToString(hash.ComputeHash(stream)); }
    [STAThread] static int Main(string[] args)
    {
        dynamic app=null,book=null; List<string> report=new List<string>();
        string source=Path.GetFullPath(args[0]),output=Path.GetFullPath(args[1]); Directory.CreateDirectory(output);
        string originalHash=Hash(source),copy=Path.Combine(output,"Workbook-TestCopy.xlsx");
        try
        {
            File.Copy(source,copy,true);
            string[] expected=File.ReadAllLines(Path.Combine(output,"expected.tsv"));
            app=Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")); app.Visible=false; app.DisplayAlerts=false;
            uint process;GetWindowThreadProcessId(new IntPtr((int)app.Hwnd),out process);File.WriteAllText(Path.Combine(output,"owned-excel-pid.txt"),process.ToString());
            book=app.Workbooks.Open(copy,0,false); int checks=0;
            for(int index=1;index<=(int)book.Worksheets.Count;index++)
            {
                dynamic sheet=book.Worksheets[index]; sheet.Activate();
                Dictionary<string,int> counts=new Dictionary<string,int>();int rows=0;
                foreach(string line in expected)
                {
                    string[] parts=line.Split((char)9);if(Int32.Parse(parts[0])!=index)continue;
                    rows=Int32.Parse(parts[1]);counts[parts[2]]=Int32.Parse(parts[3]);
                }
                if(!(bool)sheet.AutoFilterMode)
                {
                    dynamic first=sheet.Cells[5,1],last=sheet.Cells[5+rows,(int)sheet.UsedRange.Columns.Count];
                    sheet.Range[first,last].AutoFilter(); Marshal.ReleaseComObject(first);Marshal.ReleaseComObject(last);
                }
                if((bool)sheet.FilterMode)sheet.ShowAllData();
                using(ExcelContext context=ExcelContext.Capture((object)app,5,6))
                {
                    CountSnapshot snapshot=context.ReadSnapshot();
                    if(snapshot.Total!=rows)throw new Exception("Wrong total: "+sheet.Name+" actual="+snapshot.Total+" expected="+rows);checks++;
                    foreach(KeyValuePair<string,int> pair in counts)
                    {
                        int actual;if(!snapshot.TryGetCount(pair.Key,null,null,out actual)||actual!=pair.Value)throw new Exception("Wrong frequency: "+sheet.Name+" "+pair.Key);checks++;
                        context.Apply(new HashSet<string>{pair.Key});int visible=0;
                        for(int row=6;row<=5+rows;row++)
                        {
                            dynamic record=sheet.Rows[row];bool hidden=(bool)record.Hidden;Marshal.ReleaseComObject(record);
                            if(!hidden)visible++;
                        }
                        if(visible!=pair.Value)throw new Exception("Wrong filter result: "+sheet.Name+" "+pair.Key+" actual="+visible+" expected="+pair.Value);checks++;
                        report.Add("PASS sheet="+sheet.Name+" value="+pair.Key+" count="+actual+" filtered_rows="+visible);
                    }
                    HashSet<string> all=new HashSet<string>();foreach(CountItem item in snapshot.Items)all.Add(item.Label);context.Apply(all);
                }
                Marshal.ReleaseComObject(sheet);
            }
            book.Save();
            if(Hash(source)!=originalHash)throw new Exception("Source file changed");checks++;
            report.Add("PASS "+checks+" checks; original file SHA256 unchanged; test copy only");
            File.WriteAllLines(Path.Combine(output,"workbook-test-report.txt"),report);Console.WriteLine(report[report.Count-1]);return 0;
        }
        catch(Exception ex){File.WriteAllText(Path.Combine(output,"workbook-test-error.txt"),ex.ToString());Console.Error.WriteLine(ex);return 1;}
        finally
        {
            try{if(book!=null)book.Close(false);}catch{}try{if(app!=null)app.Quit();}catch{}
            foreach(object value in new object[]{(object)book,(object)app})if(value!=null&&Marshal.IsComObject(value))Marshal.FinalReleaseComObject(value);
            GC.Collect();GC.WaitForPendingFinalizers();
        }
    }
}
