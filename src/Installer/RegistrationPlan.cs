using System;
using System.Collections.Generic;
using Microsoft.Win32;
namespace SheetPaceInstaller
{
    internal sealed class RegistrationEntry
    {
        public string Path, Name; public object Value;
        public RegistryValueKind Kind { get { return Value is int ? RegistryValueKind.DWord : RegistryValueKind.String; } }
        public RegistrationEntry(string path, string name, object value) { Path = path; Name = name; Value = value; }
    }
    internal static class RegistrationPlan
    {
        internal static List<RegistrationEntry> Create(string directory, string prefix)
        {
            List<RegistrationEntry> entries = new List<RegistrationEntry>();
            string classes = prefix + @"Software\Classes\", cls = classes + @"CLSID\" + Engine.Clsid;
            entries.Add(new RegistrationEntry(classes + Engine.ProgId, "", "SheetPace Excel Add-in"));
            entries.Add(new RegistrationEntry(classes + Engine.ProgId + @"\CLSID", "", Engine.Clsid));
            entries.Add(new RegistrationEntry(cls, "", "SheetPace Excel Add-in"));
            entries.Add(new RegistrationEntry(cls + @"\ProgId", "", Engine.ProgId));
            foreach (string suffix in new[] { @"\InprocServer32", @"\InprocServer32\1.0.0.0" })
            {
                string path = cls + suffix;
                if (suffix == @"\InprocServer32") { entries.Add(new RegistrationEntry(path, "", "mscoree.dll")); entries.Add(new RegistrationEntry(path, "ThreadingModel", "Both")); }
                entries.Add(new RegistrationEntry(path, "Class", "SheetPace.Connect"));
                entries.Add(new RegistrationEntry(path, "Assembly", "SheetPace, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"));
                entries.Add(new RegistrationEntry(path, "RuntimeVersion", "v4.0.30319"));
                entries.Add(new RegistrationEntry(path, "CodeBase", new Uri(System.IO.Path.Combine(directory, "SheetPace.dll")).AbsoluteUri));
            }
            string addin = prefix + @"Software\Microsoft\Office\Excel\Addins\" + Engine.ProgId;
            entries.Add(new RegistrationEntry(addin, "FriendlyName", "SheetPace · 筛选计数与行列定位"));
            entries.Add(new RegistrationEntry(addin, "Description", "筛选值数量、鼠标行列光影、可调颜色和透明度"));
            entries.Add(new RegistrationEntry(addin, "LoadBehavior", 3)); entries.Add(new RegistrationEntry(addin, "CommandLineSafe", 0));
            return entries;
        }
        internal static string[] RemovalRoots(string prefix)
        {
            return new[] { prefix + @"Software\Microsoft\Office\Excel\Addins\" + Engine.ProgId, prefix + @"Software\Classes\" + Engine.ProgId, prefix + @"Software\Classes\CLSID\" + Engine.Clsid };
        }
    }
}
