using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
namespace SheetPace
{
    internal sealed class FilterRow { public string Name, Year, Month; public Rectangle Bounds; }
    internal sealed class FilterMenu
    {
        public IntPtr Window;
        public Rectangle Bounds;
        public readonly List<FilterRow> Rows = new List<FilterRow>();
        public long Timestamp;
    }
    internal sealed class NativeFilterWatcher : IDisposable
    {
        private readonly uint process;
        private readonly AutoResetEvent signal = new AutoResetEvent(false);
        private readonly object gate = new object();
        private volatile bool stopping;
        private FilterMenu latest;
        private Process helper;
        private readonly Thread thread;
        public NativeFilterWatcher(uint processID)
        {
            process = processID;
            thread = new Thread(Loop) { IsBackground = true, Name = "SheetPace menu IPC" };
            thread.Start();
        }
        public void Request() { lock (gate) { if (!stopping) signal.Set(); } }
        public FilterMenu Latest { get { lock (gate) return latest; } }
        private void Loop()
        {
            Process child = null;
            try
            {
                BinaryReader reader = null; BinaryWriter writer = null;
                while (signal.WaitOne())
                {
                    if (stopping) return;
                    if (child == null)
                    {
                        string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SheetPace.NativeMenu.exe");
                        lock (gate)
                        {
                            if (stopping) return;
                            child = Process.Start(new ProcessStartInfo(path, process.ToString())
                            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardInput = true, RedirectStandardOutput = true });
                            helper = child;
                        }
                        reader = new BinaryReader(child.StandardOutput.BaseStream);
                        writer = new BinaryWriter(child.StandardInput.BaseStream);
                    }
                    writer.Write((byte)1); writer.Flush();
                    FilterMenu menu = NativeMenuProtocol.Read(reader);
                    lock (gate) { if (!stopping) latest = menu; }
                }
            }
            catch (Exception ex) { if (!stopping) Log.Write("原生筛选辅助进程异常", ex); }
            finally
            {
                lock (gate)
                {
                    helper = null;
                    if (child != null) { try { child.StandardInput.Close(); if (!child.WaitForExit(1500)) child.Kill(); } catch { } child.Dispose(); }
                    stopping = true; signal.Dispose(); latest = null;
                }
            }
        }
        public void Dispose()
        {
            lock (gate)
            {
                if (stopping) return;
                stopping = true; latest = null; signal.Set();
                // EOF lets UI Automation release its providers before Excel tears down.
                if (helper != null) try { helper.StandardInput.Close(); } catch { }
            }
            // Thread.Join performs COM message pumping on an STA and bounds shutdown time.
            if (thread != Thread.CurrentThread && !thread.Join(2500)) Log.Write("原生筛选辅助进程停止超时", null);
        }
    }
}
