using System;
using SheetPace;
internal static class Program
{
    [STAThread] private static int Main(string[] args)
    {
        long window; uint process;
        if (args.Length != 2 || !Int64.TryParse(args[0], out window) || !UInt32.TryParse(args[1], out process)) return 2;
        return VisualHostService.Run(new IntPtr(window), process);
    }
}
