using System;
using SheetPace;
internal static class Program
{
    [MTAThread] private static int Main(string[] args)
    {
        uint process;
        if (args.Length != 1 || !UInt32.TryParse(args[0], out process) || process == 0) return 2;
        return NativeMenuService.Run(process);
    }
}
