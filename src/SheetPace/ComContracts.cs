using System;
using System.Runtime.InteropServices;
[assembly: ComVisible(false)]
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]
[assembly: System.Reflection.AssemblyFileVersion("1.0.0.0")]
namespace SheetPace
{
    public enum ConnectMode { AfterStartup = 0, Startup = 1, External = 2, CommandLine = 3, Solution = 4, UISetup = 5 }
    public enum DisconnectMode { HostShutdown = 0, UserClosed = 1, UISetupComplete = 2, SolutionClosed = 3 }
    [ComVisible(true), Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)] object application, ConnectMode connectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object addInInst, [In, Out] ref Array custom);
        [DispId(2)] void OnDisconnection(DisconnectMode removeMode, [In, Out] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, Out] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, Out] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, Out] ref Array custom);
    }
    [ComVisible(true), Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID);
    }
}
