using System;
using System.Runtime.InteropServices;
using System.Xml;
using SheetPace;
class ComAbiTests
{
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int RibbonCall(IntPtr self, IntPtr ribbonId, out IntPtr xml);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int ArrayCall(IntPtr self, ref IntPtr custom);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int DisconnectCall(IntPtr self, int mode, ref IntPtr custom);
    [DllImport("oleaut32.dll")] static extern IntPtr SafeArrayCreateVector(ushort type, int lower, uint count);
    [DllImport("oleaut32.dll")] static extern int SafeArrayDestroy(IntPtr array);
    static int checks;
    static void Check(bool condition, string name)
    { if (!condition) throw new Exception(name); checks++; Console.WriteLine("PASS " + name); }
    static IntPtr Slot(IntPtr instance, int index)
    { return Marshal.ReadIntPtr(Marshal.ReadIntPtr(instance), index * IntPtr.Size); }
    static void Dual(Type type)
    {
        InterfaceTypeAttribute attribute = (InterfaceTypeAttribute)Attribute.GetCustomAttribute(type, typeof(InterfaceTypeAttribute));
        Check(attribute != null && attribute.Value == ComInterfaceType.InterfaceIsDual, type.Name + " dual contract");
    }
    [STAThread] static int Main()
    {
        IntPtr ribbon = IntPtr.Zero, extensibility = IntPtr.Zero, ribbonId = IntPtr.Zero, xml = IntPtr.Zero;
        try
        {
            Dual(typeof(IDTExtensibility2)); Dual(typeof(IRibbonExtensibility));
            foreach (System.Reflection.MethodInfo method in typeof(IDTExtensibility2).GetMethods())
            {
                System.Reflection.ParameterInfo[] parameters = method.GetParameters();
                System.Reflection.ParameterInfo custom = parameters[parameters.Length - 1];
                MarshalAsAttribute attribute = (MarshalAsAttribute)Attribute.GetCustomAttribute(custom, typeof(MarshalAsAttribute));
                Check(attribute != null && attribute.Value == UnmanagedType.SafeArray && attribute.SafeArraySubType == VarEnum.VT_VARIANT && custom.IsIn && !custom.IsOut, method.Name + " Office SAFEARRAY(VARIANT) contract");
            }
            Connect addin = new Connect();
            ribbon = Marshal.GetComInterfaceForObject(addin, typeof(IRibbonExtensibility));
            RibbonCall getUi = (RibbonCall)Marshal.GetDelegateForFunctionPointer(Slot(ribbon, 7), typeof(RibbonCall));
            ribbonId = Marshal.StringToBSTR("Microsoft.Excel.Workbook");
            Marshal.ThrowExceptionForHR(getUi(ribbon, ribbonId, out xml));
            XmlDocument document = new XmlDocument(); document.LoadXml(Marshal.PtrToStringBSTR(xml));
            Check(document.DocumentElement.LocalName == "customUI", "native Ribbon vtable slot 7 returns valid XML BSTR");
            Check(document.OuterXml.Contains("sheetpaceTab"), "native Ribbon contains SheetPace tab");
            Check(document.OuterXml.Contains("label=\"光影设置\"") && !document.OuterXml.Contains("label=\"颜色与透明度\""), "Ribbon settings label renamed");
            Check(document.OuterXml.Contains("getImage=\"GetHoverImage\"") && document.OuterXml.Contains("getImage=\"GetSettingsImage\""), "Ribbon uses dedicated icons for both buttons");
            extensibility = Marshal.GetComInterfaceForObject(addin, typeof(IDTExtensibility2));
            foreach (int slot in new[] { 9, 10, 11 })
            {
                IntPtr custom = IntPtr.Zero;
                ArrayCall callback = (ArrayCall)Marshal.GetDelegateForFunctionPointer(Slot(extensibility, slot), typeof(ArrayCall));
                Marshal.ThrowExceptionForHR(callback(extensibility, ref custom));
                Check(custom == IntPtr.Zero, "native lifecycle vtable slot " + slot);
                IntPtr array = SafeArrayCreateVector((ushort)VarEnum.VT_VARIANT, 0, 2);
                Check(array != IntPtr.Zero, "SAFEARRAY allocation");
                try
                {
                    custom = array; Marshal.ThrowExceptionForHR(callback(extensibility, ref custom));
                    Check(custom == array, "native lifecycle slot " + slot + " receives SAFEARRAY");
                }
                finally { Marshal.ThrowExceptionForHR(SafeArrayDestroy(array)); }
            }
            IntPtr empty = IntPtr.Zero;
            DisconnectCall disconnect = (DisconnectCall)Marshal.GetDelegateForFunctionPointer(Slot(extensibility, 8), typeof(DisconnectCall));
            Marshal.ThrowExceptionForHR(disconnect(extensibility, (int)DisconnectMode.UserClosed, ref empty));
            Check(empty == IntPtr.Zero, "native disconnection vtable slot 8");
            GC.KeepAlive(addin); Console.WriteLine("PASS " + checks + " COM ABI checks (" + (IntPtr.Size * 8) + " bit)"); return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            if (xml != IntPtr.Zero) Marshal.FreeBSTR(xml);
            if (ribbonId != IntPtr.Zero) Marshal.FreeBSTR(ribbonId);
            if (ribbon != IntPtr.Zero) Marshal.Release(ribbon);
            if (extensibility != IntPtr.Zero) Marshal.Release(extensibility);
        }
    }
}
