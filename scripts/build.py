"""Offline Windows build; uses the .NET Framework compiler, no NuGet."""
from pathlib import Path
import hashlib, os, subprocess, sys
ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
FRAMEWORK = Path(os.environ.get("WINDIR", "C:/Windows")) / "Microsoft.NET/Framework/v4.0.30319"
CSC = FRAMEWORK / "csc.exe"

def compile_cs(output, sources, references=(), options=()):
    cmd = [str(CSC), "/nologo", "/optimize+", "/platform:anycpu", "/utf8output", "/warn:4", "/out:" + str(output)]
    cmd += ["/reference:" + str(r) for r in references] + list(options) + [str(p) for p in sources]
    print("Building", output.name, flush=True)
    subprocess.run(cmd, check=True, cwd=ROOT)

def main():
    if not CSC.exists():
        raise SystemExit("Requires Windows with .NET Framework 4.8")
    DIST.mkdir(exist_ok=True)
    refs = [FRAMEWORK / (n + ".dll") for n in ["System", "System.Core", "System.Drawing", "System.Windows.Forms", "Microsoft.CSharp"]]
    uia = [FRAMEWORK / "WPF" / (n + ".dll") for n in ["UIAutomationClient", "UIAutomationTypes", "WindowsBase"]]
    dll = DIST / "SheetPace.dll"
    compile_cs(dll, sorted((ROOT / "src/SheetPace").glob("*.cs")), refs + uia, ["/target:library"])
    helper = DIST / "SheetPace.NativeMenu.exe"
    compile_cs(helper, sorted((ROOT / "src/NativeMenu").glob("*.cs")), refs + [dll], ["/target:exe"])
    host = DIST / "SheetPace.VisualHost.exe"
    compile_cs(host, sorted((ROOT / "src/VisualHost").glob("*.cs")), refs + [dll], ["/target:winexe"])
    digest = hashlib.sha256(dll.read_bytes()).hexdigest()
    helper_digest = hashlib.sha256(helper.read_bytes()).hexdigest()
    host_digest = hashlib.sha256(host.read_bytes()).hexdigest()
    (ROOT / "src/Installer/BuildInfo.cs").write_text('namespace SheetPaceInstaller { internal static class BuildInfo { internal const string DllHash = "' + digest + '"; internal const string HelperHash = "' + helper_digest + '"; internal const string HostHash = "' + host_digest + '"; } }', encoding="utf-8")
    compile_cs(DIST / "SheetPace-Setup-1.0.4.exe", sorted((ROOT / "src/Installer").glob("*.cs")), refs,
               ["/target:winexe", "/win32manifest:" + str(ROOT / "src/Installer/app.manifest"), "/resource:" + str(dll) + ",SheetPace.dll", "/resource:" + str(helper) + ",SheetPace.NativeMenu.exe", "/resource:" + str(host) + ",SheetPace.VisualHost.exe"])
    if (ROOT / "tests/CoreTests.cs").exists():
        compile_cs(DIST / "SheetPace.Tests.exe", [ROOT / "tests/CoreTests.cs"], refs + uia + [dll], ["/target:exe"])
        x86tests = DIST / "SheetPace.Tests.x86.exe"
        compile_cs(x86tests, [ROOT / "tests/CoreTests.cs"], refs + uia + [dll], ["/target:exe", "/platform:x86"])
    if (ROOT / "tests/ExcelIntegration.cs").exists():
        compile_cs(DIST / "SheetPace.ExcelTests.exe", [ROOT / "tests/ExcelIntegration.cs"], refs + uia + [dll], ["/target:exe"])
    if (ROOT / "tests/UiProbe.cs").exists():
        compile_cs(DIST / "SheetPace.UiProbe.exe", [ROOT / "tests/UiProbe.cs"], refs + uia + [dll], ["/target:exe"])
    if (ROOT / "tests/UserWorkbookTest.cs").exists():
        compile_cs(DIST / "SheetPace.WorkbookTest.exe", [ROOT / "tests/UserWorkbookTest.cs"], refs + uia + [dll], ["/target:exe"])
    compile_cs(DIST / "SheetPace.RangeTests.exe", [ROOT / "tests/SelectionRangeTests.cs"], refs + [dll], ["/target:exe"])
    compile_cs(DIST / "SheetPace.HoverTests.exe", [ROOT / "tests/HoverGeometryTests.cs"], refs + [dll], ["/target:exe"])
    compile_cs(DIST / "SheetPace.InstalledTests.exe", [ROOT / "tests/InstalledStartupTest.cs"], refs + uia, ["/target:exe"])
    compile_cs(DIST / "SheetPace.ComAbiTests.exe", [ROOT / "tests/ComAbiTests.cs"], refs + [dll], ["/target:exe"])
    compile_cs(DIST / "SheetPace.ComAbiTests.x86.exe", [ROOT / "tests/ComAbiTests.cs"], refs + [dll], ["/target:exe", "/platform:x86"])
    hashes = [hashlib.sha256(p.read_bytes()).hexdigest() + "  " + p.name for p in sorted(DIST.glob("SheetPace-Setup-*.exe"))] + [digest + "  SheetPace.dll", helper_digest + "  SheetPace.NativeMenu.exe", host_digest + "  SheetPace.VisualHost.exe"]
    (DIST / "SHA256SUMS.txt").write_text(chr(10).join(hashes) + chr(10), encoding="utf-8")
    print("Build complete:", DIST)
if __name__ == "__main__": main()
