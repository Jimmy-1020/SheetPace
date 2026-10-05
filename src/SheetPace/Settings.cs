using System;
using System.Drawing;
using System.Globalization;
using System.IO;
namespace SheetPace
{
    public sealed class Settings
    {
        public bool HoverEnabled = true;
        public bool NativeCountsEnabled = true;
        public Color HighlightColor = Color.FromArgb(65, 145, 245);
        public int Transparency = 82;
        public int Alpha { get { return (int)Math.Round(255 * (100 - Transparency) / 100.0); } }
        public static string DirectoryPath { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SheetPace"); } }
        public static string DefaultPath { get { return Path.Combine(DirectoryPath, "settings.ini"); } }
        public static Settings Load(string path)
        {
            Settings result = new Settings();
            try
            {
                if (!File.Exists(path)) return result;
                foreach (string line in File.ReadAllLines(path))
                {
                    int equals = line.IndexOf('='); if (equals < 0) continue;
                    string key = line.Substring(0, equals), value = line.Substring(equals + 1);
                    bool flag; int number;
                    if (key == "HoverEnabled" && bool.TryParse(value, out flag)) result.HoverEnabled = flag;
                    if (key == "NativeCountsEnabled" && bool.TryParse(value, out flag)) result.NativeCountsEnabled = flag;
                    if (key == "Transparency" && int.TryParse(value, out number)) result.Transparency = Math.Max(0, Math.Min(100, number));
                    if (key == "Color" && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out number)) result.HighlightColor = Color.FromArgb(255, Color.FromArgb(number));
                }
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
            return result;
        }
        public void Save(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            string temp = path + ".tmp";
            File.WriteAllLines(temp, new string[] { "HoverEnabled=" + HoverEnabled, "NativeCountsEnabled=" + NativeCountsEnabled,
                "Transparency=" + Transparency, "Color=" + HighlightColor.ToArgb().ToString("X8", CultureInfo.InvariantCulture) });
            if (File.Exists(path)) File.Replace(temp, path, null); else File.Move(temp, path);
        }
        public Settings Copy() { return (Settings)MemberwiseClone(); }
    }
    internal static class Log
    {
        public static void Write(string message, Exception error)
        {
            try
            {
                Directory.CreateDirectory(Settings.DirectoryPath);
                string path = Path.Combine(Settings.DirectoryPath, "diagnostics.log");
                if (File.Exists(path) && new FileInfo(path).Length > 1024 * 1024) File.Delete(path);
                File.AppendAllText(path, DateTime.Now.ToString("s") + " " + message + (error == null ? "" : " " + error.GetType().Name + " " + error.Message) + Environment.NewLine);
            }
            catch { }
        }
    }
}
