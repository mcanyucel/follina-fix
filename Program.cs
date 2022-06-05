using Microsoft.Win32;
using System.Diagnostics;

Console.WriteLine("================================");
Console.WriteLine("Welcome to Follina Zero-Day Vulnerability Fix");
Console.WriteLine("For more information, see " + @"https://msrc.microsoft.com/update-guide/vulnerability/CVE-2022-30190");
Console.WriteLine("This tool searches registry for vulnerable  'ms-msdt' keys, creates a backup of them, and then deletes them.");
Console.WriteLine("There backup reg file is created under Documents/Follina Fix/fix.reg");
Console.WriteLine("Therefore if you experience issues or if the vulnerability is fixed, you can restore the functionality by restoring the backup");
Console.WriteLine("--------------------------------");

Console.WriteLine("(F)ix or (R)estore?");
var t = Console.ReadKey().Key;
Console.WriteLine();

if (t == ConsoleKey.F)
{

    try
    {
        using (var ms_msdtKey = Registry.ClassesRoot.OpenSubKey("ms-msdt", writable: true))
        {
            if (ms_msdtKey != null)
            {

                Console.WriteLine("You have the following vulnerable keys in your registry, backing them up:");
                Console.WriteLine(ms_msdtKey.Name);
                ms_msdtKey.GetValueNames().ToList().ForEach(q => Console.WriteLine(q));

                var keyPath = $"\"{ms_msdtKey.Name}\"";
                var exportDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Follina Fix");
                if (!Directory.Exists(exportDirectory)) Directory.CreateDirectory(exportDirectory);
                var exportFilePath = Path.Combine(exportDirectory, "fix.reg");

                using (Process proc = new())
                {
                    proc.StartInfo.FileName = "reg.exe";
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.StartInfo.CreateNoWindow = true;
                    proc.StartInfo.Arguments = "export \"" + keyPath + "\" \"" + exportFilePath + "\" /y";
                    proc.Start();
                    string stdout = proc.StandardOutput.ReadToEnd();
                    string stderr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();
                }

                Console.WriteLine("Backing up complete, now trying to delete the entries...");
#pragma warning disable CA1416 // Validate platform compatibility
                ms_msdtKey.DeleteSubKeyTree("");
#pragma warning restore CA1416 // Validate platform compatibility
                Console.WriteLine("Vulnerable key is deleted. Congratulations!");

            }
            else
            {
                Console.WriteLine("You do not have the vulnerable key in your registry. No action needed.");
            }
        }
    }
    catch (Exception e)
    {
        Console.WriteLine($"Failed: {e.Message}");
    }


}
else if (t == ConsoleKey.R)
{
    Console.WriteLine("Restore here");
    var exportFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Follina Fix", "fix.reg");
    if (!File.Exists(exportFilePath))
    {
        Console.WriteLine("The backup file cannot be found; it should be Documents/Follina Fix/fix.reg . Sorry.");
    }
    else
    {
        using Process proc = new();
        proc.StartInfo.FileName = "regedit.exe";
        proc.StartInfo.Arguments = exportFilePath;
        proc.StartInfo.UseShellExecute = false;
        proc.StartInfo.RedirectStandardOutput = true;
        proc.StartInfo.RedirectStandardError = true;
        proc.StartInfo.CreateNoWindow = true;
        proc.Start();
        string stdout = proc.StandardOutput.ReadToEnd();
        string stderr = proc.StandardError.ReadToEnd();
        proc.WaitForExit();

        if (stderr != null)
        {
            Console.WriteLine("Failed to restore backup, try double clicking fix.reg manually from Windows Explorer.");
            using Process expProcess = new();
            expProcess.StartInfo.FileName = "explorer";
            expProcess.StartInfo.Arguments = Path.GetDirectoryName(exportFilePath);
            expProcess.StartInfo.UseShellExecute = false;
            expProcess.StartInfo.RedirectStandardOutput = true;
            expProcess.StartInfo.RedirectStandardError = true;
            expProcess.StartInfo.CreateNoWindow = true;
            expProcess.Start();
            expProcess.WaitForExit();
        }
        else
        {
            Console.WriteLine("You vulnerable backup is restored. Be careful and do not open unknown MSOffice files (especially Word)");
        }
    }
}
Console.ReadKey();
