using NetFwTypeLib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;


namespace Connor
{
    public static class ProductHandler
    {
        public static ObservableCollection<Product> Products { get; } = new ObservableCollection<Product>();
        private static HashSet<String> appsFound = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static IEnumerable<string> ScanApps()
        {
            IEnumerable<string> basePaths = new[]
            {
                Environment.ExpandEnvironmentVariables(@"%PROGRAMFILES%\Adobe"),
                Environment.ExpandEnvironmentVariables(@"%PROGRAMFILES(x86)%\Adobe")
            }
            .Distinct(StringComparer.OrdinalIgnoreCase);
            
            HashSet<string> targetApps = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Photoshop.exe",
                "Illustrator.exe",
                "AfterFX.exe",
                "Premiere.exe",
                "InDesign.exe"
            };

            foreach (var root in basePaths)
            {
                if (!Directory.Exists(root)) continue;

                foreach (var exe in Directory.EnumerateFiles(root, "*.exe", SearchOption.AllDirectories))
                {
                    if (!targetApps.Contains(Path.GetFileName(exe))) continue;

                    if (appsFound.Add(exe))
                        yield return exe;
                }
            }
        }

        public static BitmapImage GetIcon(string executablePath)
        {
            try
            {
                Icon icon = Icon.ExtractAssociatedIcon(executablePath);
                if (icon == null)
                    return null;

                MemoryStream memoryStream = new MemoryStream();
                icon.ToBitmap().Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                memoryStream.Seek(0, SeekOrigin.Begin);

                BitmapImage bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.StreamSource = memoryStream;
                bitmapImage.EndInit();
                bitmapImage.Freeze();

                return bitmapImage;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static async Task LoadProducts()
        {
            Products.Clear();

            var seenPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            INetFwPolicy2 fwPolicy =
                (INetFwPolicy2)Activator.CreateInstance(
                    Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));

            foreach (INetFwRule rule in fwPolicy.Rules)
            {
                if (rule.Action != NET_FW_ACTION_.NET_FW_ACTION_BLOCK)
                    continue;

                if (string.IsNullOrWhiteSpace(rule.Description))
                    continue;

                if (rule.Description.Trim().EndsWith(
                        "limiting its online functionalities.",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                // Should have an application path
                if (string.IsNullOrWhiteSpace(rule.ApplicationName))
                    continue;

                string exePath = rule.ApplicationName.Trim('"');

                // Avoid duplicates
                if (!seenPaths.Add(exePath))
                    continue;

                FileVersionInfo info;
                try
                {
                    info = FileVersionInfo.GetVersionInfo(exePath);
                }
                catch
                {
                    continue; // exe not found or inaccessible
                }

                Products.Add(new Product
                {
                    Name = info.ProductName ??
                           Path.GetFileNameWithoutExtension(exePath),
                    Version = info.ProductVersion ?? "1.0.0",
                    ExecutablePath = exePath,
                    IsFirewallBlocked = true,
                    Icon = await Task.Run(() => GetIcon(exePath))
                });
            }

            foreach (string executable in ScanApps())
            {
                if (Products.Any(p =>
                    p.ExecutablePath.Equals(executable, StringComparison.OrdinalIgnoreCase)))
                    continue;

                FileVersionInfo info = FileVersionInfo.GetVersionInfo(executable);


                Products.Add(new Product
                {
                    Name = info.ProductName ??
                           Path.GetFileNameWithoutExtension(executable),
                    Version = info.ProductVersion ?? "1.0.0",
                    ExecutablePath = executable,
                    Icon = await Task.Run(() => GetIcon(executable)),
                    IsFirewallBlocked = false
                });
            }
        }

        public static async Task LoadFromFirewall()
        {

        }
    }
}