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
        private static readonly Dictionary<string, Product> productIndex = new Dictionary<string, Product>(StringComparer.OrdinalIgnoreCase);

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

        private static async Task ScanApps()
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

                foreach (var exePath in Directory.EnumerateFiles(root, "*.exe", SearchOption.AllDirectories))
                {
                    if (!targetApps.Contains(Path.GetFileName(exePath)) || productIndex.ContainsKey(exePath)) continue;
                    
                    FileVersionInfo info;
                    try
                    {
                        info = FileVersionInfo.GetVersionInfo(exePath);
                    }
                    catch
                    {
                        continue;
                    }

                    productIndex[exePath] = new Product
                    {
                        Name = info.ProductName ??
                               Path.GetFileNameWithoutExtension(exePath),
                        Version = info.ProductVersion ?? "1.0.0",
                        ExecutablePath = exePath,
                        IsFirewallBlocked = false, // Default value; will be updated in ScanFirewall
                        Icon = await Task.Run(() => GetIcon(exePath))
                    };
                }
            }
        }


        private static async Task ScanFirewall()
        {
            INetFwPolicy2 fwPolicy =
                (INetFwPolicy2)Activator.CreateInstance(
                    Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));

            foreach (INetFwRule rule in fwPolicy.Rules)
            {
                if (rule.Action != NET_FW_ACTION_.NET_FW_ACTION_BLOCK || 
                    string.IsNullOrWhiteSpace(rule.Description) || 
                    string.IsNullOrWhiteSpace(rule.ApplicationName))
                    continue;

                if (rule.Description.Trim().EndsWith("limiting its online functionalities.", StringComparison.OrdinalIgnoreCase))
                {
                    string exePath = rule.ApplicationName.Trim('"');

                    if (productIndex.TryGetValue(exePath, out Product product))
                    {
                        product.IsFirewallBlocked = true;
                        continue;
                    }

                    FileVersionInfo info;
                    try
                    {
                        info = FileVersionInfo.GetVersionInfo(exePath);
                    }
                    catch
                    {
                        continue;
                    }

                    productIndex[exePath] = new Product
                    {
                        Name = info.ProductName ?? Path.GetFileNameWithoutExtension(exePath),
                        Version = info.ProductVersion ?? "1.0.0",
                        ExecutablePath = exePath,
                        IsFirewallBlocked = true,
                        Icon = await Task.Run(() => GetIcon(exePath))
                    };
                }
            }
        }

        public static async Task InitializeScan()
        {
            await ScanApps();
            await ScanFirewall();

            Products.Clear();
            foreach (var product in productIndex.Values)
                Products.Add(product);
        }
    }
}