using NetFwTypeLib;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Imaging;
using Wpf.Ui.Controls;


namespace Connor
{
    public partial class MainWindow : FluentWindow
    {
        public ObservableCollection<Product> Products => ProductHandler.Products;

        public MainWindow()
        {
            DataContext = this;

            //Wpf.Ui.Appearance.SystemThemeWatcher.Watch(this);

            Wpf.Ui.Appearance.ApplicationThemeManager.Apply(
              Wpf.Ui.Appearance.ApplicationTheme.Dark, // Theme type
              WindowBackdropType.Auto,  // Background type
              true                                      // Whether to change accents automatically
            );
            ProductHandler.LoadProducts();
            InitializeComponent();
        }

        private void ToggleSwitch_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleSwitch toggleSwitch && toggleSwitch.DataContext is Product product)
            {
                FirewallManager.AddRule(product);
            }
        }

        private void ToggleSwitch_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleSwitch toggleSwitch && toggleSwitch.DataContext is Product product)
            {
                FirewallManager.RemoveRule(product);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                FileName = "",
                DefaultExt = ".exe",
                Filter = "Executables (*.exe)|*.exe"
            };

            bool? result = dialog.ShowDialog();

            if ((bool)result)
            {
                string filename = dialog.FileName;
                FileVersionInfo info = FileVersionInfo.GetVersionInfo(filename);
                string productName = info.ProductName ?? Path.GetFileNameWithoutExtension(filename);

                if (ProductHandler.Products.Any(p => p.ExecutablePath == filename))
                {
                    var box = new Wpf.Ui.Controls.MessageBox
                    {
                        Title = "Error",
                        Content = "El producto ya ha sido añadido."
                    };
                    box.ShowDialogAsync();
                    return;
                }
                else
                {
                    Product newProduct = new Product()
                    {
                        Name = productName,
                        Version = info.ProductVersion ?? "1.0.0",
                        ExecutablePath = filename,
                        Icon = ProductHandler.GetIcon(filename)
                    };
                    ProductHandler.Products.Add(newProduct);
                }
            }
        }
    }

    public class Product
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string ExecutablePath { get; set; }
        public bool IsFirewallBlocked { get; set; }
        public BitmapImage Icon { get; set; }
    }

    public static class ProductHandler
    {
        public static ObservableCollection<Product> Products { get; } = new ObservableCollection<Product>();
        private static readonly string[] CommonPaths =
        {
            Environment.ExpandEnvironmentVariables(@"%PROGRAMFILES%\Adobe"),
            Environment.ExpandEnvironmentVariables(@"%PROGRAMFILES(x86)%\Adobe")
        };
        private static readonly HashSet<string> ClayNames = new HashSet<string>()
        {
            "Photoshop.exe",
            "Illustrator.exe",
            "AfterFX.exe",
            "Premiere.exe",
            "InDesign.exe"
        };

        private static IEnumerable<string> FindClay()
        {
            foreach (var path in CommonPaths)
            {
                if (!Directory.Exists(path)) continue;
                foreach (var folder in Directory.EnumerateDirectories(path))
                {
                    IEnumerable<string> executables = Directory.EnumerateFiles(folder, "*.exe", SearchOption.AllDirectories)
                                               .Where(file => ClayNames.Contains(Path.GetFileName(file)));
                    foreach (var executable in executables)
                    {
                        yield return executable;
                    }
                }
            }
        }



        public static async Task LoadProducts()
        {
            Products.Clear();

            INetFwPolicy2 fwPolicy =
                (INetFwPolicy2)Activator.CreateInstance(
                    Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));

            foreach (INetFwRule rule in fwPolicy.Rules)
            {
                if (string.IsNullOrWhiteSpace(rule.Description) ||
                    !rule.Description.Trim().EndsWith(
                        "limiting its online functionalities.",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                if (string.IsNullOrWhiteSpace(rule.ApplicationName))
                    continue;

                string exePath = rule.ApplicationName.Trim('"');

                // Avoid duplicates
                if (Products.Any(p =>
                    p.ExecutablePath.Equals(exePath, StringComparison.OrdinalIgnoreCase)))
                    continue;

                FileVersionInfo info = FileVersionInfo.GetVersionInfo(exePath);

                Products.Add(new Product
                {
                    Name = info.ProductName ??
                           Path.GetFileNameWithoutExtension(exePath),
                    Version = info.ProductVersion ?? "1.0.0",
                    ExecutablePath = exePath,
                    IsFirewallBlocked = true,
                    Icon = await Task.Run(() => GetIcon(exePath)),

                });
            }

            foreach (string executable in FindClay())
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
    }

    public static class FirewallManager
    {
        private static INetFwPolicy2 fwPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));

        public static bool AddRule(Product product)
        {
            try
            {
                Type ruleType = Type.GetTypeFromProgID("HNetCfg.FwRule");

                var inboundRule = CreateRule(product, ruleType, NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN);
                var outboundRule = CreateRule(product, ruleType, NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT);

                fwPolicy.Rules.Add(inboundRule);
                fwPolicy.Rules.Add(outboundRule);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool RemoveRule(Product product)
        {
            try
            {
                var rulesToRemove = fwPolicy.Rules.Cast<INetFwRule>()
                                    .Where(rule => rule.Name == product.Name)
                                    .Select(rule => rule.Name)
                                    .ToList();

                foreach (var ruleName in rulesToRemove)
                {
                    fwPolicy.Rules.Remove(ruleName);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static INetFwRule CreateRule(Product product, Type ruleType, NET_FW_RULE_DIRECTION_ direction)
        {
            INetFwRule rule = (INetFwRule)Activator.CreateInstance(ruleType);
            rule.Name = product.Name;
            rule.Description = $"Blocks {(direction == NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN ? "inbound" : "outbound")} network access for {product.Name}, limiting its online functionalities.";
            rule.ApplicationName = product.ExecutablePath;
            rule.Action = NET_FW_ACTION_.NET_FW_ACTION_BLOCK;
            rule.Direction = direction;
            rule.Enabled = true;
            rule.Profiles = (int)NET_FW_PROFILE_TYPE2_.NET_FW_PROFILE2_ALL;
            return rule;
        }
    }
}