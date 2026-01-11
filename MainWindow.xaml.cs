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
    /// <summary>
    /// Lógica de interacción para MainWindow.xaml
    /// </summary>
    public partial class MainWindow : FluentWindow
    {
        public ObservableCollection<Product> Products => ProductHandler.Products; // Enlazar con la lista observable

        public MainWindow()
        {
            DataContext = this;

            //Wpf.Ui.Appearance.SystemThemeWatcher.Watch(this);

            Wpf.Ui.Appearance.ApplicationThemeManager.Apply(
              Wpf.Ui.Appearance.ApplicationTheme.Dark, // Theme type
              Wpf.Ui.Controls.WindowBackdropType.Auto,  // Background type
              true                                      // Whether to change accents automatically
            );

            InitializeComponent();
        }

        private void ToggleSwitch_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleSwitch toggleSwitch && toggleSwitch.DataContext is Product product)
            {
                FirewallManager.AddFirewallBlockRule(product);
            }
        }

        private void ToggleSwitch_Unchecked(object sender, RoutedEventArgs e)
        {
            if (sender is ToggleSwitch toggleSwitch && toggleSwitch.DataContext is Product product)
            {
                FirewallManager.RemoveFirewallRule(product);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Configurar el diálogo de selección de archivo
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                FileName = "", // Nombre por defecto
                DefaultExt = ".exe", // Extensión por defecto
                Filter = "Executables (*.exe)|*.exe" // Filtro de archivos
            };

            bool? result = dialog.ShowDialog();

            if (result == true)
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
                        IsFirewallBlocked = FirewallManager.FirewallRuleExists(),
                        Icon = ProductHandler.GetIcon(filename)
                    };

                    // Agregarlo a la lista
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
        private static readonly HashSet<string> ExecutableNames = new HashSet<string>()
        {
            "Photoshop.exe",
            "Illustrator.exe",
            "AfterFX.exe",
            "Premiere.exe",
            "InDesign.exe"
        };

        private static IEnumerable<string> FindExecutables()
        {
            foreach (var path in CommonPaths)
            {
                if (!Directory.Exists(path)) continue;
                IEnumerable<string> folders = Directory.EnumerateDirectories(path);
                foreach (var folder in folders)
                {
                    IEnumerable<string> executables = Directory.EnumerateFiles(folder, "*.exe", SearchOption.AllDirectories)
                                               .Where(file => ExecutableNames.Contains(Path.GetFileName(file)));
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
            IEnumerable<string> executables = FindExecutables();
            foreach (string executable in executables)
            {
                FileVersionInfo info = FileVersionInfo.GetVersionInfo(executable);
                string productName = info.ProductName ?? Path.GetFileNameWithoutExtension(executable);

                Product product = new Product()
                {
                    Name = productName,
                    Version = info.ProductVersion ?? "1.0.0",
                    ExecutablePath = executable,
                    IsFirewallBlocked = FirewallManager.FirewallRuleExists(),
                    Icon = await Task.Run(() => GetIcon(executable))
                };
                Products.Add(product);
            }
        }

        public static BitmapImage GetIcon(string executablePath)
        {
            try
            {
                var icon = Icon.ExtractAssociatedIcon(executablePath);
                if (icon == null)
                    return null;

                var bitmap = icon.ToBitmap();
                var memoryStream = new MemoryStream();
                bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                memoryStream.Seek(0, SeekOrigin.Begin);

                var bitmapImage = new BitmapImage();
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
        public static bool FirewallRuleExists()
        {
            try
            {
                INetFwPolicy2 fwPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
                return fwPolicy.Rules.Cast<INetFwRule>()
                    .Any(rule => !string.IsNullOrWhiteSpace(rule.Description) &&
                                 rule.Description.Trim().EndsWith("limiting its online functionalities.", StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool AddFirewallBlockRule(Product product)
        {
            try
            {
                if (FirewallRuleExists())
                {
                    RemoveFirewallRule(product);
                }

                INetFwPolicy2 fwPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
                Type ruleType = Type.GetTypeFromProgID("HNetCfg.FwRule");

                var inboundRule = CreateFirewallRule(product, ruleType, NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_IN);
                var outboundRule = CreateFirewallRule(product, ruleType, NET_FW_RULE_DIRECTION_.NET_FW_RULE_DIR_OUT);

                fwPolicy.Rules.Add(inboundRule);
                fwPolicy.Rules.Add(outboundRule);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool RemoveFirewallRule(Product product)
        {
            try
            {
                INetFwPolicy2 fwPolicy = (INetFwPolicy2)Activator.CreateInstance(Type.GetTypeFromProgID("HNetCfg.FwPolicy2"));
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

        private static INetFwRule CreateFirewallRule(Product product, Type ruleType, NET_FW_RULE_DIRECTION_ direction)
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