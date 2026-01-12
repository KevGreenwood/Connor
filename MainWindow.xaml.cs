using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
}