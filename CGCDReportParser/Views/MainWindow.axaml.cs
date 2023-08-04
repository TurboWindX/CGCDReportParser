using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using CGCDReportParser.Models;
using CGCDReportParser.ViewModels;
using DocumentFormat.OpenXml.InkML;
using System;
using System.ComponentModel;
using System.Threading;
using System.Threading.Tasks;

namespace CGCDReportParser.Views
{
    public partial class MainWindow : Window
    {
        string fpath;
        parseconv pc;
        bool parsing;
        public MainWindow()
        {
            InitializeComponent();
            pc = new parseconv();
            fpath = "";
            parsing = false;
        }
        public async Task VisibleProg(bool val)
        {
            //((MainWindowViewModel)DataContext).Progress = val;
            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
            {
                ((MainWindowViewModel)DataContext).Visible = val;
            });
        }
        public async Task UpdateProg(double val)
        {
            //((MainWindowViewModel)DataContext).Progress = val;
            await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() =>
            {
                ((MainWindowViewModel)DataContext).Progress = val;
            });
        }
        public async void ButtonClicked(object source, RoutedEventArgs args)
        {

            if (source == Filepicker)
            {
                await FileBrowse();
            }
            if (source == Fileparser)
            {
                if(parsing == false)
                {
                    if (fpath != "")
                    {
                        Task.Run(() => Parser());
                        parsing = true;
                    }
                    
                }

                //await Parser();
            }
            if (source == Deps)
            {
                if (deps.IsSOfficeAvailable() == true)
                {
                    await PythonBox.Show(this, "LibreOffice is already seems already installed.", "Deps", PythonBox.MessageBoxButtons.Ok);
                    
                }
                if(deps.IsSOfficeAvailable() == false)
                {
                    var x = await PythonBox.Show(this, "This will download and install LibreOffice on your system. It is used to convert the parsed DocX to PDFs.", "Deps", PythonBox.MessageBoxButtons.OkCancel);
                    if (x == PythonBox.MessageBoxResult.Ok)
                    {
                        deps.InstallDependencies();
                    }
                }
            }
        }


        public async Task FileBrowse()
        {
            var fld = new Window().StorageProvider;
            var opts = new FilePickerOpenOptions();

            var result = await fld.OpenFilePickerAsync(opts);

            if (result?.Count > 0)
            {
                fpath = result[0].Path.AbsolutePath;
                fpath = System.Web.HttpUtility.UrlDecode(fpath);
                //Debug.WriteLine(fpath);
            }

        }

        

        public async void Parser()
        {
            
            await VisibleProg(true);
            await UpdateProg(10d);
            pc.SplitDocumentAsync(fpath);
            while (pc.Done == false)
            {
                await UpdateProg(pc.Progress);
                await Task.Delay(100);
            }
            await Task.Delay(4321);
            await UpdateProg(100d);
            await Task.Delay(1000);
            await VisibleProg(false);
            parsing = false;
        }
    }
}