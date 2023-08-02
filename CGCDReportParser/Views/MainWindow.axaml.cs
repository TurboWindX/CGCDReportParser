using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using System.Threading.Tasks;

namespace CGCDReportParser.Views
{
    public partial class MainWindow : Window
    {
        string fpath;
        public MainWindow()
        {
            InitializeComponent();
            fpath = "";
        }

        public async void ButtonClicked(object source, RoutedEventArgs args)
        {
            if (source == Filepicker)
            {
                await FileBrowse();
            }
            if (source == Fileparser)
            {
                await Parser();
            }
            if (source == Pythonreq)
            {
                if (parseconv.IsDocx2PdfInstalled() == true)
                {
                    await PythonBox.Show(this, "The requirements are already installed.", "Python requirements", PythonBox.MessageBoxButtons.Ok);

                }
                else
                {
                    var x = await PythonBox.Show(this, "This will install python requirements using pip.\nIf this executable is not in a venv it will install it System Wide.\nYou only need to do this once.", "Python requirements", PythonBox.MessageBoxButtons.OkCancel);
                    if (x == PythonBox.MessageBoxResult.Ok)
                    {
                        parseconv.InstallDocx2Pdf();
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
                //Debug.WriteLine(fpath);
            }

        }

        public async Task Parser()
        {
            if (fpath != "")
            {
                await parseconv.SplitDocument(fpath);
            }
        }
    }
}