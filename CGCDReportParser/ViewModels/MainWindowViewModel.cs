using CGCDReportParser.Models;
using CGCDReportParser.Views;
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;
using ReactiveUI;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;

namespace CGCDReportParser.ViewModels
{
    public class MainWindowViewModel : ViewModelBase, INotifyPropertyChanged
    {
        public string Title => "CGCD - Report parser and converter for Jira";
        public string Deps => "Install LibreOffice";
        public string Pickreport => "Choose report";
        public string Parseconv => "Parse and convert.";
        private double progress;
        private bool visible;
        public bool Visible
        {
            get => visible;
            set
            {
                visible = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Visible"));

            }
        }
        public double Progress
        {
            get => progress;
            set
            {
                progress = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("Progress"));

            }
        }
        public MainWindowViewModel()
        {
            Progress = 0d;
        }

        public event PropertyChangedEventHandler PropertyChanged;


    }
}