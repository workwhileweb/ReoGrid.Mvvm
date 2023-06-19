using Prism.Commands;
using Prism.Mvvm;
using Prism.Regions;
using ReoGrid.Mvvm.Demo.Models;
using ReoGrid.Mvvm.Interfaces;
using System;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using unvell.ReoGrid;

namespace ReoGrid.Mvvm.Demo.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        #region [Fields]
        private string _title = "ReoGrid.Mvvm.Demo";
        private ObservableCollection<IRecordModel> _books;
        private WorksheetModel _worksheetModel;
        #endregion

        #region [Properties]
        public string Title
        {
            get => _title;
            set => SetProperty(ref _title, value);
        }

        public DelegateCommand<ReoGridControl> LoadedCommand { get; set; }
        public DelegateCommand AddRecordCommand    { get; set; }
        public DelegateCommand DeleteRecordCommand { get; set; }
        public DelegateCommand MoveRecordCommand   { get; set; }
        public DelegateCommand EditRecordCommand   { get; set; }
        public DelegateCommand GetFromUiCommand    { get; set; }

        #endregion

        public MainWindowViewModel(IRegionManager regionManager)
        {
            InitBooks();
            InitCommands();
        }

        private void InitBooks()
        {
            _books = new ObservableCollection<IRecordModel>();
            for (var i = 0; i < 10; i++)
            {
                var book = new Book
                {
                    Id = i,
                    Title = $"Title {i}",
                    Author = $"Author {i}",
                    BindingType = BindingType.Hardback,
                    IsOnSale = true,
                    Price = (decimal)(i * 10.1),
                    Publish = DateTime.Now
                };
                _books.Add(book);
            }
        }

        private void InitCommands()
        {
            LoadedCommand = new DelegateCommand<ReoGridControl>(OnLoadedCommand);
            AddRecordCommand = new DelegateCommand(OnAddRecordCommand);
            DeleteRecordCommand = new DelegateCommand(OnDeleteRecordCommand);
            MoveRecordCommand = new DelegateCommand(OnMoveRecordCommand);
            EditRecordCommand = new DelegateCommand(OnEditRecordCommand);
            GetFromUiCommand = new DelegateCommand(OnGetFromUiCommand);
        }

        private void OnLoadedCommand(ReoGridControl reoGridControl)
        {
            _worksheetModel = new WorksheetModel(reoGridControl, typeof(Book), _books);
            _worksheetModel.OnBeforeChangeRecord += WorksheetModel_OnBeforeChangeRecord;
        }

        private static bool? WorksheetModel_OnBeforeChangeRecord(IRecordModel record, System.Reflection.PropertyInfo propertyInfo, object newPropertyValue)
        {
            if (!propertyInfo.Name.Equals("Price")) return null;
            var price = Convert.ToDecimal(newPropertyValue);
            if (price <= 100m) return null; //assume the max price is 100
            MessageBox.Show("Max price is 100.", "Alert", MessageBoxButton.OK, MessageBoxImage.Warning);
            return true; // cancel the change
        }

        private void OnAddRecordCommand()
        {
            var count = _books.Count;
            var book = new Book
            {
                Id = count,
                Title = $"Title {count}",
                Author = $"Author {count}",
                BindingType = BindingType.Hardback,
                IsOnSale = true,
                Price = (decimal)(count * 10.11) > 100m ? 100m :(decimal)(count * 10.11),
                Publish = DateTime.Now
            };
            _books.Add(book);
        }

        private void OnDeleteRecordCommand()
        {
            if (_books.Count <= 0) return;
            _books.RemoveAt(_books.Count - 1);
        }

        private void OnMoveRecordCommand()
        {
            if (_books.Count <= 2) return;
            _books.Move(0, _books.Count - 1);
        }

        private void OnEditRecordCommand()
        {
            ((Book)_books[0]).Price = new Random(DateTime.Now.Millisecond).Next(1,100);
            _worksheetModel.UpdateRecord(_books[0]); // invoke UpdateRecord after editing one record.
        }

        private void OnGetFromUiCommand()
        {
            var result = string.Empty;
            foreach (var recordModel in _books)
            {
                var book = (Book)recordModel;
                result +=
                    $"Id:{book.Id}\t Title:{book.Title}\t Author:{book.Author}\t BindingType:{book.BindingType}\t Price:{book.Price}\t Publish:{book.Publish}\t RowIndex:{book.RowIndex}\r\n";
            }

            var window = new Window
            {
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                WindowStyle = WindowStyle.ToolWindow,
                SizeToContent = SizeToContent.WidthAndHeight
            };
            var textBlock = new TextBlock
            {
                Margin = new Thickness(10),
                Text = result
            };
            window.Content = textBlock;
            window.ShowDialog();
        }
    }
}
