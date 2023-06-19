using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using unvell.ReoGrid;
using unvell.ReoGrid.Events;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using unvell.ReoGrid.DataFormat;
using ReoGrid.Mvvm.Attributes;
using ReoGrid.Mvvm.Interfaces;

namespace ReoGrid.Mvvm
{
    public class WorksheetModel
    {

        #region [Fields]

        private readonly ReoGridControl _reoGridControl;
        private readonly Worksheet _worksheet;
        private readonly Type _recordModelType;
        private List<int> _columnWidthList;
        private List<int> _rowHeightList;

        #endregion

        #region [Properties]
        /// <summary>
        /// IRecordModel Collection
        /// </summary>
        public ObservableCollection<IRecordModel> Records { get; set; }

        public delegate bool? BeforeChangeRecordEventHandler(IRecordModel record, PropertyInfo propertyInfo, object newPropertyValue);
        public event BeforeChangeRecordEventHandler OnBeforeChangeRecord;

        public delegate void BeforeCellEditEventHandler(object sender, CellBeforeEditEventArgs e);
        public event BeforeCellEditEventHandler OnBeforeCellEdit;
        #endregion

        #region [Constructor]

        public WorksheetModel(ReoGridControl reoGridControl, Type recordModelType)
        {
            _reoGridControl = reoGridControl;
            _worksheet = reoGridControl.CurrentWorksheet;
            _recordModelType = recordModelType;
            Records = new ObservableCollection<IRecordModel>();

            InitWorksheet();
        }


        public WorksheetModel(ReoGridControl reoGridControl, Type recordModelType, ObservableCollection<IRecordModel> records)
        {
            _reoGridControl = reoGridControl;
            _worksheet = reoGridControl.CurrentWorksheet;
            Records = records;
            _recordModelType = recordModelType;
            Records.CollectionChanged += Records_CollectionChanged;

            InitWorksheet();
        }
        #endregion

        #region [Private Methods]

        #region [InitWorksheet] Init Worksheet
        /// <summary>
        /// Init Worksheet
        /// </summary>
        private void InitWorksheet()
        {
            if (_recordModelType.GetCustomAttribute(typeof(WorksheetAttribute)) is WorksheetAttribute classAttribue)
            {
                _worksheet.Name = classAttribue.Title;
            }
            else
            {
                _worksheet.Name = _recordModelType.Name;
            }

            var properties = _recordModelType.GetProperties();
            var ColHeaderAttributeDict = new Dictionary<PropertyInfo, ColumnHeaderAttribute>();
            foreach (var property in properties)
            {
                if (property.GetCustomAttribute(typeof(ColumnHeaderAttribute)) is ColumnHeaderAttribute headerAttribute && headerAttribute.IsVisible) //filter invisible item
                {
                    ColHeaderAttributeDict.Add(property, headerAttribute);
                }
            }
            if (ColHeaderAttributeDict.Count < 1)
            {
#if DEBUG
                Console.WriteLine("InitWorksheet Failed: HeaderAttributes.Count is 0.");
#endif
                return;
            }
            ColHeaderAttributeDict = ColHeaderAttributeDict.OrderBy(one => one.Value.Index).ToDictionary(one => one.Key, one => one.Value); // order by index
            // Re-Set Index
            for (var i = 0; i < ColHeaderAttributeDict.Keys.Count; i++)
            {
                var key = ColHeaderAttributeDict.Keys.ElementAt(i);
                ColHeaderAttributeDict[key].Index = i;
            }

           
            _worksheet.Columns = ColHeaderAttributeDict.Count;
           

            var rangePosition = new RangePosition
            {
                Cols = 1,
                Row = 0,
                Rows = _worksheet.RowCount
            };

            for (var i = 0; i < properties.Count(); i++)
            {
                var property = properties[i];
                var attribute = property.GetCustomAttribute(typeof(FormatAttributeBase));
                if (attribute == null)
                {
                    continue;
                }

                if (attribute is IFormatArgs formatArgs)
                {
                    var headerAttribute = (from key in ColHeaderAttributeDict.Keys
                             where key.Equals(property)
                             select ColHeaderAttributeDict[property]).FirstOrDefault();

                    if (headerAttribute != null && headerAttribute.IsVisible)
                    {
                        rangePosition.Col = headerAttribute.Index;
                        //not work correctly 
                        switch (formatArgs.CellDataFormatFlag)
                        {
                            case CellDataFormatFlag.General:
                                break;
                            case CellDataFormatFlag.Number:
                                {
                                    var numberFormatAttribute = formatArgs as NumberFormatAttribute;
                                    var numberFormatter = new NumberDataFormatter.NumberFormatArgs();
                                    if (numberFormatAttribute.DecimalPlaces != short.MaxValue)
                                    {
                                        numberFormatter.DecimalPlaces = numberFormatAttribute.DecimalPlaces;
                                    }
                                    numberFormatter.NegativeStyle = numberFormatAttribute.NegativeStyle;
                                    numberFormatter.UseSeparator = numberFormatAttribute.UseSeparator;
                                    numberFormatter.CustomNegativePrefix = numberFormatAttribute.CustomNegativePrefix;
                                    numberFormatter.CustomNegativePostfix = numberFormatAttribute.CustomNegativePostfix;
                                    _worksheet.SetRangeDataFormat(rangePosition, CellDataFormatFlag.Number, numberFormatter);
                                    //_ReoGridControl.DoAction(new SetRangeDataFormatAction(rangePosition, CellDataFormatFlag.Number, numberFormatter));
                                    break;
                                }
                            case CellDataFormatFlag.DateTime:
                                {
                                    var dateTimeFormatAttribute = formatArgs as DateTimeFormatAttribute;
                                    var dateTimeFormatArgs = new DateTimeDataFormatter.DateTimeFormatArgs
                                        {
                                            Format = dateTimeFormatAttribute.Format,
                                            CultureName = dateTimeFormatAttribute.CultureName
                                        };
                                    _worksheet.SetRangeDataFormat(rangePosition, CellDataFormatFlag.DateTime, dateTimeFormatArgs);
                                    break;
                                }
                            case CellDataFormatFlag.Percent:
                                break;
                            case CellDataFormatFlag.Currency:
                                break;
                            case CellDataFormatFlag.Text:
                                break;
                            case CellDataFormatFlag.Custom:
                                break;
                            default:
                                break;
                        }
                    }
                }

                
            }

            _columnWidthList = new List<int>();
            _rowHeightList = new List<int>();

            for (var i = 0; i < ColHeaderAttributeDict.Count; i++)
            {
                var headerAttribute = ColHeaderAttributeDict.ElementAt(i).Value;

                _worksheet.ColumnHeaders[i].Text = string.IsNullOrEmpty(headerAttribute.Text) ? ColHeaderAttributeDict.ElementAt(i).Key.Name : headerAttribute.Text;
                if (headerAttribute.Width <= 0)
                {
                    _worksheet.ColumnHeaders[i].IsAutoWidth = true;
                    
                }
                else
                {
                    _worksheet.ColumnHeaders[i].Width = (ushort)headerAttribute.Width;
                }
                _worksheet.ColumnHeaders[i].IsVisible = true;
                _worksheet.ColumnHeaders[i].Tag = ColHeaderAttributeDict.ElementAt(i).Key; // Tag stores PropertyInfo of Model
                _columnWidthList.Add(_worksheet.ColumnHeaders[i].Width); // store column width
            }

            //stroe row height
            for (var i = 0; i < _worksheet.Rows; i++)
            {
                _rowHeightList.Add(_worksheet.RowHeaders[i].Height);
            }

            // load data
            LoadRecords();


            _worksheet.BeforeCellEdit += Worksheet_BeforeCellEdit;
            _worksheet.CellDataChanged += Worksheet_CellDataChanged;
            _worksheet.RangeDataChanged += Worksheet_RangeDataChanged;
            _worksheet.AfterPaste += Worksheet_AfterPaste;

            _worksheet.RowsHeightChanged += Worksheet_RowsHeightChanged;
            _worksheet.ColumnsWidthChanged += Worksheet_ColumnsWidthChanged;
        }
        #endregion

        #region [Worksheet_ColumnsWidthChanged] Worksheet Columns Width Changed
        /// <summary>
        /// Worksheet Columns Width Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worksheet_ColumnsWidthChanged(object sender, ColumnsWidthChangedEventArgs e)
        {
            // Exception: Process is terminated due to StackOverflowException.
            // _Worksheet.ColumnHeaders[e.Index].Width = (ushort)e.Width;

            if (e.Index < _columnWidthList.Count)
            {
                _columnWidthList[e.Index] = e.Width;
            }
            else
            {
                _columnWidthList.Add(e.Width);
            }
        }
        #endregion

        #region [Worksheet_RowsHeightChanged] Worksheet Rows Height Changed
        /// <summary>
        /// Worksheet Rows Height Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worksheet_RowsHeightChanged(object sender, RowsHeightChangedEventArgs e)
        {
            // Exception: Process is terminated due to StackOverflowException.
            // _Worksheet.RowHeaders[e.Index].Height = (ushort)e.Height;

            if (e.Row < _rowHeightList.Count)
            {
                _rowHeightList[e.Row] = e.Height;
            }
            else
            {
                _rowHeightList.Add(e.Height);
            }
        }
        #endregion

        #region [LoadRecords] Load Records
        /// <summary>
        /// Load Records
        /// </summary>
        private void LoadRecords()
        {
            if (Records.Count <= 0) return;
            for (var rowIndex = 0; rowIndex < Records.Count; rowIndex++)
            {
                var record = Records.ElementAt(rowIndex);
                AddOrUpdateOneFromRecord(rowIndex, record);
            }
        }
        #endregion

        #region [Records_CollectionChanged] Data records collection changed
        /// <summary>
        /// Data records collection changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Records_CollectionChanged(object sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e)
        {
            // Changed event, _Records has got action result
            switch (e.Action)
            {
                case System.Collections.Specialized.NotifyCollectionChangedAction.Add:
                    {
                        foreach (IRecordModel item in e.NewItems)
                        {
                            AddOrUpdateOneFromRecord(Records.Count - 1, item);
                        }
                        break;
                    }
                case System.Collections.Specialized.NotifyCollectionChangedAction.Remove:
                    {
                        foreach (IRecordModel item in e.OldItems)
                        {
                            var rowIndex = item.RowIndex;
                            _worksheet.DeleteRows(rowIndex, 1);

                            for (var i = rowIndex; i < Records.Count; i++)
                            {
                                Records.ElementAt(rowIndex).RowIndex = rowIndex;
                            }
                        }
                        break;
                    }
                case System.Collections.Specialized.NotifyCollectionChangedAction.Replace:
                    {
                        for (var i = 0; i < e.NewItems.Count; i++)
                        {
                            var item = e.NewItems[i] as IRecordModel;
                            var rowIndex = (e.OldItems[i] as IRecordModel).RowIndex;
                            AddOrUpdateOneFromRecord(rowIndex, item);
                        }
                        break;
                    }
                case System.Collections.Specialized.NotifyCollectionChangedAction.Move:
                    {
                        for (var i = 0; i < Records.Count; i++)
                        {
                            var item = Records.ElementAt(i);
                            var rowIndex = item.RowIndex;
                            if (rowIndex != i)
                            {
                                AddOrUpdateOneFromRecord(i, item);
                            }
                        }
                        break;
                    }
                case System.Collections.Specialized.NotifyCollectionChangedAction.Reset:
                    {
                        _worksheet.DeleteRows(0, _worksheet.UsedRange.Rows);
                        break;
                    }
                default:
                    break;
            }
        }
        #endregion

        #region [Worksheet_BeforeCellEdit] Before editing cell data
        /// <summary>
        /// Before editing cell data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worksheet_BeforeCellEdit(object sender, CellBeforeEditEventArgs e)
        {

            var currentColIndex = e.Cell.Column;
            var currentRowIndex = e.Cell.Row;
            
            if (_worksheet.ColumnHeaders[currentColIndex] != null)
            {
                var propertyInfo = (PropertyInfo)_worksheet.ColumnHeaders[currentColIndex].Tag;
                if (propertyInfo.PropertyType.BaseType == typeof(Enum))
                {
                    // get enum values
                    var enumValues = Enum.GetValues(propertyInfo.PropertyType).Cast<object>().ToList();
                    SimulateComboBox(enumValues, e);
                }
                else if (propertyInfo.PropertyType == typeof(bool))
                {
                    var values = new List<object>() { bool.TrueString, bool.FalseString };
                    SimulateComboBox(values, e);
                }
            }

            OnBeforeCellEdit?.Invoke(sender, e);
        }
        #endregion

        #region [Worksheet_CellDataChanged] Worksheet Cell Data Changed
        /// <summary>
        /// Worksheet Cell Data Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worksheet_CellDataChanged(object sender, CellEventArgs e)
        {
            if (e.Cell == null) return;
            var row = e.Cell.Row;
            var col = e.Cell.Column;
                
            AddOrUpdateOneFromUi(row, col, col);
        }
        #endregion

        #region [Worksheet_AfterPaste] Paste Data
        /// <summary>
        /// Paste Data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worksheet_AfterPaste(object sender, RangeEventArgs e)
        {
            var range = e.Range;
            AddOrUpdateRecords(range);
        }
        #endregion

        #region [AddOrUpdateRecords] Add or Update mulit records
        /// <summary>
        /// </summary>
        /// <param name="range">edit range</param>
        private void AddOrUpdateRecords(RangePosition range)
        {
            for (var rowIndex = range.StartPos.Row; rowIndex <= range.EndRow; rowIndex++)
            {
                AddOrUpdateOneFromUi(rowIndex, range.StartPos.Col, range.EndPos.Col);
            }
        }
        #endregion

        #region [AddOrUpdateOneFromUi] Add or Update one record
        /// <summary>
        /// Add or Update one record
        /// </summary>
        /// <param name="rowIndex">current row index</param>
        /// <param name="startCol">start column index</param>
        /// <param name="endCol">end column index</param>
        private void AddOrUpdateOneFromUi(int rowIndex, int startCol, int endCol)
        {
            IRecordModel record;

            if (rowIndex < Records.Count) // update record
            {
                record = Records.ElementAt(rowIndex);
            }
            else // insert record
            {
                record = Activator.CreateInstance(_recordModelType) as IRecordModel;
                record.RowIndex = rowIndex;
            }

            for (var colIndex = startCol; colIndex <= endCol; colIndex++)
            {
                var propertyInfo = (PropertyInfo)_worksheet.ColumnHeaders[colIndex].Tag;
                var cellData = _worksheet.GetCellData(rowIndex, colIndex);
                object value = null;
                if (cellData != null)
                {
                    if (propertyInfo.PropertyType.BaseType == typeof(Enum))
                    {
                        try
                        {
                            value = Convert.ChangeType(Enum.Parse(propertyInfo.PropertyType, cellData.ToString()), propertyInfo.PropertyType);
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            Console.WriteLine(ex.Message);
                            break;
#endif
                        }
                    }
                    else
                    {
                        try
                        {
                            value = Convert.ChangeType(cellData, propertyInfo.PropertyType);
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            Console.WriteLine(ex.Message);
                            break;
#endif
                        }

                    }
                }

                if (OnBeforeChangeRecord != null)
                {
                    var isCancel = OnBeforeChangeRecord(record, propertyInfo, value);
                    if (isCancel.HasValue && isCancel.Value) // if has value and cancel is true, then undo the change
                    {
                        _worksheet.SetCellData(rowIndex, colIndex, propertyInfo.GetValue(record));
                        //if (rowIndex < _Records.Count)
                        //{
                        //    _Worksheet.SetCellData(rowIndex, colIndex, propertyInfo.GetValue(record));
                        //}
                        //else
                        //{

                        //}
                        continue;
                    }
                }
                propertyInfo.SetValue(record, value);
            }
            if (rowIndex >= Records.Count)
            {
                Records.Add(record);
            }
        }
        #endregion

        #region [AddOrUpdateOneFromRecord] Add or update one IRecordModel object into Worksheet
        /// <summary>
        /// Add or update one IRecordModel object into Worksheet
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="record"></param>
        private void AddOrUpdateOneFromRecord(int rowIndex, IRecordModel record)
        {
            _worksheet.SuspendDataChangedEvents();
            record.RowIndex = rowIndex; // set row index
            for (var colIndex = 0; colIndex < _worksheet.Columns; colIndex++)
            {
                var currentCellPos = new CellPosition(rowIndex, colIndex);
                if (_worksheet.ColumnHeaders[colIndex].Tag != null)
                {
                    var propertyInfo = (PropertyInfo)_worksheet.ColumnHeaders[colIndex].Tag;
                    var data = propertyInfo.GetValue(record);
                    _worksheet.SetCellData(currentCellPos, data);
                }
                else
                {
#if DEBUG
                    Console.WriteLine("LoadRecords Error: index of columns[{0}] is null!", colIndex);
#endif
                }
            }
            _worksheet.ResumeDataChangedEvents();
        }
        #endregion

        #region [Worksheet_RangeDataChanged] Worksheet Range Data Changed
        /// <summary>
        /// Worksheet Range Data Changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Worksheet_RangeDataChanged(object sender, RangeEventArgs e)
        {
            var isDeleteRows = false;
            if (e.Range.Cols >= _worksheet.UsedRange.Cols) //delete whole rows
            {
                isDeleteRows = true;
                for (var rowIndex = e.Range.StartPos.Row; rowIndex <= e.Range.EndRow; rowIndex++)
                {
                    var recordModel = (from one in Records where one.RowIndex == rowIndex select one).FirstOrDefault();
                    if (recordModel == null) continue;
                    Records.CollectionChanged -= Records_CollectionChanged;
                    Records.Remove(recordModel);
                    Records.CollectionChanged += Records_CollectionChanged;
                }

                if (!isDeleteRows) return;
                _worksheet.DeleteRows(e.Range.StartPos.Row, e.Range.Rows);

                for (var i = e.Range.StartPos.Row; i < Records.Count; i++)
                {
                    Records.ElementAt(i).RowIndex = i;
                }
            }
            else //delete parts of rows
            {
                for (var rowIndex = e.Range.StartPos.Row; rowIndex <= e.Range.EndRow; rowIndex++)
                {
                    AddOrUpdateOneFromUi(rowIndex, e.Range.StartPos.Col, e.Range.EndCol);
                }
            }
        }
        #endregion

        #region [SimulateComboBox] Simulate ComboBox
        private void SimulateComboBox(List<object> list, CellBeforeEditEventArgs e)
        {
            var currentColIndex = e.Cell.Column;
            var currentRowIndex = e.Cell.Row;
            
            var window = new Window();
            var wrapPanel = new WrapPanel();
            var listBox = new ListBox();
            listBox.SetValue( ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Hidden);
            listBox.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Hidden);
            listBox.Width = _columnWidthList.ElementAt(e.Cell.Column);

            for (var i = 0; i < list.Count; i++)
            {
                var item = list.ElementAt(i);
                var listBoxItem = new ListBoxItem
                {
                    Content = item
                };
                listBox.Items.Add(listBoxItem);
            }
            listBox.RenderTransform = new ScaleTransform(_worksheet.ScaleFactor, _worksheet.ScaleFactor);

            listBox.MouseDoubleClick += (obj, eventArgs) => {
                e.EditText = (listBox.SelectedValue as ListBoxItem)?.Content.ToString();
                window.DialogResult = true;
            };
            wrapPanel.Children.Add(listBox);
            var point = new Point();
            for (var rowIndex = 0; rowIndex <= currentRowIndex + 1; rowIndex++)
            {
                point.Y += (int)(_rowHeightList.ElementAt(rowIndex) * _worksheet.ScaleFactor);
            }
            for (var colIndex = 0; colIndex < currentColIndex; colIndex++)
            {
                point.X += (int)(_columnWidthList.ElementAt(colIndex) * _worksheet.ScaleFactor);
            }
            point.X += (int)(_worksheet.RowHeaderWidth * _worksheet.ScaleFactor);
            var screenPoint = _reoGridControl.PointToScreen(point);

            window.Width = _worksheet.ColumnHeaders[e.Cell.Column].Width;
            window.WindowStyle = WindowStyle.None;
            window.ResizeMode = ResizeMode.NoResize;
            window.BorderThickness = new Thickness(0);
            window.Content = wrapPanel;
            window.SizeToContent = SizeToContent.WidthAndHeight;
            window.WindowStartupLocation = WindowStartupLocation.Manual;
            window.Left = screenPoint.X;
            window.Top = screenPoint.Y;
            window.Loaded += (win, routedEventArgs) =>
            {
                wrapPanel.Width = listBox.RenderSize.Width * _worksheet.ScaleFactor;
                wrapPanel.Height = listBox.RenderSize.Height * _worksheet.ScaleFactor;
            };
            window.ShowDialog();
        }
        #endregion

        #endregion

        #region [Public Methods]

        #region [UpdateRecord] Update one record
        /// <summary>
        /// Update one record
        /// </summary>
        /// <param name="recordModel"></param>
        public void UpdateRecord(IRecordModel recordModel)
        {
            AddOrUpdateOneFromRecord(recordModel.RowIndex, recordModel);
        }
        #endregion

        #endregion
    }
}
