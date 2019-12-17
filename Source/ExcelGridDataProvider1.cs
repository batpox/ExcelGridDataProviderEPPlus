using SimioAPI.Extensions;
using System;
using DevExpress.Spreadsheet;
using System.IO;
using System.Collections.Generic;


namespace ExcelGridDataProvider
{
    public class ExcelGridDataPovider1 : IGridDataProvider,IGridDataProviderWithFiles
    {
        #region IGridDataProvider Members

        public string Name
        {
            get { return "Excel"; }
        }

        public string Description
        {
            get { return "Reads data from an Excel spreadsheet"; }
        }

        public System.Drawing.Image Icon
        {
            get { return Properties.Resources.Icon; }
        }

        public Guid UniqueID
        {
            get { return MY_ID; }
        }
        static readonly Guid MY_ID = new Guid("5C9020DF-7B3F-4659-BDBB-EAE5C67FDE21");

        public byte[] GetDataSettings(byte[] existingSettings)
        {
            ExcelGridDataSettings thesettings = ExcelGridDataSettings.FromBytes(existingSettings);
            if (thesettings == null)
                thesettings = new ExcelGridDataSettings();

            SettingsDialog dlg = new SettingsDialog();
            dlg.SetSettings(thesettings);

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                return thesettings.ToBytes();

            return existingSettings;
        }

        public IGridDataRecords OpenData(byte[] dataSettings, IGridDataOpenContext openContext)
        {
            ExcelGridDataSettings thesettings = ExcelGridDataSettings.FromBytes(dataSettings);
            if (thesettings == null || thesettings.FileName == null) // and maybe check that file exists and can be opened, etc?
                return null;

            thesettings.TryFindFile(openContext);
            return new ExcelGridDataRecords(thesettings, openContext);
        }

        public string GetDataSummary(byte[] dataSettings)
        {
            ExcelGridDataSettings thesettings = ExcelGridDataSettings.FromBytes(dataSettings);
            if (thesettings == null) // and maybe check that file exists and can be opened, etc?
                return null;
            string fileName = thesettings.FileName;
            if (System.IO.File.Exists(thesettings.FileName) == false) //if portal or desktop 
            {
                //send only filename not path for portal
                fileName = System.IO.Path.GetFileName(thesettings.FileName);
            }
            var summary = String.Format("Bound to Excel: {0}", fileName ?? "[No file name]");

            if (thesettings.IsNamedRange && thesettings.NamedRange != null)
                summary += String.Format(", Named Range: {0}", thesettings.NamedRange);
            else if (thesettings.IsSpecificRange && thesettings.Worksheet != null && thesettings.SpecificRange != null)
                summary += String.Format(", Worksheet: {0}, Range: {1}", thesettings.Worksheet, thesettings.SpecificRange);
            else if (thesettings.IsWorksheetRange && thesettings.Worksheet != null)
                summary += String.Format(", Worksheet: {0}", thesettings.Worksheet);

            return summary;
        }

        public string[] GetFileNamesIfAny(byte[] dataSettings)
        {
            ExcelGridDataSettings thesettings = ExcelGridDataSettings.FromBytes(dataSettings);
            if (thesettings == null) 
                return null;
            return new string[] { thesettings.FileName };
        }

        #endregion
    }

    struct RType
    {
        public const int WORKSHEET = 0;
        public const int SPECIFIC_RANGE = 1;
        public const int NAMED_RANGE = 2;
    }

    //// See http://support.microsoft.com/kb/320369
    //class ExcelCultureFence : IDisposable
    //{
    //    System.Globalization.CultureInfo _cInfo;

    //    public ExcelCultureFence()
    //    {
    //        _cInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
    //        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
    //    }

    //    #region IDisposable Members

    //    public void Dispose()
    //    {
    //        System.Threading.Thread.CurrentThread.CurrentCulture = _cInfo;
    //    }

    //    #endregion
    //}

    [Serializable]
    class ExcelGridDataSettings
    {
        string _fileName;
        public string FileName
        {
            get { return _fileName; }
            set
            {
                if (_fileName == value)
                    return;

                _fileName = value;

                // Clear out the existing selected worksheet
                Worksheet = null;

                // Clear out the existing selected named range
                if (IsNamedRange)
                    NamedRange = null;

                LoadLists();

                if (_worksheets.Count > 0)
                    Worksheet = _worksheets[0];
                if (_namedRanges.Count > 0)
                    NamedRange = _namedRanges[0];
            }
        }

        public void TryFindFile(IGridDataOpenContext openContext)
        {
            if (System.IO.File.Exists(_fileName) == false)
            {
                if (String.IsNullOrEmpty(openContext.ProjectFileName) == false)
                {
                    var newFile = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(openContext.ProjectFileName), System.IO.Path.GetFileName(_fileName));
                    if (System.IO.File.Exists(newFile))
                    {
                        _fileName = newFile;
                    }
                }
            }
        }

        void LoadLists()
        {
            if (_worksheets == null)
                _worksheets = new System.ComponentModel.BindingList<string>();
            if (_namedRanges == null)
                _namedRanges = new System.ComponentModel.BindingList<string>();

            // Clear the list of available worksheets
            _worksheets.Clear();
            // Clear the list of available named ranges
            _namedRanges.Clear();

            // Fill the list of available worksheets from the new file
            if (System.IO.File.Exists(_fileName))
            {
                Workbook workbook = new Workbook();

                // Load a workbook
                workbook.LoadDocument(_fileName, DocumentFormat.OpenXml);

                // Access a collection of worksheets.
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    _worksheets.Add(workbook.Worksheets[i].Name);
                }

                for (int i = 0; i < workbook.DefinedNames.Count; i++)
                {
                    if (!string.IsNullOrEmpty(workbook.DefinedNames[i].RefersTo) && (workbook.DefinedNames[i].Range != null))
                    {
                        //filter formula                      
                        _namedRanges.Add(workbook.DefinedNames[i].Name);
                    }
                }
            }
        }

        string _worksheet;
        public string Worksheet
        {
            get { return _worksheet; }
            set { _worksheet = value; }
        }

        [NonSerialized]
        System.ComponentModel.BindingList<string> _worksheets = new System.ComponentModel.BindingList<string>();
        public System.ComponentModel.IBindingList Worksheets
        {
            get
            {
                if (_worksheets == null)
                    LoadLists();

                return _worksheets;
            }
        }

        string _namedRange;
        public string NamedRange
        {
            get { return _namedRange; }
            set { _namedRange = value; }
        }

        [NonSerialized]
        System.ComponentModel.BindingList<string> _namedRanges = new System.ComponentModel.BindingList<string>();
        public System.ComponentModel.IBindingList NamedRanges
        {
            get
            {
                if (_namedRanges == null)
                    LoadLists();

                return _namedRanges;
            }
        }

        const string DEFAULT_SPECIFIC_RANGE = "A1:B10";
        string _specificRange = DEFAULT_SPECIFIC_RANGE;
        public string SpecificRange
        {
            get { return _specificRange; }
            set { _specificRange = value; }
        }

        int _rangeType;
        public int RangeType
        {
            get { return _rangeType; }
            set { _rangeType = value; }
        }

        public bool IsWorksheetRange
        {
            get { return RangeType == RType.WORKSHEET; }
            set { RangeType = RType.WORKSHEET; }
        }
        public bool IsSpecificRange
        {
            get { return RangeType == RType.SPECIFIC_RANGE; }
            set { RangeType = RType.SPECIFIC_RANGE; }
        }
        public bool IsNamedRange
        {
            get { return RangeType == RType.NAMED_RANGE; }
            set { RangeType = RType.NAMED_RANGE; }
        }

        public static ExcelGridDataSettings FromBytes(byte[] settings)
        {
            if (settings == null)
                return null;

            System.IO.MemoryStream memstream = new System.IO.MemoryStream(settings);
            System.Runtime.Serialization.Formatters.Binary.BinaryFormatter fmt = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();

            ExcelGridDataSettings excelsettings = (ExcelGridDataSettings)fmt.Deserialize(memstream);

            return excelsettings;
        }
        public byte[] ToBytes()
        {
            System.IO.MemoryStream memstream = new System.IO.MemoryStream();
            System.Runtime.Serialization.Formatters.Binary.BinaryFormatter fmt = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();

            fmt.Serialize(memstream, this);

            return memstream.ToArray();
        }
    }

    class ExcelGridDataRecords : IGridDataRecords
    {
        Worksheet _sheet;
        Workbook _workbook;
        int _startRowIndex, _startColumnIndex;
        int _lastRowIndex, _lastColumnIndex;
        const string __TimeStamp = "__TimeStamp";

        public ExcelGridDataRecords(ExcelGridDataSettings settings, IGridDataOpenContext openContext)
        {
            // Do we have a cached workbook?
            _workbook = GetCachedWorkbook(settings.FileName, openContext);
            if (_workbook == null)
            {
                // Load, or re-load from the file.
                _workbook = new Workbook();

                // Load a workbook
                _workbook.LoadDocument(settings.FileName, DocumentFormat.OpenXml);

                // Cache the in-memory workbook, as well as when we loaded it.
                openContext.SetNamedValue(settings.FileName, _workbook);
                openContext.SetNamedValue(settings.FileName + __TimeStamp, DateTime.Now);
            }

            if (_workbook != null)
            {
                if (settings.IsWorksheetRange)
                {
                    _startRowIndex = 0;
                    _startColumnIndex = 0;
                    _sheet = _workbook.Worksheets[settings.Worksheet];

                    if (_sheet != null)
                    {
                        Range usedRange = _sheet.GetDataRange();
                        _lastRowIndex = usedRange.RowCount;
                        _lastColumnIndex = usedRange.ColumnCount;
                    }
                }
                else if (settings.IsNamedRange || settings.IsSpecificRange)
                {
                    Range aref = null;
                    _sheet = _workbook.Worksheets[0];

                    if (settings.IsNamedRange)
                    {
                        DefinedName definedName = _workbook.DefinedNames.GetDefinedName(settings.NamedRange);
                        aref = definedName.Range;
                        _sheet = aref.Worksheet;

                    }
                    else
                    {
                        _sheet = _workbook.Worksheets[settings.Worksheet];
                        if (_sheet != null)
                        {
                            aref = _sheet.Range[settings.SpecificRange];
                        }
                    }

                    if (aref != null)
                    {
                        _lastRowIndex = aref.BottomRowIndex + 1;
                        _lastColumnIndex = aref.RightColumnIndex + 1;
                        _startRowIndex = aref.TopRowIndex;
                        _startColumnIndex = aref.LeftColumnIndex;
                    }
                    else
                    {
                        throw new InvalidOperationException("Cannot resolve specified range");
                    }
                }
            }
        }
        Workbook GetCachedWorkbook(string fileName, IGridDataOpenContext openContext)
        {
            // Do we even have one in the cache?
            Workbook workbook = openContext.GetNamedValue(fileName) as Workbook;
            if (workbook != null)
            {
                // When did we store it?
                var timeStampObject = openContext.GetNamedValue(fileName + __TimeStamp);
                if (timeStampObject is DateTime)
                {
                    var timeStamp = (DateTime)timeStampObject;
                    // Check the timestamp on the file to see if we're still current.
                    var fi = new System.IO.FileInfo(fileName);
                    if (fi.LastWriteTime < timeStamp)
                        return workbook; // Cached one is still good.
                }
            }
            return null;
        }

        #region IGridDataRecords Members

        Type TypeFromCellType(Cell c, CellValueType type)
        {
            switch (type)
            {
                case CellValueType.Text:
                    return typeof(string);
                case CellValueType.Numeric:
                    if (!c.IsDisplayedAsDateTime)
                        return typeof(double);
                    else
                        return typeof(DateTime);
                case CellValueType.DateTime:
                    return typeof(DateTime);
                case CellValueType.Boolean:
                    return typeof(bool);
                case CellValueType.None:
                case CellValueType.Error:
                case CellValueType.Unknown:
                    break;
            }

            return typeof(string);
        }

        List<GridDataColumnInfo> _columns;
        public IEnumerable<GridDataColumnInfo> Columns
        {
            get
            {
                if (_columns == null)
                {
                    //check if number of rows and columns are 0 instead of checking for the object array _data
                    if (_lastColumnIndex == 0 && _lastRowIndex == 0)
                        return null;

                    _columns = new List<GridDataColumnInfo>();

                    for (int i = _startColumnIndex; i < _lastColumnIndex; i++)
                    {
                        string colname = String.Format("Col{0}", i);
                        var cell = _sheet.Rows[_startRowIndex][i];

                        if (cell != null)
                        {
                            var type = cell.Value.Type;
                            switch (type)
                            {
                                case CellValueType.Text:
                                case CellValueType.Numeric:
                                case CellValueType.Boolean:
                                    colname = ExcelUtils.GetCellValue(cell, type);
                                    break;
                            }
                        }


                        GridDataColumnInfo info = new GridDataColumnInfo() { Name = colname, Type = typeof(string) };

                        for (int j = _startRowIndex + 1; j < _lastRowIndex; j++)
                        {
                            Cell c1 = _sheet.Rows[j][i];
                            if (c1 != null)
                            {
                                info.Type = TypeFromCellType(c1, c1.Value.Type);
                                break;
                            }
                        }

                        _columns.Add(info);
                    }
                }

                return _columns;
            }
        }

        #endregion

        class ExcelEnumerator : IEnumerator<IGridDataRecord>
        {
            Worksheet _sheet;
            ExcelGridDataRecord _current;

            int _startIndex;
            int _index;
            int _columnIndex;
            int _numRows;
            public ExcelEnumerator(Worksheet sh, int rowIndex, int colIndex, int numRows)
            {
                _sheet = sh;
                _startIndex = rowIndex;
                _index = rowIndex;
                _columnIndex = colIndex;
                _numRows = numRows;
            }

            #region IEnumerator<IGridDataRecord> Members

            public IGridDataRecord Current
            {
                get { return _current; }
            }

            #endregion

            #region IDisposable Members

            public void Dispose()
            {
            }

            #endregion

            #region IEnumerator Members

            object System.Collections.IEnumerator.Current
            {
                get { return _current; }
            }

            public bool MoveNext()
            {
                if (_index < _numRows)
                {
                    _current = new ExcelGridDataRecord(_index, _sheet, _columnIndex);
                    _index++;
                    return true;
                }

                return false;
            }

            public void Reset()
            {
                _index = _startIndex;
            }

            #endregion
        }

        #region IEnumerable<IGridDataRecord> Members

        public IEnumerator<IGridDataRecord> GetEnumerator()
        {
            return new ExcelEnumerator(_sheet, _startRowIndex + 1, _startColumnIndex, _lastRowIndex);
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ExcelEnumerator(_sheet, _startRowIndex + 1, _startColumnIndex, _lastRowIndex);
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
        }

        #endregion
    }

    class ExcelUtils
    {
        public static string GetCellValue(Cell cell, CellValueType type)
        {
            switch (type)
            {
                case CellValueType.Text:
                    return cell.Value.TextValue.ToString();
                case CellValueType.Numeric:
                    return GetDateTimeOrNumericValueAsString(cell);
                case CellValueType.Boolean:
                    return cell.Value.BooleanValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case CellValueType.None:
                    return null;
                case CellValueType.Error:
                    return cell.Value.ErrorValue.ToString();
                case CellValueType.DateTime:
                    return GetDateTimeOrNumericValueAsString(cell);
                case CellValueType.Unknown:
                    return "Unknown";
            }

            return null;
        }

        private static string GetDateTimeOrNumericValueAsString(Cell cell)
        {
            if (cell.IsDisplayedAsDateTime || cell.Value.IsDateTime)
            {
                DateTime dt = cell.Value.DateTimeValue;
                if (dt.Millisecond >= 995)
                {
                    // Excel stores things as Days from Jan 1, 1900. This can (apparently) result in some values like
                    // 1/7/2016 4:29:59.999 when what was in excel was shown as 1/7/2016 4:30:00, so....
                    // If we are very, very close to the next second, so we'll go to the next second, since 
                    //  the ToString() will simply strip off any sub-second values.
                    dt = dt.AddSeconds(1.0);
                }
                return dt.ToString(); // Simio will first try to parse dates in the current culture
            }
            else
            {
                return cell.Value.NumericValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
        }
    }


    class ExcelGridDataRecord : IGridDataRecord
    {
        int _cIndex;
        int _index;
        Worksheet _sheet;

        public ExcelGridDataRecord(int index, Worksheet sheet, int columnIndex)
        {
            _index = index;
            _sheet = sheet;
            _cIndex = columnIndex;
        }

        #region IGridDataRecord Members

        public string this[int index]
        {
            get
            {
                Cell cell = _sheet.Rows[_index][index + _cIndex];


                if (cell != null)
                {


                    return ExcelUtils.GetCellValue(cell, cell.Value.Type);
                }

                return null;
            }
        }

        #endregion
    }
}


