using SimioAPI.Extensions;
using System;
using System.IO;
using System.Collections.Generic;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace ExcelGridDataProviderEPPlus
{
    /// <summary>
    /// This version of the ExcelGridDataProvider employs OfficeOpenXml and EPPlus.
    /// Note: indices in EPPlus are 1 based!
    /// </summary>
    public class ExcelGridDataProviderEPPlus : IGridDataProvider, IGridDataProviderWithFiles
    {
        #region IGridDataProvider Members

        public string Name
        {
            get { return "ExcelEPPlus"; }
        }

        public string Description
        {
            get { return "Reads data from an Excel spreadsheet (using EPPlus)"; }
        }

        public System.Drawing.Image Icon
        {
            get { return Properties.Resources.Icon; }
        }

        public Guid UniqueID
        {
            get { return MY_ID; }
        }
        static readonly Guid MY_ID = new Guid("9E0407C8-8BC4-4E87-9234-5EDFCBBB0CAE"); // Changed GUID for EPPlus version

        /// <summary>
        /// Called when a Binding is created.
        /// </summary>
        /// <param name="existingSettings"></param>
        /// <returns></returns>
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
            var summary = $"Bound to Excel: {fileName}" ?? "[No file name]";

            if (thesettings.IsNamedRange && thesettings.NamedRange != null)
                summary += $", Named Range: {thesettings.NamedRange}";
            else if (thesettings.IsSpecificRange && thesettings.Worksheet != null && thesettings.SpecificRange != null)
                summary += $", Worksheet: {thesettings.Worksheet}, Range: {thesettings.SpecificRange}";
            else if (thesettings.IsWorksheetRange && thesettings.Worksheet != null)
                summary += $", Worksheet: { thesettings.Worksheet}";

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

    /// <summary>
    /// See http://support.microsoft.com/kb/320369
    /// </summary>
    class ExcelCultureFence : IDisposable
    {
        System.Globalization.CultureInfo _cInfo;

        public ExcelCultureFence()
        {
            _cInfo = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        #region IDisposable Members

        public void Dispose()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = _cInfo;
        }

        #endregion
    }

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

        /// <summary>
        /// Get the lists of worksheets and named ranges
        /// </summary>
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

                // Load a workbook
                ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(_fileName));
                ExcelWorkbook workbook = package.Workbook;

                // Access a collection of worksheets. EPPlus uses 1-based index
                for (int ii = 1; ii <= workbook.Worksheets.Count; ii++)
                {
                    _worksheets.Add(workbook.Worksheets[ii].Name);
                }
                
                // Create a collection of named ranges (if there are any) EPPlus uses 0-based index
                for (int ii = 0; ii < workbook.Names.Count; ii++)
                {
                    if ( workbook.Names[ii] != null )
                    {
                        //filter formula                      
                        _namedRanges.Add(workbook.Names[ii].Name);
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

        /// <summary>
        /// Deserialize the binary 'settings' object
        /// </summary>
        /// <param name="settings"></param>
        /// <returns></returns>
        public static ExcelGridDataSettings FromBytes(byte[] settings)
        {
            if (settings == null)
                return null;

            System.IO.MemoryStream memstream = new System.IO.MemoryStream(settings);
            System.Runtime.Serialization.Formatters.Binary.BinaryFormatter fmt = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();

            ExcelGridDataSettings excelsettings = (ExcelGridDataSettings)fmt.Deserialize(memstream);

            return excelsettings;
        }

        /// <summary>
        /// Serialize the binary 'settings' object.
        /// </summary>
        /// <returns></returns>
        public byte[] ToBytes()
        {
            System.IO.MemoryStream memstream = new System.IO.MemoryStream();
            System.Runtime.Serialization.Formatters.Binary.BinaryFormatter fmt = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();

            fmt.Serialize(memstream, this);

            return memstream.ToArray();
        }
    }

    /// <summary>
    /// Implements Simio's IGridDataRecords
    /// </summary>
    class ExcelGridDataRecords : IGridDataRecords
    {
        ExcelPackage _package;
        ExcelWorksheet _sheet;
        ExcelWorkbook _workbook;

        int _startRowIndex, _startColumnIndex; // 0-based for legacy reasons
        int _lastRowIndex, _lastColumnIndex;    // 0-based for legacy reasons
        const string __TimeStamp = "__TimeStamp";

        /// <summary>
        /// ctor.
        /// </summary>
        /// <param name="settings"></param>
        /// <param name="openContext"></param>
        public ExcelGridDataRecords(ExcelGridDataSettings settings, IGridDataOpenContext openContext)
        {
            // Do we have a cached package?
            _package = GetCachedPackage(settings.FileName, openContext);
            if (_package == null)
            {
                // Load, or re-load from the file, which is a 'package' file.
                _package = new ExcelPackage(new System.IO.FileInfo(settings.FileName));
                _workbook = _package.Workbook; 

                // Cache the in-memory package, as well as when we loaded it.
                openContext.SetNamedValue(settings.FileName, _package);
                openContext.SetNamedValue(settings.FileName + __TimeStamp, DateTime.Now);
            }
            else
            {
                _workbook = _package.Workbook;
            }

            if (_package != null)
            {
                if (settings.IsWorksheetRange)
                {
                    _startRowIndex = 0;
                    _startColumnIndex = 0;
                    _sheet = _workbook.Worksheets[settings.Worksheet];

                    if (_sheet != null)
                    {
                        ExcelAddressBase usedRange = _sheet.Dimension;
                        _startRowIndex = usedRange.Start.Row;
                        _startColumnIndex = usedRange.Start.Column;
                        _lastRowIndex = usedRange.End.Row;
                        _lastColumnIndex = usedRange.End.Column;
                    }
                }
                else if (settings.IsNamedRange || settings.IsSpecificRange)
                {
                    _sheet = _workbook.Worksheets[1];

                    if (settings.IsNamedRange)
                    {
                        ExcelNamedRange namedRange = _workbook.Names[settings.NamedRange];
                        _sheet = namedRange.Worksheet;  //?? add error check

                        //DefinedName definedName =   DefinedNames.GetDefinedName(settings.NamedRange);

                        //aref = definedName.Range;
                        //_sheet = aref.Worksheet;

                    }
                    else
                    {
                        _sheet = _workbook.Worksheets[settings.Worksheet];
                    }

                    if ( _sheet != null)
                    {
                        ExcelAddressBase addr = _sheet.Dimension;

                        _lastRowIndex = addr.End.Row;
                        _lastColumnIndex = addr.End.Column;
                        _startRowIndex = addr.Start.Row;
                        _startColumnIndex = addr.Start.Column;
                    }
                    else
                    {
                        throw new InvalidOperationException("Cannot resolve specified range");
                    }
                } // if named or specific range
            } // package exists
        }

        /// <summary>
        /// To avoid always opening files (which when large can be a lengthy operation) this
        /// method allows us to get the in-memory cached version instead.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="openContext"></param>
        /// <returns></returns>
        ExcelPackage GetCachedPackage(string fileName, IGridDataOpenContext openContext)
        {
            // Do we even have one in the cache?
            ExcelPackage package = openContext.GetNamedValue(fileName) as ExcelPackage;
            if (package != null)
            {
                // When did we store it?
                var timeStampObject = openContext.GetNamedValue(fileName + __TimeStamp);
                if (timeStampObject is DateTime)
                {
                    var timeStamp = (DateTime)timeStampObject;
                    // Check the timestamp on the file to see if we're still current.
                    var fi = new System.IO.FileInfo(fileName);
                    if (fi.LastWriteTime < timeStamp)
                        return package; // Cached one is still good.
                }
            }
            return null;
        }

        #region IGridDataRecords Members

        /// <summary>
        /// Mapping between Simio and Excel data types.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        Type TypeFromCellType(ExcelCell cell, CellValueType type)
        {
            switch (type)
            {
                case CellValueType.Text:
                    return typeof(string);
                case CellValueType.Numeric:
                    //if (!cell.IsDisplayedAsDateTime)
                        return typeof(double);
                    //else
                    //    return typeof(DateTime);
                case CellValueType.DateTime:
                    return typeof(DateTime);
                case CellValueType.Boolean:
                    return typeof(bool);
                case CellValueType.None:
                case CellValueType.Error:
                case CellValueType.Unknown:
                    break;

                default:
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

                    for (int cc = _startColumnIndex; cc <= _lastColumnIndex; cc++)
                    {
                        string colname = $"Col{cc}";
                        var cell = _sheet.Cells[_startRowIndex, cc]; // Get cell on first row

                        // The first row contains the column names
                        if (cell != null)
                        {
                            var type = cell.Value.GetType();

                            if ( type.Name != "String")
                            {
                                string xx = type.Name;
                            }

                            colname = ExcelUtils.GetCellValue(cell, CellValueType.Text).ToString();

                            // If OrderDate, then it comes in as a Double.

                            ////switch (type) // todo: check these out and build a unit test for it.
                            ////{
                            ////    case CellValueType.Text:
                            ////    case CellValueType.Numeric:
                            ////    case CellValueType.Boolean:
                            ////        colname = ExcelUtils.GetCellValue(cell, type);
                            ////        break;
                            ////}
                        }

                        GridDataColumnInfo info = new GridDataColumnInfo() { Name = colname, Type = typeof(string) };

                        // Look for the first non-null value and use that as the column type.
                        for (int rr = _startRowIndex + 1; rr <= _lastRowIndex; rr++)
                        {
                            var vv = _sheet.Cells[rr, cc]?.Value;
                            if (vv != null)
                            {
                                info.Type = vv.GetType(); // TypeFromCellType(c1, c1.Value.Type);
                                if ( info.Type.Name != "String")
                                {
                                    string xx = "";
                                }
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

        /// <summary>
        /// Simio API for GridDataRecord (which is 0 based)
        /// So store everything 0-based and convert when we call EPPlus methods.
        /// </summary>
        class ExcelEnumerator : IEnumerator<IGridDataRecord>
        {
            ExcelWorksheet _sheet;
            ExcelGridDataRecord _current;

            int _startIndex;    // 0 based
            int _rowIndex;      // 0 based
            int _columnIndex;   // 0 based
            int _numRows;

            public ExcelEnumerator(ExcelWorksheet sh, int rowIndex, int colIndex, int numRows)
            {
                _sheet = sh;
                _startIndex = rowIndex;
                _rowIndex = rowIndex;
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
                if (_rowIndex < _numRows)
                {
                    _current = new ExcelGridDataRecord(_rowIndex, _sheet, _columnIndex);
                    _rowIndex++;
                    return true;
                }

                return false;
            }

            public void Reset()
            {
                _rowIndex = _startIndex;
            }

            #endregion
        }

        #region IEnumerable<IGridDataRecord> Members

        public IEnumerator<IGridDataRecord> GetEnumerator()
        {
            return new ExcelEnumerator(_sheet, _startRowIndex, _startColumnIndex, _lastRowIndex);
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return new ExcelEnumerator(_sheet, _startRowIndex, _startColumnIndex, _lastRowIndex);
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
        }

        #endregion
    }

    public enum CellValueType
    {
        Text = 0,
        Numeric = 1,
        Boolean = 2,
        None = 3,
        Error = 4,
        DateTime = 5,
        Unknown = 9
    }

    class ExcelUtils
    {
        /// <summary>
        /// Get the cell value as a string
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetCellValue(ExcelRange cell, CellValueType type)
        {
            switch (type)
            {
                case CellValueType.Text:
                    return cell.Value.ToString(); 
                case CellValueType.Numeric:
                    return GetDateTimeOrNumericValueAsString(cell);
                case CellValueType.Boolean:
                    return cell.GetValue<bool>().ToString(System.Globalization.CultureInfo.InvariantCulture);
                case CellValueType.None:
                    return null;
                case CellValueType.Error:
                    return cell.Value.ToString();
                case CellValueType.DateTime:
                    return GetDateTimeOrNumericValueAsString(cell);
                case CellValueType.Unknown:
                    return "Unknown";
            }

            return null;
        }

        /// <summary>
        /// If the type of the Cell's Value is DateTime, then return
        /// the value as DateTime, with adjustments to round it to seconds
        /// when the milliseconds is 995 or greater.
        /// If the type is Numeric, the return the Cell's value.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        private static string GetDateTimeOrNumericValueAsString(ExcelRange cell)
        {

            if (cell.Value is DateTime)
            {
                DateTime dt = (DateTime) cell.Value;
                if (dt.Millisecond >= 995)
                {
                    // Excel stores things as Days from Jan 1, 1900 (or Jan 1, 1904 for the Mac)
                    // This can (apparently) result in some values like
                    // 1/7/2016 4:29:59.999 when what was in excel was shown as 1/7/2016 4:30:00, so....
                    // If we are very, very close to the next second, so we'll go to the next second, since 
                    //  the ToString() will simply strip off any sub-second values.
                    dt = dt.AddSeconds(1.0);
                }
                return dt.ToString(); // Simio will first try to parse dates in the current culture
            }
            else if (cell.Value is Decimal dd)
            {
                return dd.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            else
            {
                return cell.Value.ToString();
            }
        }

        /// <summary>
        /// Get a cell as text and try and parse it as a decimal
        /// Return false if the value is null or isn't a decimal, in which case the dd argument is untouched.
        /// Return true if a legitimate double (dd) is found and set.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="dd"></param>
        /// <returns></returns>
        public Decimal? GetCellAsDecimal(ExcelRange cell)
        {
            if (cell?.Value == null)
                return null;

            if (Decimal.TryParse(cell.Text, out decimal newValue))
                return newValue;
            else
                return null;

        }

    } // class

    /// <summary>
    /// The excel record for the Simio API (which is 0 based)
    /// So everything internally is stored 0 based, and we convert to 1-based when
    /// calling the EPPlus methods.
    /// </summary>
    class ExcelGridDataRecord : IGridDataRecord
    {
        int _cIndex; // column index (0 based)
        int _rIndex; // row index (0 based)
        ExcelWorksheet _sheet; // selected sheet


        public ExcelGridDataRecord(int index, ExcelWorksheet sheet, int columnIndex)
        {
            _rIndex = index;
            _sheet = sheet;
            _cIndex = columnIndex;
        }

        #region IGridDataRecord Members

        public string this[int index]
        {
            get
            {
                ExcelRange cell = _sheet.Cells[ _rIndex + 1, _cIndex + index ];

                if (cell != null)
                    return ExcelUtils.GetCellValue(cell, CellValueType.Text); //??todo  // cell.Value.GetType());

                return null;
            }
        }

        #endregion
    }
}


