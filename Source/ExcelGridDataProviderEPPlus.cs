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

        /// <summary>
        /// ??? This is being called for each cell ???
        /// </summary>
        /// <param name="dataSettings"></param>
        /// <returns></returns>
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
        /// <summary>
        /// A rectangular range of excel cells that is given a name
        /// </summary>
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

        // An arbitrary default is needed, so 10 columns of two rows is chosen.
        const string DEFAULT_SPECIFIC_RANGE = "A1:B10";
        string _specificRange = DEFAULT_SPECIFIC_RANGE;

        /// <summary>
        /// An excel range, of the format {upper-left}:{lower-right}. For example: A1:F20
        /// </summary>
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

        /// <summary>
        /// The rectangle of excel cells is an entire worksheet
        /// </summary>
        public bool IsWorksheetRange
        {
            get { return RangeType == RType.WORKSHEET; }
            set { RangeType = RType.WORKSHEET; }
        }

        /// <summary>
        /// The range is specified by an address, such as "B08:H23"
        /// </summary>
        public bool IsSpecificRange
        {
            get { return RangeType == RType.SPECIFIC_RANGE; }
            set { RangeType = RType.SPECIFIC_RANGE; }
        }

        /// <summary>
        /// An Excel range that is given a name.
        /// </summary>
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
        /// Constructor. Makes sure _workbook, _sheet, and _*Index's are set,
        /// depending on whether we want sheet, namedRange, or specific range.
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

            if ( _workbook == null )
                throw new InvalidDataException($"Cannot find Worksheet for File={settings.FileName}");

            _startRowIndex = 0;
            _startColumnIndex = 0;

            _sheet = _workbook.Worksheets[settings.Worksheet];
            if (_sheet == null)
                throw new InvalidDataException($"Setting Sheet={settings.Worksheet} cannot be fond.");

                if (_package != null)
            {
                if (settings.IsWorksheetRange)
                {
                    ExcelAddressBase usedRange = _sheet.Dimension;
                    _startRowIndex = usedRange.Start.Row;
                    _startColumnIndex = usedRange.Start.Column;
                    _lastRowIndex = usedRange.End.Row;
                    _lastColumnIndex = usedRange.End.Column;
                }
                else if ( settings.IsNamedRange ) // Looking for an Excel range
                {
                    ExcelNamedRange namedRange = _workbook.Names[settings.NamedRange];
                    if ( namedRange != null )
                    {
                        // Change the sheet to where the name range was found
                        _sheet = namedRange.Worksheet;

                        string addr = namedRange.Address;
                        _lastRowIndex = namedRange.End.Row;
                        _lastColumnIndex = namedRange.End.Column;
                        _startRowIndex = namedRange.Start.Row;
                        _startColumnIndex = namedRange.Start.Column;

                    }
                }
                else if ( settings.IsSpecificRange ) // looking for address like "A3:E20"
                {
                    var addr = new ExcelAddress(settings.SpecificRange);
                    if (addr != null)
                    {
                        _lastRowIndex = addr.End.Row;
                        _lastColumnIndex = addr.End.Column;
                        _startRowIndex = addr.Start.Row;
                        _startColumnIndex = addr.Start.Column;
                    }
                }
                else
                {
                    throw new InvalidOperationException("Setting is not Worksheet, Named, or Specific");
                } // if named or specific range
            } // package exists
        }

        /// <summary>
        /// To avoid always opening files (which - when large - can be a lengthy operation) this
        /// method allows us to get the in-memory cached version instead.
        /// Null is returned if there is on in-cache valid version.
        /// Reasons for being invalid: 
        /// 1. Disk version is newer than cached version.
        /// 2. Cannot find cached version.
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
        /// <summary>
        /// This method expects:  _sheet, _columns, _startRowIndex, _lastRowIndex.
        /// </summary>
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
                            colname = ExcelUtils.GetCellValue(cell );
                        }

                        GridDataColumnInfo info = new GridDataColumnInfo() { Name = colname, Type = typeof(string) };

                        // Now scan the rows below the header, looking for the first non-null value
                        // Infer the type and use this as the column type.
                        for (int rr = _startRowIndex + 1; rr <= _lastRowIndex; rr++)
                        {
                            cell = _sheet.Cells[rr, cc];
                            var vv = cell?.Value; // Get cell value according to OpenOfficeXml
                            if ( vv.GetType().FullName == "System.Double")
                            {
                                DateTime dt = DateTime.MinValue;
                                if ( DateTime.TryParse(cell.Text, out dt))
                                {
                                    cell.Value = dt;
                                }
                            }

                            if (vv != null)
                            {
                                info.Type = vv.GetType(); //  Store the System type (e.g. System.Double) 
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

    /// <summary>
    /// Excel recognizes four kinds of information:
    /// Logical, Numeric, Text, and Error.
    /// DateTime values are numeric (they are a float of days since 1 Jan 1900, 
    /// so you use the fractional component to get the time)
    /// If you want to get a Simio date you have to convert.
    /// </summary>
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
                    return ExcelUtils.GetCellValue(cell); //??todo  // cell.Value.GetType());

                return null;
            }
        }

        #endregion
    }
}


