namespace Ninjanaut.IO
{
    public class BinaryExcelReaderOptions
    {
        /// <summary>
        /// Default value is null. I recommend setting this value so that you don't accidentally load empty columns.
        /// </summary>
        public int? MaxColumns { get; set; }

        /// <summary>
        /// Default value is 0. Keep in mind that OLE DB driver does not take into account blank rows. 
        /// For example, if you have 4 additional non-header rows from top and two of them are blank, 
        /// the header row index is 2. Warning: if the row contains formatting, it is not considered blank.
        /// </summary>
        public int? HeaderRowIndex { get; set; }


        /// <summary>
        /// Default value is true. If set to false and the row does not contains anything (even formatting), 
        /// then the row will not be loaded anyway.
        /// </summary>
        public bool RemoveEmptyRows { get; set; }


        /// <summary>
        /// Default value is true. If set to false, HeaderRowIndex property is ignored.
        /// </summary>
        public bool HeaderExists { get; set; }

        public BinaryExcelReaderOptions()
        {
            MaxColumns = null;
            HeaderRowIndex = 0;
            RemoveEmptyRows = true;
            HeaderExists = true;
        }
    }
}
