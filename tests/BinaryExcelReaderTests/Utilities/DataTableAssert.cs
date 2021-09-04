using System.Data;
using Xunit;

namespace BinaryExcelReaderTests.Utilities
{
    public static class DataTableAssert
    {
        public static void DataTables(DataTable actual, DataTable expected)
        {
            Columns(actual, expected);
            Rows(actual, expected);
        }

        public static void Rows(DataTable actual, DataTable expected)
        {
            for (var rowIndex = 0; rowIndex < expected.Rows.Count; rowIndex++)
            {
                for (var colIndex = 0; colIndex < expected.Rows[rowIndex].ItemArray.Length; colIndex++)
                {
                    var expectedValue = expected.Rows[rowIndex].ItemArray[colIndex]?.ToString();
                    var excelValue = actual.Rows[rowIndex][colIndex].ToString();
                    Assert.Equal(expectedValue, excelValue);
                }
            }
        }

        public static void Columns(DataTable actual, DataTable expected)
        {
            for (var headerIndex = 0; headerIndex < expected.Columns.Count; headerIndex++)
            {
                var expectedValue = expected.Columns[headerIndex]?.ToString();
                var excelValue = actual.Columns[headerIndex].ToString();

                Assert.Equal(expectedValue, excelValue, ignoreCase: true);
            }
        }

        public static void ColumnsWithDuplication(DataTable actual, DataTable expected, int duplicatedColumnNumber)
        {
            for (var headerIndex = 0; headerIndex < expected.Columns.Count; headerIndex++)
            {
                var expectedValue = expected.Columns[headerIndex]?.ToString();
                var excelValue = actual.Columns[headerIndex].ToString();

                if (headerIndex == duplicatedColumnNumber - 1)
                {
                    Assert.NotNull(expectedValue);
                    Assert.NotEmpty(expectedValue);
                    Assert.True(excelValue.Length == expectedValue.Length);
                    continue;
                }

                Assert.Equal(expectedValue, excelValue, ignoreCase: true);
            }
        }
    }
}
