using BinaryExcelReaderTests.Utilities;
using Ninjanaut.IO;
using System;
using System.Data;
using System.Globalization;
using Xunit;

namespace BinaryBinaryExcelReaderTests
{
    public class UnitTests
    {
        [Fact]
        public void Load_excel_with_empty_header_rows_from_top()
        {
            // Act
            // OleDbDataReader automatically ignore blank rows from the top!
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\EmptyRowsFromTop.xlsb", 
                "CustomSheetName");

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_invalid_header_rows_from_top()
        {
            // Act
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\InvalidRowsFromTop.xlsb", 
                "CustomSheetName", new() { HeaderRowIndex = 2 });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_and_remove_empty_rows()
        {
            // Act
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\EmptyRows.xlsb", 
                "CustomSheetName", new() { RemoveEmptyRows = true });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_empty_rows()
        {
            // Act
            // Each empty row contains formatting so that excel also knows that it is the row to include.
            // If there is really nothing in the row, the row is still ignored.
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\EmptyRows.xlsb", 
                "CustomSheetName", new() { RemoveEmptyRows = false });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "", "", "" });
            dt.AddRow(new object[] { "", "", "" });

            Assert.NotNull(datatable);
            Assert.Equal(9, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_max_columns_option()
        {
            // Act
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\AdditionalColumns.xlsb", 
                "CustomSheetName", new() { MaxColumns = 3 });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_duplicated_columns()
        {
            // Act
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\DuplicatedColumns.xlsb", 
                "CustomSheetName");

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("B1"),
                new DataColumn("C"),
                new DataColumn("C1"),
                new DataColumn("C11")
            });

            dt.AddRow(new object[] { "1", "2", "3", "4", "5", "6" });
            dt.AddRow(new object[] { "1", "2", "3", "4", "5", "6" });
            dt.AddRow(new object[] { "1", "2", "3", "4", "5", "6" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_excel_via_sheet_name()
        {
            // Act
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\CustomSheetName.xlsb", 
                "Custom Sheet Name");

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C")
            });

            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });
            dt.AddRow(new object[] { "1", "2", "3" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_max_column_option_that_is_larger_than_header()
        {
            // Act
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\AdditionalColumns.xlsb",
                "CustomSheetName", new() { MaxColumns = 20,  });

            // Assert
            var dt = new DataTable();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("A"),
                new DataColumn("B"),
                new DataColumn("C"),
                new DataColumn("D"),
                new DataColumn("E")
            });

            dt.AddRow(new object[] { "1", "2", "3", "4", "5" });
            dt.AddRow(new object[] { "1", "2", "3", "4", "5" });
            dt.AddRow(new object[] { "1", "2", "3", "4", "5" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.DataTables(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_known_edge_cases()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xlsb", 
                "CustomSheetName", new() { HeaderRowIndex = 1 });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToShortDateString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("F1"),
                new DataColumn("B"),
                new DataColumn("B1"),
                new DataColumn("C"),
                new DataColumn("F5"),
                new DataColumn("D"),
                new DataColumn("TRUE"),
                new DataColumn("1"),
                new DataColumn(dateValue),
                new DataColumn("12#56"),
            });

            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "False", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_excel_with_known_edge_cases_without_header()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xlsb",
                "CustomSheetName", new() { HeaderExists = false });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToShortDateString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("F1"),
                new DataColumn("F2"),
                new DataColumn("F3"),
                new DataColumn("F4"),
                new DataColumn("F5"),
                new DataColumn("F6"),
                new DataColumn("F7"),
                new DataColumn("F8"),
                new DataColumn("F9"),
                new DataColumn("F10"),
            });

            dt.AddRow(new object[] { "", "B", "B", "C", "", "D", "TRUE", "1", "28/08/2021", "12.56" });
            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "FALSE", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(4, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_xlsx_excel_with_known_edge_cases()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xlsx", 
                "CustomSheetName", new() { HeaderRowIndex = 1 });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToShortDateString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("F1"),
                new DataColumn("B"),
                new DataColumn("B1"),
                new DataColumn("C"),
                new DataColumn("F5"),
                new DataColumn("D"),
                new DataColumn("TRUE"),
                new DataColumn("1"),
                new DataColumn(dateValue),
                new DataColumn("12#56"),
            });

            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "False", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_xls_excel_with_known_edge_cases()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xls",
                "CustomSheetName", new() { HeaderRowIndex = 1 });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToShortDateString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("F1"),
                new DataColumn("B"),
                new DataColumn("B1"),
                new DataColumn("C"),
                new DataColumn("F5"),
                new DataColumn("D"),
                new DataColumn("TRUE"),
                new DataColumn("1"),
                new DataColumn(dateValue),
                new DataColumn("12#56"),
            });

            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "False", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }

        [Fact]
        public void Load_xlsm_excel_with_known_edge_cases()
        {
            // Act
            // Excel has first sheet hidden.
            var datatable = BinaryExcelReader.ToDataTable(@"TestData\KnownEdgeCases.xlsm", 
                "CustomSheetName", new() { HeaderRowIndex = 1 });

            // Assert
            var dt = new DataTable();

            var dateValue = DateTime.ParseExact("28/08/2021", "dd/MM/yyyy", CultureInfo.InvariantCulture).ToShortDateString();

            dt.Columns.AddRange(new[]
            {
                new DataColumn("F1"),
                new DataColumn("B"),
                new DataColumn("B1"),
                new DataColumn("C"),
                new DataColumn("F5"),
                new DataColumn("D"),
                new DataColumn("TRUE"),
                new DataColumn("1"),
                new DataColumn(dateValue),
                new DataColumn("12#56"),
            });

            dt.AddRow(new object[] { "1", "Value of B4", "3", "", "5", "", "7", "", "", "1325,48" });
            dt.AddRow(new object[] { "", "", "3", "", "4", "", "", "False", "", "" });
            dt.AddRow(new object[] { "1", "7", "3", "4", dateValue, "6", "7", "", "1234.4895", "" });

            Assert.NotNull(datatable);
            Assert.Equal(3, datatable.Rows.Count);
            DataTableAssert.ColumnsWithDuplication(datatable, dt, duplicatedColumnNumber: 3);
            DataTableAssert.Rows(datatable, dt);
        }
    }
}
