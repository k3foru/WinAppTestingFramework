//===============================================================================
// Ellis WinApp Testing Framework Library
// By Kiran Kumar
//===============================================================================

using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;

namespace Ellis.WinApp.Testing.Framework.Actions
{
    public class TableActions : AppContext
    {
        public static WinCell SelectCellFromTable(UITestControl windowProperties, string tableName, string rowName, string cellName)
        {
            var table = (WinTable)Actions.GetWindowChild(windowProperties, tableName);
            var row = table.Container.SearchFor<WinRow>(new { Name = rowName });
            var cell = row.Container.SearchFor<WinCell>(new { Name = cellName });
            cell.SetFocus();
            return cell;
        }

        public static WinRow SelectRowFromTable(UITestControl windowProperties, string tableName, string rowName)
        {
            var table = (WinTable)Actions.GetWindowChild(windowProperties, tableName);
            var row = table.Container.SearchFor<WinRow>(new { Name = rowName });
            row.SetFocus();
            return row;
        }

        public static bool OpenRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue)
        {
            var winTable = (WinTable)Actions.GetWindowChild(windowInstence, tableControlName);
            foreach (var uiTestControl in winTable.Rows)
            {
                uiTestControl.SetFocus();
                var winCell = CodedUIExtension.SearchFor<WinCell>(winTable.Container, new
                {
                    Name = columnName
                });
                if (winCell.GetProperty("Value").ToString() != columnValue) continue;
                winCell.SetFocus();
                Mouse.DoubleClick(winCell);
                return true;
            }
            return false;
        }

        public static bool SelectRecordFromTable(UITestControl windowInstence, string tableControlName, string columnName, string columnValue)
        {
            var tableName = Actions.GetWindowChild(windowInstence, tableControlName);
            var table = (WinTable)tableName;

            foreach (var rowC in table.Rows)
            {
                rowC.SetFocus();
                var rowHeader = table.Container.SearchFor<WinCell>(new { Name = columnName });
                var callValue = rowHeader.GetProperty("Value").ToString();

                if (callValue != columnValue) continue;
                rowHeader.SetFocus();
                Mouse.Click(rowHeader);
                return true;
            }
            return false;
        }
    }
}
