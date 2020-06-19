using System;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using WA_Test_V5.Interface.TreeView;

namespace WA_Test_V5.GetData.Excel
{
    public class GetExcelData
    {
        private ExcelPackage pack;
        private ExcelWorksheet sheet;
        private string path;
        private int increment = 1; //TODO: make concurrent safety

        public GetExcelData(string filePath)
        {
            path = filePath;
            FileInfo fInfo = new FileInfo(path);
            if (fInfo.Exists != true) throw new Exception();
            pack = new ExcelPackage(fInfo);
            sheet = pack.Workbook.Worksheets.First(); // TODO: unsafe
        }
        public List<TreeViewElements> GetSample()
        {
            var sheets = pack.Workbook.Worksheets;
            var dataSheet = sheets.First();
            var numberOfRows = dataSheet.Dimension.End.Row;
            var numberOfCols = dataSheet.Dimension.End.Column;
            ExcelRange Cells = dataSheet.Cells;
            var _SampleTreeView = new List<TreeViewElements>();
            for (int rowIterator = 2; rowIterator <= numberOfRows; rowIterator++)
            {
                var e = new TreeViewElements()
                {
                    ID = Cells[rowIterator, 1].Value.ToString(),
                    Parent_ID = Cells[rowIterator, 2].Value.ToString(),
                    Name = Cells[rowIterator, 3].Value.ToString(),
                    CID = Convert.ToInt32(Cells[rowIterator, 4].Value),
                };
                _SampleTreeView.Add(e);
            }
            return _SampleTreeView;
        }

        public List<TreeViewElements> GetData()
        {
            var sheets = pack.Workbook.Worksheets;

            var numberOfRows = sheet.Dimension.End.Row;
            var numberOfCols = sheet.Dimension.End.Column;
            ExcelRange Cells = sheet.Cells;




            var dictionary = new SortedDictionary<string, object>();

            for (int i = 2; i <= sheet.Dimension.Rows; i++)
            {
                AddRow(i, 1, dictionary);
            }




            var _SampleTreeView = new List<TreeViewElements>();
            Read(dictionary, _SampleTreeView, 0);








            return _SampleTreeView;
        }

        private void Read(
            //Dictionary<string, object> dictionary,
            object dictionary,
            List<TreeViewElements> treeViewElements,
            int inherritanceLevel,
            string parentId = "0")
        {
            if (dictionary is SortedDictionary<string, List<string>>)
            {
                var d2 = (SortedDictionary<string, List<string>>)dictionary;

                foreach (var key in d2.Keys)
                {
                    foreach (var key2 in d2[key])
                    {
                        if (treeViewElements.Any(
                            x => x.CID == int.Parse(key2) &&
                            x.Name == key &&
                            x.Parent_ID == parentId))
                        {
                            continue;
                        }

                        treeViewElements.Add(new TreeViewElements()
                        {
                            ID = increment.ToString(),
                            Name = key,
                            Parent_ID = parentId,
                            CID = int.Parse(key2) // unsafe
                        });

                        increment++;
                    }
                }
                return;
            }
            if (dictionary is SortedDictionary<string, object>)
            {
                var d1 = (SortedDictionary<string, object>)dictionary;

                foreach (var key in d1.Keys)
                {
                    treeViewElements.Add(new TreeViewElements()
                    {
                        ID = increment.ToString(),
                        Name = key,
                        Parent_ID = parentId,
                        CID = -2
                    });

                    increment++;

                    Read(d1[key], treeViewElements, inherritanceLevel + 1, (increment - 1).ToString());
                }
                return;
            }
        }

        // TODO: consider using HybridDictionary
        private void AddRow(int row, int column, SortedDictionary<string, object> dictionary)
        {
            if (column == sheet.Dimension.End.Column - 2)
            {
                if (!dictionary.ContainsKey(sheet.GetValue(row, column).ToString()))
                {
                    dictionary[sheet.GetValue(row, column).ToString()] = new SortedDictionary<string, List<string>>();
                }

                if (!((SortedDictionary<string, List<string>>)dictionary[sheet.GetValue(row, column).ToString()])
                    .ContainsKey(sheet.GetValue(row, column + 1).ToString()))
                {
                    ((SortedDictionary<string, List<string>>)dictionary[sheet.GetValue(row, column).ToString()])
                                        [sheet.GetValue(row, column + 1).ToString()] = new List<string>();
                }

                //TODO: учесть, что строки могут повторяться
                ((SortedDictionary<string, List<string>>)dictionary[sheet.GetValue(row, column).ToString()])
                    [sheet.GetValue(row, column + 1).ToString()]
                    .Add(sheet.GetValue(row, column + 2).ToString());

                return;
            }

            if (!dictionary.ContainsKey(sheet.GetValue(row, column).ToString()))
            {
                dictionary[sheet.GetValue(row, column).ToString()] = new SortedDictionary<string, object>();
            }

            AddRow(row, column + 1, (SortedDictionary<string, object>)dictionary[sheet.GetValue(row, column).ToString()]);
        }
    }
}