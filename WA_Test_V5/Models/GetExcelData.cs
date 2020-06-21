namespace WA_Test_V5.GetData.Excel
{
	using OfficeOpenXml;
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using WA_Test_V5.Interface.TreeView;

	public class GetExcelData
	{
		private readonly ExcelWorksheet sheet;

		private int increment = 1;

		private const int defaultCID = -2;

		public GetExcelData(string filePath)
		{
			var fInfo = new FileInfo(filePath);
			if (fInfo.Exists != true)
			{
				throw new Exception();
			}

			var pack = new ExcelPackage(fInfo);
			sheet = pack.Workbook.Worksheets.First();
		}

		public List<TreeViewElements> GetSample()
		{
			var numberOfRows = sheet.Dimension.End.Row;
			var Cells = sheet.Cells;
			var _SampleTreeView = new List<TreeViewElements>();
			for (var rowIterator = 2; rowIterator <= numberOfRows; rowIterator++)
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
			var dictionary = new SortedDictionary<string, object>();

			for (var i = 2; i <= sheet.Dimension.Rows; i++)
			{
				AddRowPartialData(i, 1, dictionary);
			}

			var _SampleTreeView = new List<TreeViewElements>();
			MapRowsDataToList(dictionary, _SampleTreeView, 0);
			return _SampleTreeView;
		}

		private void MapRowsDataToList(
			SortedDictionary<string, object> dictionary,
			List<TreeViewElements> treeViewElements,
			int inherritanceLevel,
			string parentId = "0")
		{
			if (inherritanceLevel == sheet.Dimension.Columns - 2)
			{
				foreach (var key in dictionary.Keys)
				{
					foreach (var cId in (List<string>)dictionary[key])
					{
						if (treeViewElements.Any(
							x => x.CID == int.Parse(cId) &&
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
							CID = int.Parse(cId)
						});

						increment++;
					}
				}
				return;
			}

			foreach (var key in dictionary.Keys)
			{
				treeViewElements.Add(new TreeViewElements()
				{
					ID = increment.ToString(),
					Name = key,
					Parent_ID = parentId,
					CID = defaultCID
				});

				increment++;

				MapRowsDataToList(
					(SortedDictionary<string, object>)dictionary[key],
					treeViewElements, inherritanceLevel + 1,
					(increment - 1).ToString());
			}
			return;
		}

		private void AddRowPartialData(int rowIndex, int columnIndex, SortedDictionary<string, object> dictionary)
		{
			var cellValue = sheet.GetValue(rowIndex, columnIndex).ToString();

			if (columnIndex == sheet.Dimension.End.Column - 1)
			{
				if (!dictionary.ContainsKey(cellValue))
				{
					dictionary[cellValue] = new List<string>();
				}

				((List<string>)dictionary[cellValue])
					.Add(sheet.GetValue(rowIndex, columnIndex + 1).ToString());

				return;
			}

			if (!dictionary.ContainsKey(cellValue))
			{
				dictionary[cellValue] = new SortedDictionary<string, object>();
			}

			AddRowPartialData(rowIndex, columnIndex + 1, (SortedDictionary<string, object>)dictionary[cellValue]);
		}
	}
}
