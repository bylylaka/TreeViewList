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
		private int increment = 1;//TODO: make concurrent safety
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
			var dictionary = new SortedDictionary<string, object>();

			for (int i = 2; i <= sheet.Dimension.Rows; i++)
			{
				AddRowPartialData(i, 1, dictionary);
			}

			var _SampleTreeView = new List<TreeViewElements>();
			Read(dictionary, _SampleTreeView, 0);

			return _SampleTreeView;
		}

		private void Read(
			object dictionary,
			List<TreeViewElements> treeViewElements,
			int inherritanceLevel,
			string parentId = "0")
		{
			if (dictionary is SortedDictionary<string, List<string>>)
			{
				// Может иметь вид
				// dictionary is SortedDictionary<string, List<string>> as d2
				// в более новой версии Microsoft.Net.Compilers (сейчас 1.3.2)
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
						CID = defaultCID
					});

					increment++;

					Read(d1[key], treeViewElements, inherritanceLevel + 1, (increment - 1).ToString());
				}
				return;
			}
		}

		// TODO: consider using HybridDictionary
		private void AddRowPartialData(int rowIndex, int columnIndex, SortedDictionary<string, object> dictionary)
		{
			var cellValue = sheet.GetValue(rowIndex, columnIndex).ToString();

			if (columnIndex == sheet.Dimension.End.Column - 2)
			{
				if (!dictionary.ContainsKey(cellValue))
				{
					dictionary[cellValue] = new SortedDictionary<string, List<string>>();
				}

				if (!((SortedDictionary<string, List<string>>)dictionary[cellValue])
					.ContainsKey(sheet.GetValue(rowIndex, columnIndex + 1).ToString()))
				{
					((SortedDictionary<string, List<string>>)dictionary[cellValue])
										[sheet.GetValue(rowIndex, columnIndex + 1).ToString()] = new List<string>();
				}

				((SortedDictionary<string, List<string>>)dictionary[cellValue])
					[sheet.GetValue(rowIndex, columnIndex + 1).ToString()]
					.Add(sheet.GetValue(rowIndex, columnIndex + 2).ToString());

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
