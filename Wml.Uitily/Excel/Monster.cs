using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic.FileIO;
using NPOI.HSSF.UserModel;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Web;

namespace Wml.Uitily.Excel
{
	public class Monster
	{
		public DataTable XlsxToDataTable(string vFilePath, string vSheetName)
		{
			DataTable dataTable = new DataTable();
			try
			{
				SLDocument sLDocument = new SLDocument(vFilePath, vSheetName);
				dataTable.TableName = vSheetName;
				SLWorksheetStatistics worksheetStatistics = sLDocument.GetWorksheetStatistics();
				int startColumnIndex = worksheetStatistics.StartColumnIndex;
				int endColumnIndex = worksheetStatistics.EndColumnIndex;
				int startRowIndex = worksheetStatistics.StartRowIndex;
				int endRowIndex = worksheetStatistics.EndRowIndex;
				for (int num = startColumnIndex; num <= endColumnIndex; num++)
				{
					SLRstType cellValueAsRstType = sLDocument.GetCellValueAsRstType(1, num);
					dataTable.Columns.Add(new DataColumn(cellValueAsRstType.GetText(), typeof(string)));
				}
				for (int num2 = startRowIndex + 1; num2 <= endRowIndex; num2++)
				{
					DataRow dataRow = dataTable.NewRow();
					for (int num = startColumnIndex; num <= endColumnIndex; num++)
					{
						dataRow[num - 1] = sLDocument.GetCellValueAsString(num2, num);
					}
					dataTable.Rows.Add(dataRow);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Xlsx to DataTable: \n" + ex.Message);
			}
			return dataTable;
		}
		public DataTable XlsxToDataTable(string vFilePath)
		{
			DataTable dataTable = new DataTable();
			try
			{
				SLDocument sLDocument = new SLDocument(vFilePath);
				dataTable.TableName = sLDocument.GetSheetNames()[0];
				SLWorksheetStatistics worksheetStatistics = sLDocument.GetWorksheetStatistics();
				int startColumnIndex = worksheetStatistics.StartColumnIndex;
				int endColumnIndex = worksheetStatistics.EndColumnIndex;
				int startRowIndex = worksheetStatistics.StartRowIndex;
				int endRowIndex = worksheetStatistics.EndRowIndex;
				for (int num = startColumnIndex; num <= endColumnIndex; num++)
				{
					SLRstType cellValueAsRstType = sLDocument.GetCellValueAsRstType(1, num);
					dataTable.Columns.Add(new DataColumn(cellValueAsRstType.GetText(), typeof(string)));
				}
				for (int num2 = startRowIndex + 1; num2 <= endRowIndex; num2++)
				{
					DataRow dataRow = dataTable.NewRow();
					for (int num = startColumnIndex; num <= endColumnIndex; num++)
					{
						dataRow[num - 1] = sLDocument.GetCellValueAsString(num2, num);
					}
					dataTable.Rows.Add(dataRow);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Xlsx to DataTable: \n" + ex.Message);
			}
			return dataTable;
		}
		public DataTable XlsxToDataTable(string vFilePath, string vSheetName, int vJumpToRow)
		{
			DataTable dataTable = new DataTable();
			try
			{
				SLDocument sLDocument = new SLDocument(vFilePath, vSheetName);
				dataTable.TableName = vSheetName;
				SLWorksheetStatistics worksheetStatistics = sLDocument.GetWorksheetStatistics();
				int startColumnIndex = worksheetStatistics.StartColumnIndex;
				int endColumnIndex = worksheetStatistics.EndColumnIndex;
				int endRowIndex = worksheetStatistics.EndRowIndex;
				for (int num = startColumnIndex; num <= endColumnIndex; num++)
				{
					SLRstType cellValueAsRstType = sLDocument.GetCellValueAsRstType(vJumpToRow, num);
					dataTable.Columns.Add(new DataColumn(cellValueAsRstType.GetText(), typeof(string)));
				}
				for (int num2 = vJumpToRow + 1; num2 <= endRowIndex; num2++)
				{
					DataRow dataRow = dataTable.NewRow();
					for (int num = startColumnIndex; num <= endColumnIndex; num++)
					{
						dataRow[num - 1] = sLDocument.GetCellValueAsString(num2, num);
					}
					dataTable.Rows.Add(dataRow);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Xlsx to DataTable: \n" + ex.Message);
			}
			return dataTable;
		}
		public DataTable XlsxToDataTable(string vFilePath, int vJumpToRow)
		{
			DataTable dataTable = new DataTable();
			try
			{
				SLDocument sLDocument = new SLDocument(vFilePath);
				dataTable.TableName = sLDocument.GetSheetNames()[0];
				SLWorksheetStatistics worksheetStatistics = sLDocument.GetWorksheetStatistics();
				int startColumnIndex = worksheetStatistics.StartColumnIndex;
				int endColumnIndex = worksheetStatistics.EndColumnIndex;
				int endRowIndex = worksheetStatistics.EndRowIndex;
				for (int num = startColumnIndex; num <= endColumnIndex; num++)
				{
					SLRstType cellValueAsRstType = sLDocument.GetCellValueAsRstType(vJumpToRow, num);
					dataTable.Columns.Add(new DataColumn(cellValueAsRstType.GetText()));
				}
				for (int num2 = vJumpToRow + 1; num2 <= endRowIndex; num2++)
				{
					DataRow dataRow = dataTable.NewRow();
					for (int num = startColumnIndex; num <= endColumnIndex; num++)
					{
						dataRow[num - 1] = sLDocument.GetCellValueAsString(num2, num);
					}
					dataTable.Rows.Add(dataRow);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Xlsx to DataTable: \n" + ex.Message);
			}
			return dataTable;
		}
		public DataTable CsvToDataTable(string vFilePath)
		{
			DataTable dataTable = new DataTable();
			try
			{
				TextFieldParser textFieldParser = new TextFieldParser(vFilePath, Encoding.UTF8);
				textFieldParser.SetDelimiters(new string[]
				{
					","
				});
				textFieldParser.HasFieldsEnclosedInQuotes = true;
				string[] array = textFieldParser.ReadFields();
				string[] array2 = array;
				for (int i = 0; i < array2.Length; i++)
				{
					string text = array2[i];
					DataColumn dataColumn = new DataColumn();
					dataColumn.AllowDBNull = true;
					if (!string.IsNullOrEmpty(text))
					{
						dataColumn.ColumnName = text;
					}
					dataTable.Columns.Add(dataColumn);
				}
				while (!textFieldParser.EndOfData)
				{
					string[] array3 = textFieldParser.ReadFields();
					for (int j = 0; j < array3.Length; j++)
					{
						if (array3[j] == "")
						{
							array3[j] = null;
						}
					}
					dataTable.Rows.Add(array3);
				}
				textFieldParser.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("Csv to DataTable : \n" + ex.Message);
			}
			return dataTable;
		}
		public string DataTableToCsv(DataTable vContent, string vOutputFilePath)
		{
			string result;
			try
			{
				if (File.Exists(vOutputFilePath))
				{
					File.Delete(vOutputFilePath);
				}
				StringBuilder stringBuilder = new StringBuilder();
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					stringBuilder.Append(vContent.Columns[i].ColumnName);
					stringBuilder.Append((i == vContent.Columns.Count - 1) ? "\n" : ",");
				}
				foreach (DataRow dataRow in vContent.Rows)
				{
					for (int i = 0; i < vContent.Columns.Count; i++)
					{
						stringBuilder.Append(dataRow[i].ToString().Trim());
						stringBuilder.Append((i == vContent.Columns.Count - 1) ? "\n" : ",");
					}
				}
				File.WriteAllText(vOutputFilePath, stringBuilder.ToString(), Encoding.UTF8);
				result = "OK";
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx : \n" + ex.Message);
			}
			return result;
		}
		public string DataTableToXlsx(DataTable vContent, string vOutputFilePath)
		{
			string result;
			try
			{
				SLDocument sLDocument = new SLDocument();
				if (File.Exists(vOutputFilePath))
				{
					File.Delete(vOutputFilePath);
				}
				sLDocument.ImportDataTable(1, 1, vContent, true);
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					SLStyle sLStyle = sLDocument.CreateStyle();
					if (vContent.Columns[i].DataType.FullName.Equals("System.String"))
					{
						sLStyle.FormatCode = "@";
					}
					else
					{
						if (vContent.Columns[i].DataType.FullName.Equals("System.DateTime"))
						{
							sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
						}
						else
						{
							if (vContent.Columns[i].DataType.FullName.Equals("System.Int16"))
							{
								sLStyle.FormatCode = "#";
							}
							else
							{
								if (vContent.Columns[i].DataType.FullName.Equals("System.Int32"))
								{
									sLStyle.FormatCode = "#";
								}
								else
								{
									if (vContent.Columns[i].DataType.FullName.Equals("System.Int64"))
									{
										sLStyle.FormatCode = "#";
									}
									else
									{
										sLStyle.FormatCode = "General";
									}
								}
							}
						}
					}
					sLDocument.SetColumnStyle(i + 1, sLStyle);
					sLDocument.AutoFitColumn(i + 1, vContent.Columns.Count, 300.0);
				}
				SLTable sLTable = sLDocument.CreateTable(1, 1, vContent.Rows.Count + 1, vContent.Columns.Count);
				sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
				sLDocument.InsertTable(sLTable);
				sLDocument.SaveAs(vOutputFilePath);
				result = "OK";
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx : \n" + ex.Message);
			}
			return result;
		}
		public string DataTableToXlsx(DataTable vContent, string vOutputFilePath, string vTitle)
		{
			string result;
			try
			{
				SLDocument sLDocument = new SLDocument();
				if (File.Exists(vOutputFilePath))
				{
					File.Delete(vOutputFilePath);
				}
				int num = 2;
				int num2 = 1;
				int num3 = 1;
				int num4 = 1;
				sLDocument.ImportDataTable(num, num2, vContent, true);
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					SLStyle sLStyle = sLDocument.CreateStyle();
					if (vContent.Columns[i].DataType.FullName.Equals("System.String"))
					{
						sLStyle.FormatCode = "@";
					}
					else
					{
						if (vContent.Columns[i].DataType.FullName.Equals("System.DateTime"))
						{
							sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
						}
						else
						{
							if (vContent.Columns[i].DataType.FullName.Equals("System.Int16"))
							{
								sLStyle.FormatCode = "#";
							}
							else
							{
								if (vContent.Columns[i].DataType.FullName.Equals("System.Int32"))
								{
									sLStyle.FormatCode = "#";
								}
								else
								{
									if (vContent.Columns[i].DataType.FullName.Equals("System.Int64"))
									{
										sLStyle.FormatCode = "#";
									}
									else
									{
										sLStyle.FormatCode = "General";
									}
								}
							}
						}
					}
					sLDocument.SetColumnStyle(i + 1, sLStyle);
					sLDocument.AutoFitColumn(i + 1, vContent.Columns.Count, 300.0);
				}
				SLTable sLTable = sLDocument.CreateTable(num, num2, vContent.Rows.Count + 1, vContent.Columns.Count);
				sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
				sLDocument.InsertTable(sLTable);
				SLStyle sLStyle2 = sLDocument.CreateStyle();
				sLStyle2.Alignment.Horizontal = HorizontalAlignmentValues.Center;
				sLStyle2.Alignment.Vertical = VerticalAlignmentValues.Center;
				sLDocument.SetCellStyle(num3, num4, sLStyle2);
				sLDocument.SetCellValue(num3, num4, vTitle);
				sLDocument.MergeWorksheetCells(num3, num4, num3, vContent.Columns.Count);
				sLDocument.SaveAs(vOutputFilePath);
				result = "OK";
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx : \n" + ex.Message);
			}
			return result;
		}
		public string DataTableToXlsx(DataTable vContent, string vOutputFilePath, string vTitle, string vPrintTime)
		{
			string result;
			try
			{
				SLDocument sLDocument = new SLDocument();
				if (File.Exists(vOutputFilePath))
				{
					File.Delete(vOutputFilePath);
				}
				int num = 2;
				int num2 = 1;
				int num3 = 1;
				int num4 = 1;
				sLDocument.ImportDataTable(num, num2, vContent, true);
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					SLStyle sLStyle = sLDocument.CreateStyle();
					if (vContent.Columns[i].DataType.FullName.Equals("System.String"))
					{
						sLStyle.FormatCode = "@";
					}
					else
					{
						if (vContent.Columns[i].DataType.FullName.Equals("System.DateTime"))
						{
							sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
						}
						else
						{
							if (vContent.Columns[i].DataType.FullName.Equals("System.Int16"))
							{
								sLStyle.FormatCode = "#";
							}
							else
							{
								if (vContent.Columns[i].DataType.FullName.Equals("System.Int32"))
								{
									sLStyle.FormatCode = "#";
								}
								else
								{
									if (vContent.Columns[i].DataType.FullName.Equals("System.Int64"))
									{
										sLStyle.FormatCode = "#";
									}
									else
									{
										sLStyle.FormatCode = "General";
									}
								}
							}
						}
					}
					sLDocument.SetColumnStyle(i + 1, sLStyle);
					sLDocument.AutoFitColumn(i + 1, vContent.Columns.Count, 300.0);
				}
				SLTable sLTable = sLDocument.CreateTable(num, num2, vContent.Rows.Count + 1, vContent.Columns.Count);
				sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
				sLDocument.InsertTable(sLTable);
				SLStyle sLStyle2 = sLDocument.CreateStyle();
				sLStyle2.Alignment.Horizontal = HorizontalAlignmentValues.Center;
				sLStyle2.Alignment.Vertical = VerticalAlignmentValues.Center;
				sLDocument.SetCellStyle(num3, num4, sLStyle2);
				sLDocument.SetCellValue(num3, num4, vTitle);
				SLStyle sLStyle3 = sLDocument.CreateStyle();
				sLStyle3.Alignment.Horizontal = HorizontalAlignmentValues.Left;
				sLStyle3.Alignment.Vertical = VerticalAlignmentValues.Center;
				sLStyle3.Font.FontSize = new double?(8.0);
				int num5;
				if (vContent.Columns.Count > 2)
				{
					num5 = vContent.Columns.Count - 1;
					int rowIndex = 1;
					int columnIndex = num5 + 1;
					sLDocument.SetCellStyle(rowIndex, columnIndex, sLStyle3);
					sLDocument.SetCellValue(rowIndex, columnIndex, "time:" + vPrintTime);
				}
				else
				{
					num5 = vContent.Columns.Count;
					int rowIndex = 1;
					int columnIndex = num5 + 1;
					sLDocument.SetCellStyle(rowIndex, columnIndex, sLStyle3);
					sLDocument.SetCellValue(rowIndex, columnIndex, "time:" + vPrintTime);
				}
				sLDocument.MergeWorksheetCells(num3, num4, num3, num5);
				sLDocument.SaveAs(vOutputFilePath);
				result = "OK";
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx : \n" + ex.Message);
			}
			return result;
		}
		public string DataTableToXlsx(DataTable vContent, string vOutputFilePath, string vTitle, Dictionary<string, string> vParams)
		{
			string result;
			try
			{
				SLDocument sLDocument = new SLDocument();
				if (File.Exists(vOutputFilePath))
				{
					File.Delete(vOutputFilePath);
				}
				int num = 2 + vParams.Keys.Count;
				int num2 = 1;
				int num3 = 1;
				int num4 = 1;
				sLDocument.ImportDataTable(num, num2, vContent, true);
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					SLStyle sLStyle = sLDocument.CreateStyle();
					if (vContent.Columns[i].DataType.FullName.Equals("System.String"))
					{
						sLStyle.FormatCode = "@";
					}
					else
					{
						if (vContent.Columns[i].DataType.FullName.Equals("System.DateTime"))
						{
							sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
						}
						else
						{
							if (vContent.Columns[i].DataType.FullName.Equals("System.Int16"))
							{
								sLStyle.FormatCode = "#";
							}
							else
							{
								if (vContent.Columns[i].DataType.FullName.Equals("System.Int32"))
								{
									sLStyle.FormatCode = "#";
								}
								else
								{
									if (vContent.Columns[i].DataType.FullName.Equals("System.Int64"))
									{
										sLStyle.FormatCode = "#";
									}
									else
									{
										sLStyle.FormatCode = "General";
									}
								}
							}
						}
					}
					sLDocument.SetColumnStyle(i + 1, sLStyle);
					sLDocument.AutoFitColumn(i + 1, vContent.Columns.Count, 300.0);
				}
				SLTable sLTable = sLDocument.CreateTable(num, num2, vContent.Rows.Count + 1, vContent.Columns.Count);
				sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
				sLDocument.InsertTable(sLTable);
				SLStyle sLStyle2 = sLDocument.CreateStyle();
				sLStyle2.Alignment.Horizontal = HorizontalAlignmentValues.Center;
				sLStyle2.Alignment.Vertical = VerticalAlignmentValues.Center;
				sLDocument.SetCellStyle(num3, num4, sLStyle2);
				sLDocument.SetCellValue(num3, num4, vTitle);
				sLDocument.MergeWorksheetCells(num3, num4, num3, vContent.Columns.Count);
				if (vParams.Keys.Count > 0)
				{
					SLStyle sLStyle3 = sLDocument.CreateStyle();
					sLStyle3.Alignment.Horizontal = HorizontalAlignmentValues.Right;
					int i = 0;
					foreach (string current in vParams.Keys)
					{
						int rowIndex = 2 + i;
						int columnIndex = 1;
						sLDocument.SetCellValue(rowIndex, columnIndex, current);
						sLDocument.SetCellStyle(rowIndex, columnIndex, sLStyle3);
						int num5 = 2 + i;
						int num6 = 2;
						int count = vContent.Columns.Count;
						sLDocument.MergeWorksheetCells(num5, num6, num5, count);
						sLDocument.SetCellValue(num5, num6, vParams[current].ToString());
						i++;
					}
				}
				sLDocument.SaveAs(vOutputFilePath);
				result = "OK";
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx : \n" + ex.Message);
			}
			return result;
		}
		public string DataTableToXlsx(DataTable vContent, string vOutputFilePath, string vTitle, string vPrintTime, Dictionary<string, string> vParams)
		{
			string result;
			try
			{
				SLDocument sLDocument = new SLDocument();
				if (File.Exists(vOutputFilePath))
				{
					File.Delete(vOutputFilePath);
				}
				int num = 2 + vParams.Keys.Count;
				int num2 = 1;
				int num3 = 1;
				int num4 = 1;
				sLDocument.ImportDataTable(num, num2, vContent, true);
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					SLStyle sLStyle = sLDocument.CreateStyle();
					if (vContent.Columns[i].DataType.FullName.Equals("System.String"))
					{
						sLStyle.FormatCode = "@";
					}
					else
					{
						if (vContent.Columns[i].DataType.FullName.Equals("System.DateTime"))
						{
							sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
						}
						else
						{
							if (vContent.Columns[i].DataType.FullName.Equals("System.Int16"))
							{
								sLStyle.FormatCode = "#";
							}
							else
							{
								if (vContent.Columns[i].DataType.FullName.Equals("System.Int32"))
								{
									sLStyle.FormatCode = "#";
								}
								else
								{
									if (vContent.Columns[i].DataType.FullName.Equals("System.Int64"))
									{
										sLStyle.FormatCode = "#";
									}
									else
									{
										sLStyle.FormatCode = "General";
									}
								}
							}
						}
					}
					sLDocument.SetColumnStyle(i + 1, sLStyle);
					sLDocument.AutoFitColumn(i + 1, vContent.Columns.Count, 300.0);
				}
				SLTable sLTable = sLDocument.CreateTable(num, num2, vContent.Rows.Count + 1, vContent.Columns.Count);
				sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
				sLDocument.InsertTable(sLTable);
				SLStyle sLStyle2 = sLDocument.CreateStyle();
				sLStyle2.Alignment.Horizontal = HorizontalAlignmentValues.Center;
				sLStyle2.Alignment.Vertical = VerticalAlignmentValues.Center;
				sLDocument.SetCellStyle(num3, num4, sLStyle2);
				sLDocument.SetCellValue(num3, num4, vTitle);
				SLStyle sLStyle3 = sLDocument.CreateStyle();
				sLStyle3.Alignment.Horizontal = HorizontalAlignmentValues.Left;
				sLStyle3.Alignment.Vertical = VerticalAlignmentValues.Center;
				sLStyle3.Font.FontSize = new double?(8.0);
				int num5;
				if (vContent.Columns.Count > 2)
				{
					num5 = vContent.Columns.Count - 1;
					int rowIndex = 1;
					int columnIndex = num5 + 1;
					sLDocument.SetCellStyle(rowIndex, columnIndex, sLStyle3);
					sLDocument.SetCellValue(rowIndex, columnIndex, "time:" + vPrintTime);
				}
				else
				{
					num5 = vContent.Columns.Count;
					int rowIndex = 1;
					int columnIndex = num5 + 1;
					sLDocument.SetCellStyle(rowIndex, columnIndex, sLStyle3);
					sLDocument.SetCellValue(rowIndex, columnIndex, "time:" + vPrintTime);
				}
				sLDocument.MergeWorksheetCells(num3, num4, num3, num5);
				if (vParams.Keys.Count > 0)
				{
					SLStyle sLStyle4 = sLDocument.CreateStyle();
					sLStyle4.Alignment.Horizontal = HorizontalAlignmentValues.Right;
					int i = 0;
					foreach (string current in vParams.Keys)
					{
						int rowIndex2 = 2 + i;
						int columnIndex2 = 1;
						sLDocument.SetCellValue(rowIndex2, columnIndex2, current);
						sLDocument.SetCellStyle(rowIndex2, columnIndex2, sLStyle4);
						int num6 = 2 + i;
						int num7 = 2;
						int count = vContent.Columns.Count;
						sLDocument.MergeWorksheetCells(num6, num7, num6, count);
						sLDocument.SetCellValue(num6, num7, vParams[current].ToString());
						i++;
					}
				}
				sLDocument.SaveAs(vOutputFilePath);
				result = "OK";
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx : \n" + ex.Message);
			}
			return result;
		}
		public void XlsxToWebResponse(DataTable vContent, HttpResponse vResponse, string vFileName)
		{
			try
			{
				SLDocument sLDocument = new SLDocument();
				sLDocument.ImportDataTable(1, 1, vContent, true);
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					SLStyle sLStyle = sLDocument.CreateStyle();
					if (vContent.Columns[i].DataType.FullName.Equals("System.String"))
					{
						sLStyle.FormatCode = "@";
					}
					else
					{
						if (vContent.Columns[i].DataType.FullName.Equals("System.DateTime"))
						{
							sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
						}
						else
						{
							if (vContent.Columns[i].DataType.FullName.Equals("System.Int16"))
							{
								sLStyle.FormatCode = "#";
							}
							else
							{
								if (vContent.Columns[i].DataType.FullName.Equals("System.Int32"))
								{
									sLStyle.FormatCode = "#";
								}
								else
								{
									if (vContent.Columns[i].DataType.FullName.Equals("System.Int64"))
									{
										sLStyle.FormatCode = "#";
									}
									else
									{
										sLStyle.FormatCode = "General";
									}
								}
							}
						}
					}
					sLDocument.SetColumnStyle(i + 1, sLStyle);
					sLDocument.AutoFitColumn(i + 1, vContent.Columns.Count, 300.0);
				}
				SLTable sLTable = sLDocument.CreateTable(1, 1, vContent.Rows.Count + 1, vContent.Columns.Count);
				sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
				sLDocument.InsertTable(sLTable);
				MemoryStream memoryStream = new MemoryStream();
				sLDocument.SaveAs(memoryStream);
				byte[] buffer = memoryStream.ToArray();
				vResponse.Buffer = true;
				vResponse.Clear();
				vResponse.ContentType = "application/force-download";
				vResponse.AddHeader("content-disposition", "attachment;filename=" + vFileName);
				vResponse.BinaryWrite(buffer);
				memoryStream.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx/Web Response : \n" + ex.Message);
			}
		}
		public byte[] XlsxToByte(DataTable vContent)
		{
			byte[] result;
			try
			{
				SLDocument sLDocument = new SLDocument();
				sLDocument.ImportDataTable(1, 1, vContent, true);
				for (int i = 0; i < vContent.Columns.Count; i++)
				{
					SLStyle sLStyle = sLDocument.CreateStyle();
					if (vContent.Columns[i].DataType.FullName.Equals("System.String"))
					{
						sLStyle.FormatCode = "@";
					}
					else
					{
						if (vContent.Columns[i].DataType.FullName.Equals("System.DateTime"))
						{
							sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
						}
						else
						{
							if (vContent.Columns[i].DataType.FullName.Equals("System.Int16"))
							{
								sLStyle.FormatCode = "#";
							}
							else
							{
								if (vContent.Columns[i].DataType.FullName.Equals("System.Int32"))
								{
									sLStyle.FormatCode = "#";
								}
								else
								{
									if (vContent.Columns[i].DataType.FullName.Equals("System.Int64"))
									{
										sLStyle.FormatCode = "#";
									}
									else
									{
										sLStyle.FormatCode = "General";
									}
								}
							}
						}
					}
					sLDocument.SetColumnStyle(i + 1, sLStyle);
					sLDocument.AutoFitColumn(i + 1, vContent.Columns.Count, 300.0);
				}
				SLTable sLTable = sLDocument.CreateTable(1, 1, vContent.Rows.Count + 1, vContent.Columns.Count);
				sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
				sLDocument.InsertTable(sLTable);
				MemoryStream memoryStream = new MemoryStream();
				sLDocument.SaveAs(memoryStream);
				result = memoryStream.ToArray();
				memoryStream.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("DataTable to Xlsx/Web Response : \n" + ex.Message);
			}
			return result;
		}
		public string DataSetToXlsx(DataSet vContent, string vOutputFilePath)
		{
			string result;
			try
			{
				if (vContent != null && vContent.Tables.Count > 0)
				{
					using (SLDocument sLDocument = new SLDocument())
					{
						if (File.Exists(vOutputFilePath))
						{
							File.Delete(vOutputFilePath);
						}
						foreach (DataTable dataTable in vContent.Tables)
						{
							sLDocument.AddWorksheet(dataTable.TableName);
							sLDocument.ImportDataTable(1, 1, dataTable, true);
							for (int i = 0; i < dataTable.Columns.Count; i++)
							{
								SLStyle sLStyle = sLDocument.CreateStyle();
								if (dataTable.Columns[i].DataType.FullName.Equals("System.String"))
								{
									sLStyle.FormatCode = "@";
								}
								else
								{
									if (dataTable.Columns[i].DataType.FullName.Equals("System.DateTime"))
									{
										sLStyle.FormatCode = "yyyy/mm/dd hh:mm:ss";
									}
									else
									{
										if (dataTable.Columns[i].DataType.FullName.Equals("System.Int16"))
										{
											sLStyle.FormatCode = "#";
										}
										else
										{
											if (dataTable.Columns[i].DataType.FullName.Equals("System.Int32"))
											{
												sLStyle.FormatCode = "#";
											}
											else
											{
												if (dataTable.Columns[i].DataType.FullName.Equals("System.Int64"))
												{
													sLStyle.FormatCode = "#";
												}
												else
												{
													sLStyle.FormatCode = "General";
												}
											}
										}
									}
								}
								sLDocument.SetColumnStyle(i + 1, sLStyle);
								sLDocument.AutoFitColumn(i + 1, dataTable.Columns.Count, 300.0);
							}
							SLTable sLTable = sLDocument.CreateTable(1, 1, dataTable.Rows.Count + 1, dataTable.Columns.Count);
							sLTable.SetTableStyle(SLTableStyleTypeValues.Medium1);
							sLDocument.InsertTable(sLTable);
						}
						sLDocument.DeleteWorksheet("Sheet1");
						sLDocument.SaveAs(vOutputFilePath);
					}
					result = "OK";
				}
				else
				{
					result = "DataSet is null";
				}
			}
			catch (Exception ex)
			{
				throw new Exception("DataSet to Xlsx : \n" + ex.Message);
			}
			return result;
		}
		public DataTable XlsToDataTable(string vFilePath, string vSheetName)
		{
			DataTable dataTable = new DataTable();
			Stream stream = null;
			try
			{
				stream = File.OpenRead(vFilePath);
				HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(stream);
				HSSFSheet hSSFSheet = (HSSFSheet)hSSFWorkbook.GetSheet(vSheetName);
				HSSFRow hSSFRow = (HSSFRow)hSSFSheet.GetRow(0);
				int lastCellNum = (int)hSSFRow.LastCellNum;
				for (int i = (int)hSSFRow.FirstCellNum; i < lastCellNum; i++)
				{
					DataColumn column = new DataColumn(hSSFRow.GetCell(i).StringCellValue);
					dataTable.Columns.Add(column);
				}
				dataTable.TableName = vSheetName;
				int lastRowNum = hSSFSheet.LastRowNum;
				for (int i = hSSFSheet.FirstRowNum + 1; i <= hSSFSheet.LastRowNum; i++)
				{
					HSSFRow hSSFRow2 = (HSSFRow)hSSFSheet.GetRow(i);
					DataRow dataRow = dataTable.NewRow();
					for (int j = (int)hSSFRow2.FirstCellNum; j < lastCellNum; j++)
					{
						dataRow[j] = hSSFRow2.GetCell(j).ToString();
					}
					dataTable.Rows.Add(dataRow);
				}
				stream.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("Xls to DataTable: \n" + ex.Message);
			}
			finally
			{
				if (stream != null)
				{
					stream.Close();
				}
			}
			return dataTable;
		}
		public DataTable XlsToDataTable(string vFilePath)
		{
			DataTable dataTable = new DataTable();
			Stream stream = null;
			try
			{
				stream = File.OpenRead(vFilePath);
				HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(stream);
				HSSFSheet hSSFSheet = (HSSFSheet)hSSFWorkbook.GetSheetAt(hSSFWorkbook.ActiveSheetIndex);
				HSSFRow hSSFRow = (HSSFRow)hSSFSheet.GetRow(0);
				int lastCellNum = (int)hSSFRow.LastCellNum;
				for (int i = (int)hSSFRow.FirstCellNum; i < lastCellNum; i++)
				{
					DataColumn column = new DataColumn(hSSFRow.GetCell(i).StringCellValue);
					dataTable.Columns.Add(column);
				}
				dataTable.TableName = hSSFSheet.SheetName;
				int lastRowNum = hSSFSheet.LastRowNum;
				for (int i = hSSFSheet.FirstRowNum + 1; i <= hSSFSheet.LastRowNum; i++)
				{
					HSSFRow hSSFRow2 = (HSSFRow)hSSFSheet.GetRow(i);
					DataRow dataRow = dataTable.NewRow();
					for (int j = (int)hSSFRow2.FirstCellNum; j < lastCellNum; j++)
					{
						dataRow[j] = hSSFRow2.GetCell(j).ToString();
					}
					dataTable.Rows.Add(dataRow);
				}
				stream.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("Xls to DataTable: \n" + ex.Message);
			}
			finally
			{
				if (stream != null)
				{
					stream.Close();
				}
			}
			return dataTable;
		}
		public DataTable XlsToDataTable(string vFilePath, string vSheetName, int vJumpToRow)
		{
			DataTable dataTable = new DataTable();
			Stream stream = null;
			try
			{
				stream = File.OpenRead(vFilePath);
				HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(stream);
				HSSFSheet hSSFSheet = (HSSFSheet)hSSFWorkbook.GetSheet(vSheetName);
				HSSFRow hSSFRow = (HSSFRow)hSSFSheet.GetRow(vJumpToRow - 1);
				int lastCellNum = (int)hSSFRow.LastCellNum;
				for (int i = (int)hSSFRow.FirstCellNum; i < lastCellNum; i++)
				{
					DataColumn column = new DataColumn(hSSFRow.GetCell(i).StringCellValue);
					dataTable.Columns.Add(column);
				}
				dataTable.TableName = vSheetName;
				int lastRowNum = hSSFSheet.LastRowNum;
				for (int i = hSSFSheet.FirstRowNum + vJumpToRow; i <= hSSFSheet.LastRowNum; i++)
				{
					HSSFRow hSSFRow2 = (HSSFRow)hSSFSheet.GetRow(i);
					DataRow dataRow = dataTable.NewRow();
					for (int j = (int)hSSFRow2.FirstCellNum; j < lastCellNum; j++)
					{
						if (hSSFRow2.GetCell(j) == null)
						{
							dataRow[j] = "";
						}
						else
						{
							dataRow[j] = hSSFRow2.GetCell(j).ToString();
						}
					}
					dataTable.Rows.Add(dataRow);
				}
				stream.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("Xls to DataTable: \n" + ex.Message);
			}
			finally
			{
				if (stream != null)
				{
					stream.Close();
				}
			}
			return dataTable;
		}
		public DataTable XlsToDataTable(string vFilePath, int vJumpToRow)
		{
			DataTable dataTable = new DataTable();
			Stream stream = null;
			try
			{
				stream = File.OpenRead(vFilePath);
				HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(stream);
				HSSFSheet hSSFSheet = (HSSFSheet)hSSFWorkbook.GetSheetAt(hSSFWorkbook.ActiveSheetIndex);
				HSSFRow hSSFRow = (HSSFRow)hSSFSheet.GetRow(vJumpToRow - 1);
				int lastCellNum = (int)hSSFRow.LastCellNum;
				for (int i = (int)hSSFRow.FirstCellNum; i < lastCellNum; i++)
				{
					DataColumn column = new DataColumn(hSSFRow.GetCell(i).StringCellValue);
					dataTable.Columns.Add(column);
				}
				dataTable.TableName = hSSFSheet.SheetName;
				int lastRowNum = hSSFSheet.LastRowNum;
				for (int i = hSSFSheet.FirstRowNum + vJumpToRow; i <= hSSFSheet.LastRowNum; i++)
				{
					HSSFRow hSSFRow2 = (HSSFRow)hSSFSheet.GetRow(i);
					DataRow dataRow = dataTable.NewRow();
					for (int j = (int)hSSFRow2.FirstCellNum; j < lastCellNum; j++)
					{
						if (hSSFRow2.GetCell(j) == null)
						{
							dataRow[j] = "";
						}
						else
						{
							dataRow[j] = hSSFRow2.GetCell(j).ToString();
						}
					}
					dataTable.Rows.Add(dataRow);
				}
				stream.Close();
			}
			catch (Exception ex)
			{
				throw new Exception("Xls to DataTable: \n" + ex.Message);
			}
			finally
			{
				if (stream != null)
				{
					stream.Close();
				}
			}
			return dataTable;
		}
		public DataSet XlsToDataSet(string vFilePath)
		{
			DataSet dataSet = new DataSet();
			Stream stream = null;
			try
			{
				stream = File.OpenRead(vFilePath);
				HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(stream);
				int numberOfSheets = hSSFWorkbook.NumberOfSheets;
				for (int i = 0; i < numberOfSheets; i++)
				{
					HSSFSheet hSSFSheet = (HSSFSheet)hSSFWorkbook.GetSheetAt(i);
					DataTable dataTable = new DataTable();
					dataTable.TableName = hSSFSheet.SheetName;
					HSSFRow hSSFRow = (HSSFRow)hSSFSheet.GetRow(0);
					int lastCellNum = (int)hSSFRow.LastCellNum;
					for (int j = (int)hSSFRow.FirstCellNum; j < lastCellNum; j++)
					{
						DataColumn column = new DataColumn(hSSFRow.GetCell(j).StringCellValue);
						dataTable.Columns.Add(column);
					}
					int lastRowNum = hSSFSheet.LastRowNum;
					for (int j = hSSFSheet.FirstRowNum + 1; j < hSSFSheet.LastRowNum; j++)
					{
						HSSFRow hSSFRow2 = (HSSFRow)hSSFSheet.GetRow(j);
						DataRow dataRow = dataTable.NewRow();
						for (int k = (int)hSSFRow2.FirstCellNum; k < lastCellNum; k++)
						{
							dataRow[k] = hSSFRow2.GetCell(k).ToString();
						}
						dataTable.Rows.Add(dataRow);
					}
					dataSet.Tables.Add(dataTable);
				}
			}
			catch (Exception ex)
			{
				throw new Exception("Xls to DataSet: \n" + ex.Message);
			}
			finally
			{
				if (stream != null)
				{
					stream.Close();
				}
			}
			return dataSet;
		}
	}
}
