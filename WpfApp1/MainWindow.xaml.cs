using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;
using System.Collections.Generic;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Color = System.Windows.Media.Color;
using System.Text;
using System.Windows.Documents;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;


namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            this.Background = new SolidColorBrush(Color.FromRgb(240, 240, 240));
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.Width = SystemParameters.PrimaryScreenWidth / 4; // Ограничение масштаба
            this.Height = SystemParameters.PrimaryScreenHeight / 4;
        }

        private void SelectWordFile_Click(object sender, RoutedEventArgs e)
        {
            string wordFilePath = GetSelectedFilePath("Word files (*.docx)|*.docx|All files (*.*)|*.*");

            if (wordFilePath != null)
            {
                wordFilePathTextBox.Text = wordFilePath;
                MessageBox.Show("Word файл найден!");
                LoadWordColumns(wordFilePath);
            }
        }

        private void SelectExcelFile_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = GetSelectedFilePath("Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*");

            if (excelFilePath != null)
            {
                excelFilePathTextBox.Text = excelFilePath;
                MessageBox.Show("Excel файл найден!");

                // Очищаем и загружаем новые данные в ComboBox'ы
                excelFromColumnComboBox.Items.Clear();
                excelToColumnComboBox.Items.Clear();
                sheetComboBox.Items.Clear();

                // Открываем Excel файл и загружаем информацию о листах и столбцах
                try
                {
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                        foreach (Sheet sheet in workbookPart.Workbook.Sheets)
                        {
                            sheetComboBox.Items.Add(sheet.Name);
                        }

                        sheetComboBox.SelectedIndex = 0;
                        LoadColumnsForSelectedSheet();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при открытии Excel файла: {ex.Message}");
                }
            }
        }

        private void sheetComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            LoadColumnsForSelectedSheet();
        }

        private void LoadColumnsForSelectedSheet()
        {
            string excelFilePath = excelFilePathTextBox.Text;

            if (string.IsNullOrEmpty(excelFilePath) || sheetComboBox.SelectedIndex == -1)
            {
                return;
            }

            string selectedSheetName = sheetComboBox.SelectedItem.ToString();

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    Sheet selectedSheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == selectedSheetName);
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(selectedSheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    if (sheetData != null)
                    {
                        excelFromColumnComboBox.Items.Clear();
                        excelToColumnComboBox.Items.Clear();

                        Row firstRow = sheetData.Elements<Row>().FirstOrDefault();
                        if (firstRow != null)
                        {
                            foreach (Cell cell in firstRow.Elements<Cell>())
                            {
                                string columnName = GetColumnNameFromCellReference(cell.CellReference.Value);
                                excelFromColumnComboBox.Items.Add(columnName);
                                excelToColumnComboBox.Items.Add(columnName);
                            }

                            excelFromColumnComboBox.SelectedIndex = 0;
                            excelToColumnComboBox.SelectedIndex = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке данных листа Excel: {ex.Message}");
            }
        }

        private void LoadWordColumns(string wordFilePath)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, false))
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                    if (mainPart == null)
                        throw new Exception("Не удалось открыть часть основного документа Word.");

                    Body body = mainPart.Document.Body;

                    // Находим таблицу в документе Word
                    Table wordTable = FindTargetTable(body);
                    if (wordTable == null)
                        throw new Exception("Не удалось найти таблицу в документе Word.");

                    // Очищаем ComboBox'ы
                    wordFromColumnComboBox.Items.Clear();
                    wordToColumnComboBox.Items.Clear();

                    // Получаем количество столбцов из первой строки таблицы Word
                    TableRow firstRow = wordTable.Elements<TableRow>().FirstOrDefault();
                    if (firstRow != null)
                    {
                        int columnIndex = 0;
                        foreach (TableCell cell in firstRow.Elements<TableCell>())
                        {
                            string columnName = $"Column {columnIndex + 1}";
                            wordFromColumnComboBox.Items.Add(columnName);
                            wordToColumnComboBox.Items.Add(columnName);
                            columnIndex++;
                        }

                        wordFromColumnComboBox.SelectedIndex = 0;
                        wordToColumnComboBox.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке столбцов Word файла: {ex.Message}");
            }
        }

        private void CopyData_Click(object sender, RoutedEventArgs e)
        {
            string excelFilePath = excelFilePathTextBox.Text;
            string wordFilePath = wordFilePathTextBox.Text;

            if (string.IsNullOrEmpty(excelFilePath) || string.IsNullOrEmpty(wordFilePath))
            {
                MessageBox.Show("Выберите файлы Excel и Word.");
                return;
            }

            int excelFromColumnIndex = excelFromColumnComboBox.SelectedIndex;
            int excelToColumnIndex = excelToColumnComboBox.SelectedIndex;
            int wordFromColumnIndex = wordFromColumnComboBox.SelectedIndex;
            int wordToColumnIndex = wordToColumnComboBox.SelectedIndex;

            if (excelFromColumnIndex == -1 || excelToColumnIndex == -1 || wordFromColumnIndex == -1 || wordToColumnIndex == -1)
            {
                MessageBox.Show("Выберите корректные столбцы для Excel и Word.");
                return;
            }

            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    Sheet selectedSheet = workbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetComboBox.SelectedItem.ToString());
                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(selectedSheet.Id);
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();

                    // Определяем столбцы для копирования в Excel
                    List<string> selectedColumns = new List<string>();
                    for (int i = excelFromColumnIndex; i <= excelToColumnIndex; i++)
                    {
                        selectedColumns.Add(GetExcelColumnName(i + 1));
                    }

                    // Открываем существующий Word документ
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFilePath, true))
                    {
                        MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                        if (mainPart == null)
                            throw new Exception("Не удалось открыть часть основного документа Word.");

                        Body body = mainPart.Document.Body;

                        // Находим таблицу в документе Word
                        Table wordTable = FindTargetTable(body);
                        if (wordTable == null)
                            throw new Exception("Не удалось найти таблицу в документе Word.");

                        // Получаем данные из Excel и записываем их в таблицу Word
                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            TableRow tableRow = new TableRow();

                            for (int i = 0; i < wordTable.Descendants<TableRow>().First().Elements<TableCell>().Count(); i++)
                            {
                                TableCell tableCell;
                                if (i >= wordFromColumnIndex && i <= wordToColumnIndex && selectedColumns.Count > i - wordFromColumnIndex)
                                {
                                    string column = selectedColumns[i - wordFromColumnIndex];
                                    Cell cell = row.Elements<Cell>().FirstOrDefault(c => GetColumnNameFromCellReference(c.CellReference.Value) == column);
                                    string cellValue = cell != null ? GetCellValue(cell, workbookPart) : "";
                                    tableCell = new TableCell(new Paragraph(new Run(new Text(cellValue))));

                                    // Устанавливаем выравнивание текста в ячейке Word
                                    SetAlignment(tableCell, cell, workbookPart);
                                }
                                else
                                {
                                    tableCell = new TableCell(new Paragraph(new Run(new Text(""))));
                                }
                                tableRow.Append(tableCell);
                            }

                            wordTable.Append(tableRow);
                        }

                        // Сохраняем изменения в документе Word
                        mainPart.Document.Save();
                        MessageBox.Show($"Данные из Excel успешно записаны в таблицу Word файла: {wordFilePath}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при записи данных из Excel в Word файл: {ex.Message}");
            }
        }
        private void SetAlignment(TableCell tableCell, Cell excelCell, WorkbookPart workbookPart)
        {
            if (excelCell != null && excelCell.StyleIndex != null)
            {
                CellFormat cellFormat = (CellFormat)workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt((int)excelCell.StyleIndex.Value);
                ParagraphProperties paragraphProperties = new ParagraphProperties();

                if (cellFormat.Alignment != null)
                {
                    // Устанавливаем горизонтальное выравнивание
                    var horizontalAlignment = cellFormat.Alignment.Horizontal?.Value;
                    if (horizontalAlignment == DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)
                    {
                        paragraphProperties.Justification = new Justification() { Val = JustificationValues.Center };
                    }
                    else if (horizontalAlignment == DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right)
                    {
                        paragraphProperties.Justification = new Justification() { Val = JustificationValues.Right };
                    }
                    else
                    {
                        paragraphProperties.Justification = new Justification() { Val = JustificationValues.Left };
                    }

                    // Устанавливаем вертикальное выравнивание
                    var verticalAlignment = cellFormat.Alignment.Vertical?.Value;
                    TableCellProperties tableCellProperties = new TableCellProperties();
                    if (verticalAlignment == DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Center)
                    {
                        tableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
                    }
                    else if (verticalAlignment == DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues.Bottom)
                    {
                        tableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Bottom };
                    }
                    else
                    {
                        tableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Top };
                    }

                    tableCell.TableCellProperties = tableCellProperties;
                }

                // Добавляем ParagraphProperties в TableCell
                tableCell.Elements<Paragraph>().FirstOrDefault()?.InsertAt(paragraphProperties, 0);
            }
        }

        private string GetExcelColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell.CellValue == null)
            {
                return string.Empty;
            }

            SharedStringTablePart stringTablePart = workbookPart.SharedStringTablePart;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[int.Parse(cell.CellValue.InnerText)].InnerText;
            }
            else
            {
                return cell.CellValue.InnerText;
            }
        }

        private Table FindTargetTable(Body body)
        {
            foreach (Table table in body.Elements<Table>())
            {
                return table;
            }
            return null;
        }

        private string GetSelectedFilePath(string filter)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = filter;

            if (openFileDialog.ShowDialog() == true)
            {
                return openFileDialog.FileName;
            }

            return null;
        }

        private string GetColumnNameFromCellReference(string cellReference)
        {
            string columnName = "";
            foreach (char c in cellReference)
            {
                if (!char.IsDigit(c))
                {
                    columnName += c;
                }
                else
                {
                    break;
                }
            }
            return columnName;
        }

    }
}