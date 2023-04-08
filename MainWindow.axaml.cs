using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MessageBox.Avalonia;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace LuckyDrawApp
{
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
#if DEBUG
            this.AttachDevTools();
#endif
        }

        private async void ButtonDraw_Click(object sender, RoutedEventArgs e)
        {
            // Get the participants from the multiline textbox
            var participants = textBoxParticipants.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            int numWinners = 5;

            if (participants.Length < numWinners)
            {
                await MessageBoxManager.GetMessageBoxStandardWindow("Error", $"Please enter at least {numWinners} participants.").Show();
                return;
            }

            // Shuffle the participants array
            var random = new Random();
            for (int i = participants.Length - 1; i > 0; i--)
            {
                int j = random.Next(i + 1);
                string temp = participants[i];
                participants[i] = participants[j];
                participants[j] = temp;
            }

            // Get the first 5 winners from the shuffled participants
            string[] winners = participants.Take(numWinners).ToArray();

            // Display the winners in the Winners textbox
            textBoxWinners.Text = string.Join(Environment.NewLine, winners);
        }

        private async void ButtonLoadExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "Select Excel file",
                AllowMultiple = false,
                Filters = new List<FileDialogFilter>
                {
                    new FileDialogFilter { Name = "Excel files", Extensions = { "xlsx", "xls" } },
                    new FileDialogFilter { Name = "All files", Extensions = { "*" } }
                }
            };
            string[] result = await openFileDialog.ShowAsync(this);

            if (result != null && result.Length > 0)
            {
                LoadExcelData(result[0]);
            }
        }


        private void LoadExcelData(string filePath)
        {
            List<string> participants = new List<string>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                SharedStringTablePart sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                int columnIndex = 1; // Change this value to the desired column index (1-based)

                foreach (Row row in sheetData.Elements<Row>().Skip(1))
                {
                    Cell cell = row.Descendants<Cell>().FirstOrDefault(c => GetColumnIndex(c.CellReference.Value) == columnIndex);

                    if (cell != null)
                    {
                        string cellValue;

                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            int sharedStringIndex = int.Parse(cell.InnerText);
                            cellValue = sharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
                        }
                        else
                        {
                            cellValue = cell.InnerText;
                        }

                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            participants.Add(cellValue);
                        }
                    }
                }
            }

            textBoxParticipants.Text = string.Join(Environment.NewLine, participants);
        }

        private int GetColumnIndex(string cellReference)
        {
            string columnReference = new string(cellReference.ToCharArray().Where(c => char.IsLetter(c)).ToArray());
            int columnIndex = 0;

            foreach (char c in columnReference)
            {
                columnIndex = columnIndex * 26 + (c - 'A' + 1);
            }

            return columnIndex;
        }
    }
}
