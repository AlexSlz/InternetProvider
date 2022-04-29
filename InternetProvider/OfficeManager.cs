using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace InternetProvider
{
    internal class OfficeManager
    {
        private FileInfo _blankFile;
        private string _path = "";
        private string blankFileName = "a.docx";
        DataBaseManager _dataBaseManager;

        private string _wordFolder = "Word";
        private string _excelFolder = "Excel";

        public OfficeManager(DataBaseManager d, string path)
        {
            _path = path;
            _blankFile = new FileInfo(Path.Combine(path, blankFileName));
            _dataBaseManager = d;

            if (!Directory.Exists(Path.Combine(path, _excelFolder)))
            {
                Directory.CreateDirectory(Path.Combine(path, _excelFolder));
            }
            if (!Directory.Exists(Path.Combine(path, _wordFolder)))
            {
                Directory.CreateDirectory(Path.Combine(path, _wordFolder));
            }

            foreach (Process proc in Process.GetProcessesByName("WINWORD"))
            {
                proc.Kill();
            }

        }

        public void FillBlankFile(Dictionary<string, string> items)
        {
            var app = new Word.Application();

            Object file = _blankFile.FullName;

            Object miss = Type.Missing;

            try
            {
                Word.Document document = app.Documents.Open(file);
                foreach (var item in items)
                {
                    Word.Find find = app.Selection.Find;
                    find.Text = item.Key;
                    if (item.Key.Contains("{dataTable}"))
                    {
                        for (int i = 1; i < document.Paragraphs.Count; i++)
                        {
                            if (document.Paragraphs[i].Range.Text.Contains(item.Key))
                                GenerateTable(item.Value, document, document.Paragraphs[i]);
                        }
                    }
                    else
                    {
                        find.Replacement.Text = item.Value;

                        Object wrap = Word.WdFindWrap.wdFindContinue;
                        Object replace = Word.WdReplace.wdReplaceAll;

                        find.Execute(FindText: Type.Missing, MatchCase: false, MatchWholeWord: false, MatchWildcards: false, MatchSoundsLike: miss, MatchAllWordForms: false, Forward: true, Wrap: wrap, Format: false, ReplaceWith: miss, Replace: replace);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Зі створенням договору, сталася помилка! \n{ex.Message}", "!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                Object newFileName = Path.Combine(_path, _wordFolder, $"Договір{items["{dId}"]}.{DateTime.Now.ToString("dd-MM-(HH-mm-ss)")}.docx");
                app.ActiveDocument.SaveAs2(newFileName);
                MessageBox.Show($"Файл договору {newFileName} успішно створено!", "FillBlankFile", MessageBoxButtons.OK, MessageBoxIcon.Information);
                app.ActiveDocument.Close();
                app.Quit();
            }
        }

        public void CreateWordFile(string tableName)
        {
            Word.Application application = new Word.Application();
            Word.Document document = application.Documents.Add();

            try
            {
                application.Visible = false;
                Word.Paragraph p1 = document.Paragraphs.Add();
                p1.Range.Font.Size = 14;
                p1.Range.Text = $"Таблиця {tableName}:";
                p1.Range.InsertParagraphAfter();
                GenerateTable(tableName, document, p1, 12);
            }
            catch (Exception ex)
            {
                document.Close();
                application.Quit();
                MessageBox.Show($"З експортуванням таблиці {tableName}, сталася помилка!", "CreateWordFile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                document.SaveAs2(Path.Combine(_path, _wordFolder, $"{tableName}({DateTime.Now.ToString("HH-mm-ss")}).docx"));
                document.Close();
                application.Quit();
                MessageBox.Show($"Таблиця {tableName} успішно експортована!", "CreateWordFile", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void CreateExcelFile(string tableName)
        {
            Excel.Application application = new Excel.Application();
            application.Visible = false;
            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing);
            try
            {
                List<Dictionary<string, string>> tables = _dataBaseManager.GetFullData($"SELECT * FROM [{tableName}]");
                List<string> columns = _dataBaseManager.GetColumnNameList($"SELECT * FROM [{tableName}]");

                bool column = true;
                int offset = 1;
                for (int i = 0; i < tables.Count; i++)
                {
                    for (int j = 0; j < tables[i].Count; j++)
                    {
                        application.Cells[1, j + 1].Font.Bold = 1;
                        application.Cells[i + offset, j + 1] = (i == 0 && column) ? columns[j] : tables[i].ElementAt(j).Value;
                    }
                    if (column)
                    {
                        column = false;
                        i = -1;
                        offset++;
                    }
                }
            }
            catch (Exception ex)
            {
                application.Quit();
                MessageBox.Show($"З експортуванням таблиці {tableName}, сталася помилка!", "CreateExcelFile", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                workbook.SaveAs(Path.Combine(_path, _excelFolder, $"{tableName}({DateTime.Now.ToString("HH-mm-ss")}).xls"));
                application.Quit();
                MessageBox.Show($"Таблиця {tableName} успішно експортована!", "CreateExcelFile", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void GenerateTable(string tableName, Word.Document document, Word.Paragraph p, int fontSize = 14)
        {
            List<Dictionary<string, string>> tables = _dataBaseManager.GetFullData($"SELECT * FROM [{tableName}]");
            List<string> column = _dataBaseManager.GetColumnNameList($"SELECT * FROM [{tableName}]");
            Word.Table table = document.Tables.Add(p.Range, tables.Count + 1, tables[0].Count);
            table.Borders.Enable = 1;
            foreach (Word.Row row in table.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                {
                    cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cell.Range.Font.Size = fontSize;

                    if (cell.RowIndex == 1)
                    {
                        cell.Range.Text = column[cell.ColumnIndex - 1];
                        cell.Range.Font.Bold = 1;
                    }
                    else
                    {
                        cell.Range.Text = tables[row.Index - 2].ElementAt(cell.Column.Index - 1).Value;
                        cell.Shading.BackgroundPatternColor = ((cell.Column.Index + row.Index) % 2 == 0) ? Word.WdColor.wdColorGray15 : Word.WdColor.wdColorWhite;
                    }

                }
            }
        }

    }
}
