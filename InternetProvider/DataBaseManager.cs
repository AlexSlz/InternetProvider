using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace InternetProvider
{
    internal class DataBaseManager
    {
        public Dictionary<string, string> visuallyRelatedTables = new Dictionary<string, string>()
        {
            { "Код тарифу", "Тарифи" },
            { "Статус підключення", "Статус клієнта" },
            { "Код района", "Райони" },
        };

        public static Dictionary<string, Dictionary<string, Color>> TablePalette = new Dictionary<string, Dictionary<string, Color>>() {
            { "Статус клієнта", new Dictionary<string, Color> {
                            { "Підключено", Color.LightGreen },
                            { "Очікування оплати", Color.LightPink },
                            { "Помилка", Color.LightYellow } } 
            },
            { "Тарифи", new Dictionary<string, Color> {
                            { "Базовий", Color.LightGray },
                            { "Стандарт", Color.LightSeaGreen },
                            { "Преміум", Color.Gold } }
            },
            { "Райони", new Dictionary<string, Color> {
                            { "rand", Color.Black } }
            },
            { "Не фарбувати", new Dictionary<string, Color> {
                            { "none", Color.Black } }
            }
        };
        private string _aceFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "a.accdb");
        
        private OleDbConnection conn = null; // new OleDbConnection(connString);

        private struct relatedBlanks
        {
            public string firstTable;
            public List<string> secondTable;
        }

        private List<relatedBlanks> relatedBlanksList = new List<relatedBlanks>
        {
            new relatedBlanks { firstTable = "Клієнт" , secondTable = new List<string> { "Договір" }},
            new relatedBlanks { firstTable = "Договір" , secondTable = new List<string> { "Поповнення", "Витрати клієнтів" }}
        };

        OpenFileDialog _openFileDialog = new OpenFileDialog();
        public DataBaseManager(OpenFileDialog openFileDialog)
        {
            _openFileDialog = openFileDialog;
            ReOpenConnection();
            /*DataTable table = new OleDbEnumerator().GetElements();
            string inf = "";
            foreach (DataRow row in table.Rows)
            {
                inf += row["SOURCES_NAME"] + "\n";
            }
            MessageBox.Show(inf);*/
        }

        private bool CheckTables()
        {
            try
            {
                string q = @"Select * From
                                ((((([Витрати клієнтів] Inner Join
                                Договір On [Витрати клієнтів].[Номер договору] = Договір.[Номер договору]) Inner Join
                                Клієнт On Договір.[Код клієнта] = Клієнт.[Код клієнта]) Inner Join
                                Поповнення On Поповнення.[Номер договору] = Договір.[Номер договору]) Inner Join
                                Райони On Клієнт.[Код района] = Райони.Код) Inner Join
                                [Статус клієнта] On Клієнт.[Статус підключення] = [Статус клієнта].Код) Inner Join
                                Тарифи On Договір.[Код тарифу] = Тарифи.Код";

                OleDbCommand cmd = new OleDbCommand(q, conn);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show($"Щось пішло не так... \n\n {ex.Message}", "CheckTables", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (dialogResult == DialogResult.OK)
                {
                    System.Environment.Exit(0);
                }
                return false;
            }
            return true;
        }

        public bool DeleteRelatedRows(string table, string index)
        {
            DialogResult dialogResult = MessageBox.Show($"Ви точно хочете видалити елемент з таблиці {table}?", "DeleteRelatedRows", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                DeleteRow(table, GetColumnNameList($"SELECT * FROM [{table}]")[0], index);
                return true;
            }
            return false;
        }

        private void DeleteRow(string table, string columnName, string index)
        {
            foreach (relatedBlanks blanks in relatedBlanksList)
            {
                if (blanks.firstTable == table)
                {
                    blanks.secondTable.ForEach(_table => {
                        DeleteRow(_table, GetColumnNameList($"SELECT * FROM [{table}]")[0], index);
                    });
                }
            }
            if (CountRow(table) > 0)
            {
                string q = $"DELETE FROM [{table}] WHERE [{columnName}]={index}";
                ExecuteQueryBool(q);
            }
        }


        public List<string> GetTableNameList()
        {
            List<string> list = new List<string>();
            DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            foreach (DataRow dr in dt.Rows)
            {
                if (!dr["TABLE_NAME"].ToString().Contains("~"))
                    list.Add(dr["TABLE_NAME"].ToString());
            }
            return list;
        }

        public List<string> GetColumnNameList(string q)
        {
            List<string> list = new List<string>();
            OleDbCommand cmd = new OleDbCommand(q, conn);
            
            OleDbDataReader reader = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
            var table = reader.GetSchemaTable();
            var nameCol = table.Columns["ColumnName"];
            foreach (DataRow row in table.Rows)
            {
                list.Add(row[nameCol].ToString());
            }
            reader.Close();
            return list;
        }

        private Dictionary<string, Color> FillRandColor(Dictionary<string, Color> value, string name, int index = 1)
        {
            value.Clear();
            Random random = new Random();
            int color = 256;
            int minColor = 128;
            GetFullData($"SELECT * FROM [{name}]").ForEach(item => {
                int r = random.Next(minColor += 15, color),
                g = random.Next(minColor += 10, color),
                b = random.Next(minColor -= 25, color);
                value.Add(item.ElementAt(index).Value, Color.FromArgb(r, g, b));
            });
            return value;
        }


        public void FillDataGridView(string q, DataGridView dataGridView, string palette, bool autoCorrect = true)
        {
            Dictionary<string, Color> visuallyColors = TablePalette[palette];
            if (visuallyColors.ContainsKey("rand"))
                visuallyColors = FillRandColor(visuallyColors, palette);
            dataGridView.Columns.Clear();

            OleDbCommand cmd = new OleDbCommand(q, conn);
            DataTable dataTable = new DataTable();
            OleDbDataReader reader = null;
            try
            {
                reader = cmd.ExecuteReader();
            }
            catch {
                return;
            }

            GetColumnNameList(q).ForEach(item => {
                dataGridView.Columns.Add("", item);
            });

            int index = 0;
            while (reader.Read())
            {
                dataGridView.Rows.Add();
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    dataGridView.Rows[index].Cells[i].Value = (autoCorrect) ? GetCustomName(dataGridView.Columns[i].HeaderText, reader.GetValue(i).ToString(), 1) : reader.GetValue(i);
                    if(visuallyColors.ContainsKey(dataGridView.Rows[index].Cells[i].Value.ToString()))
                    dataGridView.Rows[index].DefaultCellStyle.BackColor = visuallyColors[dataGridView.Rows[index].Cells[i].Value.ToString()];
                }
                index++;
            }
            dataGridView.ClearSelection();
            reader.Close();
        }
        public List<TabManager.inputBlanks> GetInputBlanks(string q, List<string> disableName = null)
        {
            List<TabManager.inputBlanks> inputBlanks = new List<TabManager.inputBlanks>();

            Dictionary<string, string> data = GetRowData(q);
            int index = 0;
            GetColumnNameList(q).ForEach(header =>
            {
                bool enabled = true;
                if (disableName != null)
                    disableName.ForEach(item => {
                        if (header.Contains(item)) enabled = false;
                    });
                inputBlanks.Add(new TabManager.inputBlanks { columnName = header, value = GetCustomName(header, data.ElementAt(index++).Value, 1), enabled = enabled }); ;
            });
            return inputBlanks;
        }


        public string GetName(string index, string table, int i = 1, string customName = "Код")
        {
            if (!int.TryParse(index, out int a))
            {
                index = $"'{index}'";
            }
            OleDbCommand cmd = new OleDbCommand($"SELECT * FROM [{table}] WHERE [{customName}] = {index}", conn);
            OleDbDataReader reader = cmd.ExecuteReader();
            reader.Read();
            string value = reader.GetValue(i).ToString();
            reader.Close();
            return value;
        }


        public string GetCustomName(string header, string value, int i = 0)
        {
            foreach (var item in visuallyRelatedTables)
            {
                if (header.Contains(item.Key))
                {
                    List<string> columnList = GetColumnNameList($"SELECT * FROM [{item.Value}]");
                    return GetName(value, item.Value, i, columnList[(i == 0) ? 1 : 0]);
                }
            }
            return value;
        }

        public List<string> StandartColumnData(string table)
        {
            return (visuallyRelatedTables.ContainsKey(table)) ? GetColumnData($"SELECT * FROM [{visuallyRelatedTables[table]}]", 1) : new List<string>();
        }

        public List<string> GetColumnData(string q, int columnIndex)
        {
            List<string> list = new List<string>();
            OleDbCommand cmd = new OleDbCommand(q, conn);
            OleDbDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                list.Add(reader.GetString(columnIndex).ToString());
            }
            reader.Close();
            return list;
        }
        public List<Dictionary<string, string>> GetFullData(string q)
        {
            List<Dictionary<string, string>> list = new List<Dictionary<string, string>>();
            OleDbCommand cmd = new OleDbCommand(q, conn);
            try
            {
                using(OleDbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Dictionary<string, string> temp = new Dictionary<string, string>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            temp.Add(reader.GetName(i), reader.GetValue(i).ToString());
                        }
                        list.Add(temp);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{q}\n\n{ex.Message}", "GetFullData", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return list;
        }

        public Dictionary<string, string> GetRowData(string q)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            List<string> colNames = GetColumnNameList(q);
            OleDbCommand cmd = new OleDbCommand(q, conn);
            OleDbDataReader reader = cmd.ExecuteReader();
            reader.Read();
            int i = 0;
            colNames.ForEach(key =>
            {
                dict.Add(key, reader.GetValue(i++).ToString());
            });
            reader.Close();
            return dict;
        }

        public int CountRow(string table, string addons = "")
        {
            string q = $"SELECT COUNT(*) FROM [{table}] {addons}";
            OleDbCommand command = new OleDbCommand(q, conn);
            return (int)command.ExecuteScalar();
        }

        public string GetMax(string table, string column, string addons = "")
        {
            string q = $"SELECT MAX([{column}]) FROM [{table}] {addons}";
            return ExecuteQueryValue(q);
        }
        public string GetMin(string table, string column, string addons = "")
        {
            string q = $"SELECT MIN([{column}]) FROM [{table}] {addons}";
            return ExecuteQueryValue(q);
        }

        public List<string> GenerateQuery(List<string> tables, TabPage tabPage, bool editMode = false)
        {
            List<List<string>> list = new List<List<string>>();
            tables.ForEach(name =>
            {
                string q = $"SELECT * FROM [{name}]";
                list.Add(GetColumnNameList(q));
            });
            List<string> queryList = new List<string>();
            int tableIndex = 0;
            int index = 0;
            list.ForEach(table =>
            {
                string mainQ = editMode ? $"UPDATE [{tables[tableIndex++]}] SET " : $"INSERT INTO [{tables[tableIndex++]}] VALUES (";

                Dictionary<string, string> idColumn = new Dictionary<string, string>();
                table.ForEach(column =>
                {
                    if (tabPage.Controls.Find($"text{index}", true).First().Text.Contains(column))
                    {
                        Control control = tabPage.Controls.Find($"input{index}", true).First();
                        string value = GetCustomName(column, control.Text, 0);
                        if (!int.TryParse(value, out int a))
                            value = $"'{value}'";
                        mainQ += editMode ? $"[{column}] = {value}" : value;
                        mainQ += ", ";

                        idColumn.Add(column, value);

                    }
                    index++;
                });
                mainQ = mainQ.Remove(mainQ.Length - 2, 2) + (editMode ? $" WHERE [{idColumn.First().Key}] = {idColumn.First().Value}" : ")");
                queryList.Add(mainQ);
            });

            foreach (var q in queryList)
            {
                if (!ExecuteQueryBool(q))
                {
                    return null;
                }
            }

            return queryList;
        }
        public int GetLastId(string table, string addons = "")
        {
            return (CountRow(table) > 0) ? int.Parse(GetMax(table, GetColumnNameList($"SELECT * FROM [{table}]")[0], addons)) : 0;
        }

        public void UpdateUsers(double addTime, bool autoDelete)
        {
            ShowMoneyError = true;
            string q = @"Select Клієнт.[Код клієнта], Договір.[Номер договору], Договір.[Код тарифу], Тарифи.[Ціна на Місяць] From
                                            (Клієнт Inner Join
                                            Договір On Договір.[Код клієнта] = Клієнт.[Код клієнта]) Inner Join
                                            Тарифи On Договір.[Код тарифу] = Тарифи.Код";
            List<Dictionary<string,string>> data = GetFullData(q);
            foreach (Dictionary<string, string> info in data)
            {
                DateTime time = (CountRow("Витрати клієнтів", $"WHERE [Номер договору] = {info["Номер договору"]}") > 0) ? DateTime.Parse(ExecuteQueryValue($"Select [Витрати клієнтів].[Дата наступної оплати] From [Витрати клієнтів] WHERE ([Код] = {GetLastId("Витрати клієнтів", $"WHERE [Номер договору] = {info["Номер договору"]}")}) AND ([Номер договору] = {info["Номер договору"]})")) : DateTime.Now;
                int status = (time > DateTime.Now) ? 1 : 2;
                if (GetMoney(info["Номер договору"]) >= long.Parse(info["Ціна на Місяць"]) && time <= DateTime.Now)
                {
                    ExecuteQueryBool($"INSERT INTO [Витрати клієнтів] VALUES ({GetLastId("Витрати клієнтів") + 1}, {info["Номер договору"]}, '{DateTime.Now.ToString("dd/MM/yyyy HH:mm")}', '{DateTime.Now.AddDays(addTime).ToString("dd/MM/yyyy HH:mm")}', {info["Ціна на Місяць"]})");
                    status = 1;
                }
                else if (GetMoney(info["Номер договору"]) <= 0 && time <= DateTime.Now && autoDelete)
                {
                    DeleteRow("Поповнення", "Номер договору", info["Номер договору"]);
                    DeleteRow("Витрати клієнтів", "Номер договору", info["Номер договору"]);
                }
                if (GetMoney(info["Номер договору"]) <= -1)
                    status = 3;
                ExecuteQueryBool($"UPDATE [Клієнт] SET [Баланс] = {GetMoney(info["Номер договору"])}, [Статус підключення] = {status} WHERE [Код клієнта] = {info["Код клієнта"]}");
            }
        }

        bool ShowMoneyError = true;
        private long GetMoney(string index)
        {
            long sum = 0;
            long spending = 0;
            try
            {
                sum = (CountRow("Поповнення", $"WHERE [Номер договору] = {index}") > 0) ? long.Parse(ExecuteQueryValue($"SELECT SUM(Сума) FROM [Поповнення] WHERE [Номер договору] = {index}")) : 0;
                spending = (CountRow("Витрати клієнтів", $"WHERE [Номер договору] = {index}") > 0) ? long.Parse(ExecuteQueryValue($"SELECT SUM(Сума) FROM [Витрати клієнтів] WHERE [Номер договору] = {index}")) : 0;

            }
            catch (Exception)
            {
                if (ShowMoneyError)
                {
                    DialogResult dialogResult = MessageBox.Show($"У клієнта з номером договору = {index}, сталась помилка! \nВони скинуті до -1!", "GetMoney", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    ShowMoneyError = false;
                }
                return -1;
            }
            return sum - spending;
        }

        public bool ExecuteQueryBool(string q)
        {
            OleDbCommand cmd = new OleDbCommand(q, conn);
            try
            {
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{q}\n\n{ex.Message}", "ExecuteQueryBool", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        public string ExecuteQueryValue(string q)
        {
            OleDbCommand command = new OleDbCommand(q, conn);
            try
            {
                return command.ExecuteScalar().ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{q}\n\n{ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }
        }

        public void connectOtherDataBase()
        {
            try
            {
                _openFileDialog.ShowDialog();
                _aceFilePath = _openFileDialog.FileName;

                ReOpenConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}", "connectOtherDataBase", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ReOpenConnection()
        {
            string connString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_aceFilePath};Persist Security Info=False;";
            conn = new OleDbConnection(connString);
            try
            {
                conn.Open();
                if (CheckTables())
                {
                    MessageBox.Show("База даних успішно підключена!", "ReOpenConnection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                DialogResult dialogResult = MessageBox.Show($"З підключенням до бази даних сталася помилка! \n\n{ex.Message}\n\nСпробувати підключити до іншої бази?", "ReOpenConnection", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                
                if(dialogResult == DialogResult.Yes)
                {
                    connectOtherDataBase();
                }else if(dialogResult == DialogResult.No)
                {
                    System.Environment.Exit(0);
                }
            }
        }
    }
}
