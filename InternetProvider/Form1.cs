using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace InternetProvider
{
    public partial class Form1 : Form
    {
        private static string _path = AppDomain.CurrentDomain.BaseDirectory;
        Dictionary<string, string> queryBlanks = new Dictionary<string, string>()
        {
            { "Клієнт", @"Select
                            Клієнт.[Код клієнта],
                            Договір.[Номер договору],
                            Клієнт.ПІБ,
                            Клієнт.[Код района],
                            Клієнт.Баланс,
                            Клієнт.[Статус підключення],
                            Договір.[Код тарифу],
                            Тарифи.[Ціна на Місяць]
                        From
                                (Договір Inner Join
                                Клієнт On Договір.[Код клієнта] = Клієнт.[Код клієнта]) Inner Join
                                Тарифи On Договір.[Код тарифу] = Тарифи.Код"
            }
        };

        List<Control> disableWhenQuery = new List<Control>();

        static DataBaseManager _dataBaseManager;
        static ChartEditor _chartEditor;
        static TabManager _tabManager;
        static OfficeManager _officeManager;
        public Form1()
        {

/*            CultureInfo newCulture = (CultureInfo)System.Threading.Thread.CurrentThread.CurrentCulture.Clone();
            newCulture.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            newCulture.DateTimeFormat.DateSeparator = "/";
            Thread.CurrentThread.CurrentCulture = newCulture;*/

            InitializeComponent();

            _dataBaseManager = new DataBaseManager(openFileDialog1);
            _chartEditor = new ChartEditor(_dataBaseManager);
            _tabManager = new TabManager(_dataBaseManager);
            _officeManager = new OfficeManager(_dataBaseManager, _path);

            disableWhenQuery.Add(panel4);
            disableWhenQuery.Add(button7);
            disableWhenQuery.Add(button6);
            disableWhenQuery.Add(button8);
            disableWhenQuery.Add(button9);
            disableWhenQuery.Add(button10);

            comboBox1.Items.AddRange(_dataBaseManager.GetTableNameList().ToArray());
            comboBox5.Items.AddRange(DataBaseManager.TablePalette.ToList().Select(i => i.Key).ToArray());
            comboBox5.SelectedIndex = 0; //comboBox5.Items.Count - 1;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            UpdateSearchBox();
        }
        void DataGridFill(DataGridView dataGrid, Label label, string tableName, string q = "")
        {
            if (String.IsNullOrEmpty(q))
            {
                q = $"SELECT * FROM [{tableName}]";
                if (queryBlanks.ContainsKey(tableName))
                {
                    q = queryBlanks[tableName];
                }
            }
            _dataBaseManager.FillDataGridView(q, dataGrid, comboBox5.Text, checkBox1.Checked);
            label.Text = $"{tableName} | Кількість записів: {dataGrid.RowCount}";
        }

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            button4.Visible = true;
            label3.Text = "Повна інформація про клієнта:\n\n";

            try
            {
                string q = $"SELECT * FROM [Клієнт] WHERE [Код клієнта] = {dataGridView1.Rows[e.RowIndex].Cells[0].Value}";
                DrawInformation(q, label3);
                label3.Text += "\nДоговір: \n\n";
                q = $"SELECT * FROM [Договір] WHERE [Код клієнта] = {dataGridView1.Rows[e.RowIndex].Cells[0].Value}";
                DrawInformation(q, label3);
                label3.Text += "\n";

                DataGridFill(dataGridView2, label2, "", $"SELECT * FROM [{comboBox2.Text}] WHERE [Номер договору] = {dataGridView1.Rows[e.RowIndex].Cells[1].Value}");
            }
            catch { }
        }

        private void DrawInformation(string q, Label label)
        {
            List<string> data = _dataBaseManager.GetRowData(q);
            int index = 0;
            _dataBaseManager.GetColumnNameList(q).ForEach(header =>
            {
                label.Text += $"{header}: {_dataBaseManager.GetCustomName(header, data[index++], 1)}\n";
            });
        }

        private void CreateTabPage(string tableName, string q, string customName = "", List<TabManager.inputBlanks> _inputBlanks = null, EventHandler OkEvent = null)
        {
            tabControl1.TabPages.Add(_tabManager.GenerateTabPage(tableName, q, customName, CancelMethod, OkEvent, _inputBlanks));
            tabControl1.SelectedIndex = tabControl1.TabCount - 1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<TabManager.inputBlanks> _inputBlanks = new List<TabManager.inputBlanks>()
            {
                new TabManager.inputBlanks { columnName = "Баланс", value = "1000", enabled = false },
                new TabManager.inputBlanks { columnName = "Статус підключення", value = "1", enabled = false },
                new TabManager.inputBlanks { columnName = "Номер договору", value = "input0", enabled = false },
                new TabManager.inputBlanks { columnName = "Договір.Код клієнта", value = "input0", enabled = false },
                new TabManager.inputBlanks { columnName = "Дата укладення договору", value = "", enabled = false },
                new TabManager.inputBlanks { columnName = "Логін", value = "U", enabled = true }
            };
            if (tabControl1.SelectedTab.Text == "Головна сторінка" && TableValidation(dataGridView1, true))
            {
                CreateTabPage("Клієнт,Договір", @"Select * From Клієнт Inner Join Договір On Договір.[Код клієнта] = Клієнт.[Код клієнта]",
                    "Укладання нового договору", _inputBlanks, AddMethod);
                if (checkBox2.Checked)
                    _tabManager.FillTabPage(tabControl1.SelectedTab);
            }
            else if (TableValidation(dataGridView3, true))
            {
                CreateTabPage(comboBox1.Text, "", "", null, AddMethod);
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 1)
            {

                List<TabManager.inputBlanks> _inputBlanks = new List<TabManager.inputBlanks>()
            {
                new TabManager.inputBlanks { columnName = "Номер договору", value = dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value.ToString(), enabled = false },
                new TabManager.inputBlanks { columnName = "Дата", value = "", enabled = false },
            };
                if (TableValidation(dataGridView2, true))
                    CreateTabPage("Поповнення", "", "Поповнення балансу клієнта", _inputBlanks, AddMethod);
            }
        }

        private void AddMethod(object s, EventArgs e)
        {
            List<string> list = tabControl1.SelectedTab.Text.Split('"')[1].Split(',').ToList();
            List<string> result = _dataBaseManager.GenerateQuery(list, tabControl1.SelectedTab);
            if (result != null)
            {
                Refresh();
                tabControl1.TabPages.Remove(tabControl1.SelectedTab);
            }
        }

        private void EditMethod(object s, EventArgs e)
        {
            List<string> list = tabControl1.SelectedTab.Text.Split('"')[1].Split(',').ToList();
            List<string> result = _dataBaseManager.GenerateQuery(list, tabControl1.SelectedTab, true);
            if (result != null)
            {
                Refresh();
                tabControl1.TabPages.Remove(tabControl1.SelectedTab);
            }
        }

        private void CancelMethod(object s, EventArgs e)
        {
            _tabManager.EditMode = false;
            tabControl1.TabPages.RemoveAt(tabControl1.TabCount - 1);
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            disableWhenQuery.ForEach(item =>
            {
                item.Enabled = !_tabManager.EditMode && tabControl1.SelectedTab.Text != "Статистика";
            });
            UpdateSearchBox();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab.Text == "Головна сторінка" && TableValidation(dataGridView1, false) && dataGridView2.SelectedCells.Count <= 0)
            {
                string q = $"Select * From [Клієнт] Inner Join [Договір] On Договір.[Код клієнта] = Клієнт.[Код клієнта] WHERE Клієнт.[Код клієнта] = {dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value}";
                List<TabManager.inputBlanks> _inputBlanks = _dataBaseManager.GetInputBlanks(q, new List<string>() { "Код клієнта", "Номер договору" });

                CreateTabPage("Клієнт,Договір", q, "Редагування договору", _inputBlanks, EditMethod);
            }
            else if (tabControl1.SelectedTab.Text == "Налаштування" && TableValidation(dataGridView3, false))
            {
                string q = $"Select * From [{comboBox1.Text}] WHERE [{dataGridView3.Columns[0].HeaderText}] = {dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[0].Value}";
                List<TabManager.inputBlanks> _inputBlanks = _dataBaseManager.GetInputBlanks(q);

                CreateTabPage(comboBox1.Text, "", "Редагування елемента в таблиці", _inputBlanks, EditMethod);
            } else if (TableValidation(dataGridView2, false))
            {
                string q = $"Select * From [{comboBox2.Text}] WHERE [{dataGridView2.Columns[0].HeaderText}] = {dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[0].Value}";
                List<TabManager.inputBlanks> _inputBlanks = _dataBaseManager.GetInputBlanks(q);

                CreateTabPage(comboBox2.Text, "", "Редагування елемента в таблиці", _inputBlanks, EditMethod);
            }
        }

        private bool TableValidation(DataGridView dataGridView, bool checkZeroRow)
        {
            bool result = (dataGridView.RowCount > 0) ? dataGridView.CurrentRow.Index < dataGridView.RowCount : checkZeroRow;

            if (!result)
            {
                MessageBox.Show("Вже йде додавання/редагування таблиці.\nАбо сталася помилка.", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return result;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.Visible = false;
            dataGridView1.ClearSelection();
            DataGridFill(dataGridView2, label2, "", $"SELECT * FROM [{comboBox2.Text}]");
            label3.Text = "Повна інформація про клієнта:\n\n";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SelectBoxManager(comboBox1))
            {
                UpdateSearchBox();
                Refresh();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            disableWhenQuery.ForEach(item =>
            {
                item.Enabled = false;
            });
            Refresh();
            disableWhenQuery.ForEach(item =>
            {
                item.Enabled = true;
            });
        }
        private void Refresh()
        {
            _tabManager.EditMode = false;
            textBox1.Text = "";
            if (checkBox3.Checked)
                _dataBaseManager.UpdateUsers((double)numericUpDown1.Value, checkBox4.Checked);
            DataGridFill(dataGridView1, label1, "Клієнт");
            DataGridFill(dataGridView2, label2, "", $"SELECT * FROM [{comboBox2.Text}]");
            DataGridFill(dataGridView3, label4, "", $"SELECT * FROM [{comboBox1.Text}]");//comboBox1.Text
            button4.Visible = false;
            UpdateDatePicker();
            _chartEditor.CircleChartUpdate(chart1);
        }
        private void UpdateDatePicker()
        {
            List<DateTimePicker> timePickers = new List<DateTimePicker> { dateTimePicker1, dateTimePicker2 };
            foreach (var item in timePickers)
            {
                item.Format = DateTimePickerFormat.Custom;
                item.CustomFormat = "dd/MM/yyyy"; // HH:mm
            }
            dateTimePicker1.Text = _dataBaseManager.GetMin("Витрати клієнтів", "Дата операції");
            dateTimePicker2.Text = DateTime.Now.ToString();
        }

        private void DateTimePickerChanged(object sender, EventArgs e)
        {
            _chartEditor.ChartUpdate(chart2, dateTimePicker1.Text, dateTimePicker2.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            bool update = false;
            if (tabControl1.SelectedTab.Text == "Головна сторінка" && TableValidation(dataGridView1, false) && dataGridView2.SelectedCells.Count <= 0)
            {
                update = _dataBaseManager.DeleteRelatedRows("Клієнт", dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value.ToString());
            }
            else if (tabControl1.SelectedTab.Text == "Налаштування" && TableValidation(dataGridView3, false))
            {
                update = _dataBaseManager.DeleteRelatedRows(comboBox1.Text, dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[0].Value.ToString());
            }else if(TableValidation(dataGridView2, false))
            {
                update = _dataBaseManager.DeleteRelatedRows(comboBox2.Text, dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[0].Value.ToString());
            }
            if (update)
                Refresh();
        }
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (SelectBoxManager(comboBox2))
            {
                if (dataGridView1.SelectedCells.Count == 1)
                    DataGridFill(dataGridView2, label2, "", $"SELECT * FROM [{comboBox2.Text}] WHERE [Номер договору] = {dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[1].Value}");
                else
                    Refresh();
            }
        }

        private void UpdateSearchBox()
        {
            comboBox3.Items.Clear();
            comboBox3.Items.AddRange(_dataBaseManager.GetColumnNameList((tabControl1.SelectedIndex == 0) ? queryBlanks["Клієнт"] : $"SELECT * FROM [{comboBox1.Text}]").ToArray());
            comboBox3.SelectedIndex = 0;
        }

        private string GetTableName(string id)
        {
            Dictionary<string, string> tableNameBlanks = new Dictionary<string, string>()
            {
                { "Код клієнта", "Клієнт" },
            };
            foreach (var item in tableNameBlanks)
            {
                if (item.Key.Contains(id))
                    return item.Value + ".";
            }
            return "";
        }

        private void Search()
        {
            if (textBox1.Text != "")
            {
                string symbol = comboBox4.Text;
                string value = textBox1.Text;
                if (!int.TryParse(value, out int a))
                {
                    value = $"'%{value}%'";
                    symbol = ((comboBox4.Text == "!=") ? "NOT " : "") + "LIKE";
                }
                symbol = (comboBox4.Text == "!=") ? "<>" : symbol;
                string q = ((tabControl1.SelectedIndex == 0) ? queryBlanks["Клієнт"] : $"SELECT * FROM [{comboBox1.Text}]") + $" WHERE {GetTableName(comboBox3.Text)}[{comboBox3.Text}] {symbol} {value}";
                if (tabControl1.SelectedIndex == 0)
                    DataGridFill(dataGridView1, label1, "Клієнт", q);
                else
                    DataGridFill(dataGridView3, label4, "", q);
            }
            else
            {
                Refresh();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            Search();
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && SelectBoxManager(comboBox4))
            {
                Search();
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (textBox1.Text != "" && SelectBoxManager(comboBox3))
                Search();
        }

        Dictionary<string, string> comboSelectIndex = new Dictionary<string, string>();
        int errorCount = 0;
        private bool SelectBoxManager(ComboBox comboBox)
        {
            if (errorCount-- > 0)
                return false;
            if(!comboSelectIndex.ContainsKey(comboBox.Name))
            {
                comboSelectIndex.Add(comboBox.Name, comboBox.SelectedIndex + "");
            }
            else
            {
                if (comboSelectIndex[comboBox.Name].Contains(comboBox.SelectedIndex + ""))
                {
                    return false;
                }
                comboSelectIndex[comboBox.Name] = comboBox.SelectedIndex + "";
            }

            return true;

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if(TableValidation(dataGridView1, true))
            {
                Dictionary<string, string> clientInfo = _dataBaseManager.GetRowDataDic($"Select * From (Клієнт Inner Join Договір On Договір.[Код клієнта] = Клієнт.[Код клієнта]) Inner Join Тарифи On Договір.[Код тарифу] = Тарифи.Код WHERE Клієнт.[Код клієнта] = {dataGridView1.Rows[dataGridView1.CurrentRow.Index].Cells[0].Value}");
                Dictionary<string, string> d = new Dictionary<string, string>()
                        {
                            { "{dId}", clientInfo["Номер договору"] },
                            { "{date}", DateTime.Now.ToString("dd/MM/yyyy")  },
                            { "{fullName}", clientInfo["ПІБ"]  },
                            { "{dataTable}", "Тарифи"  },
                            { "{address}", $"{_dataBaseManager.GetName(clientInfo["Код района"], "Райони")} вул. {clientInfo["Вулиця"]} буд. {clientInfo["Номер будинку"]} кв. {clientInfo["Номер квартири"]}"  },
                            { "{phone}", clientInfo["Номер телефону"]  },
                            { "{login}", clientInfo["Логін"]  },
                            { "{password}", clientInfo["Пароль"]  },
                        };
                _officeManager.FillBlankFile(d);
            }
        }
        
        private void button8_Click(object sender, EventArgs e)
        {
            _officeManager.CreateWordFile(comboBox1.Text);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            _officeManager.CreateExcelFile(comboBox1.Text);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            _dataBaseManager.connectOtherDataBase();
        }
    }
}
