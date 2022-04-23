using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InternetProvider
{
    internal class TabManager
    {
        public struct inputBlanks
        {
            public string columnName;
            public string value;
            public bool enabled;
        }

        DataBaseManager _dataBaseManager = null;
        PresetData _presetData = new PresetData();
        public TabManager(DataBaseManager dataBaseManager)
        {
            _dataBaseManager = dataBaseManager;
        }
        public bool EditMode = false;
        public void FillTabPage(TabPage tabPage)
        {
            string phoneNumber = PresetData.GetPhoneNumber();
            for (int i = 0; i < (tabPage.Controls.Count / 2) - 1; i++)
            {
                Label label = (Label)tabPage.Controls.Find($"text{i}", true).First();
                Control input = tabPage.Controls.Find($"input{i}", true).First();
                if (input.Enabled)
                {
                    try
                    {
                        ComboBox combo = (ComboBox)input;
                        combo.SelectedIndex = PresetData.GetRandomValue(combo.Items.Count);
                    }
                    catch (Exception)
                    {
                        string result = "";
                        switch (label.Text)
                        {
                            case "ПІБ":
                                result = PresetData.GetFullName();
                                break;                            
                            case "Вулиця":
                                result = PresetData.GetStreet();
                                break;                            
                            case "Номер будинку":
                                result = PresetData.GetRandomValue(999, 1).ToString();
                                break;                            
                            case "Номер квартири":
                                result = PresetData.GetRandomValue(100, 1).ToString();
                                break;                                                       
                            case "Номер телефону":
                                result = phoneNumber;
                                break;                            
                            case "Логін":
                                result = $"U{phoneNumber.Substring(phoneNumber.Length - 3)}";
                                break;                            
                            case "Пароль":
                                result = PresetData.GetRandomValue(999999, 100000).ToString(); ;
                                break;
                        }
                        TextBox textBox = (TextBox)input;
                        textBox.Text = result;
                    }
                }
            }
        }

        public TabPage GenerateTabPage(string table, string q = "", string customTitle = "", EventHandler cancel = null, EventHandler confirm = null, List<inputBlanks> inputBlanksList = null)
        {
            EditMode = true;
            TabPage tabPage = new TabPage();
            tabPage.Text = ((customTitle == "") ? $"Додавання елемента до таблиці" : customTitle) + $" \"{table}\"";
            table = table.Split(',')[0];
            q = (q == "") ? $"SELECT * FROM [{table}]" : q;


            int disableIndex = 1;

            int id = _dataBaseManager.GetLastId(table);

            List<string> columnName = _dataBaseManager.GetColumnNameList(q);

            Point startedPoint = new Point(15, 55);
            Point point = startedPoint;
            //int x = 15, y = 100;
            int index = 0;
            columnName.ForEach(item =>
            {
                string value = "";
                bool enabled = true;

                if (inputBlanksList != null)
                    inputBlanksList.ForEach(blank => {
                        if (item.Contains(blank.columnName))
                        {
                            value = blank.value;
                            if (tabPage.Controls.ContainsKey(blank.value))
                            {
                                value = tabPage.Controls.Find(value, true).First().Text;
                            }
                            enabled = blank.enabled;
                        }
                    });

                Label label = new Label();
                label.Location = new Point(point.X, point.Y);
                label.Text = item;
                label.Name = "text" + index;

                label.Width = 250;

                List<string> comboItem = _dataBaseManager.StandartColumnData(item);

                Size inputSize;

                Control inputArea;

                if (comboItem.Count > 0)
                {
                    ComboBox comboBox = (ComboBox)FillElement(new ComboBox(), index++, label);
                    comboBox.Items.AddRange(comboItem.ToArray());
                    comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
                    inputSize = new Size(comboBox.Width, comboBox.Height);
                    int a = 0;
                    int.TryParse(value, out a);
                    comboBox.SelectedIndex = a;
                    inputArea = comboBox;
                }
                else if (item.Contains("Дата")) { 
                    DateTimePicker dateTimePicker = (DateTimePicker)FillElement(new DateTimePicker(), index++, label);
                    dateTimePicker.Format = DateTimePickerFormat.Custom;
                    //
                    dateTimePicker.CustomFormat = "dd/MM/yyyy HH:mm";
                    inputSize = new Size(dateTimePicker.Width, dateTimePicker.Height);
                    inputArea= dateTimePicker;
                }
                else
                {
                    TextBox textBox = (TextBox)FillElement(new TextBox(), index++, label);
                    inputSize = new Size(textBox.Width, textBox.Height);
                    textBox.MaxLength = 32;
                    inputArea = textBox;
                }
                point.Y += label.Size.Height + inputSize.Height + 15;
                if (index == columnName.Count / 2)
                {
                    point.X += inputSize.Width + label.Width;
                    point.Y = startedPoint.Y;
                }
                inputArea.Text = value;
                inputArea.Enabled = enabled;

                if(index == disableIndex)
                {
                    inputArea.Enabled = false;
                }
                if (index == 1 && value == "")
                {
                    inputArea.Text = (id + 1).ToString();
                }
                tabPage.Controls.Add(inputArea);
                tabPage.Controls.Add(label);
            });

            Button buttonConfirm = new Button();
            buttonConfirm.Size = new Size(50, 50);
            buttonConfirm.Location = new Point(point.X + ((250 - (buttonConfirm.Width + 5 + 100)) /2), point.Y + 5);  //new Point(80, 50);
            buttonConfirm.Text = "OK";
            buttonConfirm.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            buttonConfirm.Font = new Font(Button.DefaultFont.FontFamily, 12);

            Button buttonCancel = new Button();
            buttonCancel.Location = new Point(buttonConfirm.Location.X + buttonConfirm.Width + 5, buttonConfirm.Location.Y); //new Point(130, 50);
            buttonCancel.Text = "Скасувати";
            buttonCancel.Size = new Size(100, 50);
            buttonCancel.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            buttonCancel.Font = new Font(Button.DefaultFont.FontFamily, 12);

            buttonConfirm.Click += confirm;
            tabPage.Controls.Add(buttonConfirm);

            buttonCancel.Click += cancel;
            tabPage.Controls.Add(buttonCancel);

            return tabPage;

        }

        private Control FillElement(Control element, int index, Label label, string customName = "input")
        {
            element.Location = new Point(label.Location.X, label.Location.Y + label.Size.Height);
            element.Name = customName + index;
            element.Width = label.Width;

            return element;
        }

    }
}
