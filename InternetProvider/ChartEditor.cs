using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace InternetProvider
{
    internal class ChartEditor
    {
        DataBaseManager _dataBaseManager = null;
        public ChartEditor(DataBaseManager dataBaseManager)
        {
            _dataBaseManager = dataBaseManager;
        }

        public struct MyChart
        {
            public string SelectedItems;
            public string Table;
            public SeriesChartType ChartType;
        }

        public Chart CircleChartUpdate(Chart chart)
        {
            chart.Series.Clear();
            chart.Series.Add("Тарифи");
            chart.Series["Тарифи"].ChartType = SeriesChartType.Pie;
            chart.Series["Тарифи"].IsValueShownAsLabel = true;
            chart.Series["Тарифи"].Font = new Font(FontFamily.GenericSansSerif, 24);
            _dataBaseManager.GetColumnData($"SELECT * FROM [Тарифи]", 1).ForEach(name =>
            {
                chart.Series["Тарифи"].Points.AddXY(name, _dataBaseManager.CountRow("Договір", $"WHERE [Код тарифу] = {_dataBaseManager.GetName(name, "Тарифи", 0, "Найменування")}"));
            });
            return chart;
        }

        public Chart ChartUpdate(Chart chart, string min, string max)
        {
            //SeriesChartType.Area
            List<MyChart> list = new List<MyChart> {
                new MyChart { Table = "Поповнення", SelectedItems = "[Дата],Сума", ChartType = SeriesChartType.Spline },
                 new MyChart { Table = "Витрати клієнтів", SelectedItems = "[Дата Операції],Сума", ChartType = SeriesChartType.Spline } };
            chart.Series.Clear();
            list.ForEach(ch =>
            {
                chart.Series.Add(ch.Table);
                chart.Series[ch.Table].ChartType = ch.ChartType;
                chart.Series[ch.Table].BorderWidth = 7;
                chart.Series[ch.Table].MarkerStyle = MarkerStyle.Square;
                chart.Series[ch.Table].MarkerColor = Color.Black;
                chart.Series[ch.Table].MarkerSize = 7;
                chart.Series[ch.Table].IsValueShownAsLabel = true;
                chart.Series[ch.Table].Font = new Font(FontFamily.GenericSansSerif, 10);

                if (min == "")
                    min = (_dataBaseManager.GetMin("Витрати клієнтів", "Дата операції") != "") ? _dataBaseManager.GetMin("Витрати клієнтів", "Дата операції") : DateTime.Now.Date.ToString("dd/MM/yyyy");
                if (max == "")
                    max = DateTime.Now.Date.ToString("dd/MM/yyyy");
                for (DateTime date = DateTime.Parse(min); date <= DateTime.Parse(max); date = date.AddDays(1))
                {
                    int.TryParse(_dataBaseManager.ExecuteQueryValue($"Select SUM({ch.SelectedItems.Split(',')[1]}) From [{ch.Table}] WHERE {ch.SelectedItems.Split(',')[0]} LIKE '%{date.ToString("dd/MM/yyyy")}%'"), out int sum);
                    if(sum > 0)
                        chart.Series[ch.Table].Points.AddXY(date.ToShortDateString(), sum);
                }
            });
            return chart;
        }
    }
}
