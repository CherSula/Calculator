using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NPOI.HSSF.UserModel; // Для .xls
using NPOI.XSSF.UserModel; // Для .xlsx
using NPOI.SS.UserModel;

namespace Calculator
{
    public partial class Form1 : Form
    {
        private Dictionary<string, double> _indicators = new Dictionary<string, double>();
        private Dictionary<string, double> _prices = new Dictionary<string, double>();
        private List<string> _researchNames = new List<string>();

        public Form1()
        {
            InitializeComponent();

            // Подписываемся на событие CellValueChanged
            dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;

            // Также необходимо подписаться на событие CurrentCellDirtyStateChanged,
            // чтобы обновить значение сразу после изменения ячейки.
            dataGridView1.CurrentCellDirtyStateChanged += (s, e) =>
            {
                if (dataGridView1.IsCurrentCellDirty)
                {
                    dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                }
            };
        }

        private void Form1_Load(object sender, EventArgs e)
        {
#if DEBUG
            ReserchExcells();
#endif
        }

        private void ReserchExcells()
        {
            var researchPath = @"C:\Users\svetl\Desktop\Маша\sheet\analysis-parameter.xlsx";
            var analisysPath = @"C:\Users\svetl\Desktop\Маша\sheet\parameters_lab.cost.xlsx";

            label1.Text = researchPath;
            label2.Text = analisysPath;

            LoadResearch(researchPath);
            LoadPrices(analisysPath);
        }

        private void btnLoadResearch_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    label1.Text = openFileDialog.FileName;
                    LoadResearch(openFileDialog.FileName);
                }
            }
        }


        private void LoadResearch(string filePath)
        {
            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = Path.GetExtension(filePath) == ".xls" ? (IWorkbook)new HSSFWorkbook(file) : new XSSFWorkbook(file);

                for (int i = 0; i < workbook.NumberOfSheets; i++)
                {
                    var sheet = workbook.GetSheetAt(i);
                    string researchName = sheet.SheetName;
                    
                    if (sheet.SheetName == "Все показатели")
                    {
                        continue;
                    }

                    if (!_researchNames.Contains(researchName))
                    {
                        _researchNames.Add(researchName);
                    }

                    for (int row = 0; row <= sheet.LastRowNum; row++)
                    {
                        var cell = sheet.GetRow(row)?.GetCell(0); // Столбец A
                        if (cell != null)
                        {
                            string indicatorName = cell.ToString();

                            if (string.IsNullOrWhiteSpace(indicatorName))
                            {
                                continue;
                            }

                            if (_indicators.ContainsKey(indicatorName))
                            {
                                _indicators[indicatorName] += 1;
                            }
                            else
                            {
                                _indicators[indicatorName] = 1;
                            }
                        }
                    }
                }
            }
        }

        private void btnLoadPrices_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    label2.Text = openFileDialog.FileName;
                    LoadPrices(openFileDialog.FileName);
                }
            }
        }

        private void LoadPrices(string filePath)
        {
            IWorkbook workbook;
            using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = Path.GetExtension(filePath) == ".xls" ? (IWorkbook)new HSSFWorkbook(file) : new XSSFWorkbook(file);

                var sheet = workbook.GetSheetAt(0); // Предполагаем, что данные на первом листе

                for (int row = 1; row <= sheet.LastRowNum; row++) // Пропускаем заголовок
                {
                    var nameCell = sheet.GetRow(row)?.GetCell(0); // Столбец A
                    var priceCellFinal = sheet.GetRow(row)?.GetCell(1); // Столбец D


                    if (nameCell != null && priceCellFinal != null)
                    {
                        string indicatorName = nameCell.ToString();
                        double priceFinal;

                        if (double.TryParse(priceCellFinal.ToString(), out priceFinal))
                        {
                            _prices[indicatorName] = priceFinal;
                        }
                    }
                }
            }
        }


        private void btnShowUniqueIndicators_Click(object sender, EventArgs e)
        {
            MessageBox.Show($"Количество уникальных показателей: {_indicators.Count}");

            dataGridView1.Rows.Clear();

            dataGridView1.Columns.Add("parameter", "Показатель");
            dataGridView1.Columns.Add("count", "Кол-во");
            dataGridView1.Columns.Add("eachCost", "Цена за шт");
            dataGridView1.Columns.Add("coefficient", "Коэффициент");
            dataGridView1.Columns.Add("cost", "Цена показателя всего");

            foreach (var pair in _indicators)
            {
                // Проверяем, существует ли цена для данного показателя
                if (_prices.TryGetValue(pair.Key, out var priceFinal))
                {
                    double eachCost = priceFinal; // Получаем цену за единицу
                    double count = pair.Value; // Количество показателя
                    double coefficient = 1; // Здесь вы можете установить коэффициент по умолчанию или получить его из другого источника
                    double cost = count * coefficient * eachCost; // Расчет стоимости

                    // Добавляем строку с данными
                    dataGridView1.Rows.Add(pair.Key, count, eachCost, coefficient, cost);
                }
                else
                {
                    MessageBox.Show($"Цена для показателя '{pair.Key}' не найдена.");
                }
            }
        }

        // Обработчик события изменения значения ячейки
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // if (измененная ячейка не относится к столбцу coefficient)
            // {
            //     return;
            // }

            // if (значение ячейки не число или меньше нуля)
            // {
            //    присваиваем дефолтное значение
            //    вызываем мессаге бокс и говорим где пользователь ошибся
            // }
            // Проверяем, что измененная ячейка - это столбец коэффициента
            if (e.ColumnIndex == dataGridView1.Columns["coefficient"].Index && e.RowIndex < 0)
            {
                MessageBox.Show("Коэффициент не может быть меньше нуля");
            }
            if (string.IsNullOrWhiteSpace(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString()))
            {
                MessageBox.Show("Коэффициент не может быть пустым. Устанавливается значение по умолчанию: 1.");
                dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 1; // Устанавливаем значение по умолчанию
            }
            if (e.ColumnIndex == dataGridView1.Columns["coefficient"].Index && e.RowIndex >= 0)
            {
                // Получаем новое значение коэффициента
                if (double.TryParse(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), out double newCoefficient))
                {
                    if (string.IsNullOrWhiteSpace(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString()))
                    {
                        MessageBox.Show("Коэффициент не может быть пустым. Устанавливается значение по умолчанию: 1.");
                        dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = 1; // Устанавливаем значение по умолчанию
                    }
                    // Получаем количество и цену за единицу из текущей строки
                    double count = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["count"].Value);
                    double eachCost = Convert.ToDouble(dataGridView1.Rows[e.RowIndex].Cells["eachCost"].Value);

                    // Пересчитываем стоимость
                    double newCost = count * newCoefficient * eachCost;

                    // Обновляем значение стоимости в DataGridView
                    dataGridView1.Rows[e.RowIndex].Cells["cost"].Value = newCost;
                    // Если значение пустое, устанавливаем его обратно на 1 (или любое другое значение по умолчанию)

                }
            }
        }


        private void btnCalculateCost_Click(object sender, EventArgs e)
        {
            //foreach (DataGridViewRow row in dataGridView.Rows)
            //{
            //    string indicatorName = row.Cells[0].Value.ToString();

            //    if (prices.ContainsKey(indicatorName) && double.TryParse(row.Cells[1].Value.ToString(), out double coefficient))
            //    {
            //        double totalCostPerIndicator = prices[indicatorName] * coefficient;
            //        row.Cells[2].Value = totalCostPerIndicator; // Устанавливаем стоимость в третью колонку
            //    }
            //}

            MessageBox.Show("Стоимость рассчитана.");
        }

        private void btnTotalSumForOrder_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Итоговая стоимость заказа - ");
        }
    }
}