using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using GroupBox = System.Windows.Forms.GroupBox;
using Word = Microsoft.Office.Interop.Word;


namespace Accounting
{
    public partial class Form1 : Form
    {
        // Создаем словари для хранения результатов
        private Dictionary<string, (int YesCount, int NoCount)> results =
            new Dictionary<string, (int YesCount, int NoCount)>();

        // Создаем словари для хранения результатов
        private Dictionary<string, (double Positive, double Negative)> resultsPercentage =
            new Dictionary<string, (double Positive, double Negative)>();
       
        private Dictionary<string, int> resultIndex = new Dictionary<string, int>();

        private Dictionary<string, bool> radioButtonResults = new Dictionary<string, bool>();





        public Form1()
        {
            InitializeComponent();

            // Отключение вкладки при загрузке формы
            tabControl1.Selecting += tabControl1_Selecting;


        }


        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage == tabPage9) // Если пытаемся выбрать tabPage3
            {
                e.Cancel = true; // Отменяем выбор
                MessageBox.Show("Для перехода в 'Результаты' во вкладке 'Меню' сформируйте отчет в программе.\n" +
                    "Также все ответы должны быть получены!");
            }
        }



        private void ExportExcel()
        {
            // Создаем приложение Word
            Word.Application wordApp = new Word.Application();
            

            // Создаем новый документ
            Word.Document doc = wordApp.Documents.Add();

            // Определяем количество строк и столбцов
            int rowCount = tabControl1.TabCount + 1; 
            int columnCount = 4; // 4 столбца

            // Создаем таблицу
            Word.Table table = doc.Tables.Add(doc.Range(), rowCount, columnCount);
            table.Borders.Enable = 1; // Включаем границы в таблице

            table.Cell(2, 2).Range.Text = "80 и выше (система эффективная, требует контроля и минимальных улучшений)";
            table.Cell(2, 3).Range.Text = "75-80 (система в целом эффективна, требуются корректировки по отдельным разделам работ)";
            table.Cell(2, 4).Range.Text = "75 и ниже (система неэффективна, требуются существенные изменения)";

            table.Cell(3, 1).Range.Text = "№ 2 Активное выявление, учет и регистрация, анализ ИСМП среди пациентов и персонала"; // Пример данных для первого столбца
            table.Cell(4, 1).Range.Text = "№ 3 Организация микробиологических исследований (включая случай подозрения ИСМП)"; // Пример данных для второго столбца
            table.Cell(5, 1).Range.Text = "№ 5 Организация стерилизации"; // Пример данных для первого столбца
            table.Cell(6, 1).Range.Text = "№ 6 Обеспечение эпидемиологической безопасности среды"; // Пример данных для второго столбца
            table.Cell(7, 1).Range.Text = "№ 7 Обеспечение эпидемиологической безопасности медицинских технологий"; // Пример данных для третьего столбца
            table.Cell(8, 1).Range.Text = "№ 8 Порядок оказания помощи пациентам, требующим изоляции (с инфекциями, передающимися воздушно-капельным путем, опасными инфекциями)"; // Пример данных для четвертого столбца
            table.Cell(9, 1).Range.Text = "№ 9 Наличие полностью оборудованных мест для мытья и обработки рук"; // Пример данных для четвертого столбца
            table.Cell(10, 1).Range.Text = "№ 10 Соблюдение правил гигиены рук персоналам"; // Пример данных для четвертого столбца
            table.Cell(11, 1).Range.Text = "№ 11 Соблюдение персоналом алгоритма использования индивидуальных средств защиты"; // Пример данных для четвертого столбца
            table.Cell(12, 1).Range.Text = "№ 12 Профилактика ИСМП у медицинского персонала"; // Пример данных для четвертого столбца
            
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= columnCount; j++)
                {
                    table.Cell(i, j).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;      
                }
                
            }

            // Объединяем все три заголовка в одну ячейку
            var cellToMerge = table.Cell(1, 2);
            cellToMerge.Merge(table.Cell(1, 3)); // Сначала объединяем ячейки 1 и 2
            cellToMerge.Merge(table.Cell(1, 3)); // Затем объединяем с 3

            // Устанавливаем текст и выравнивание
            cellToMerge.Range.Text = "Оценка показателя, % положительных ответов";

            // Объединяем все три заголовка в одну ячейку
            var cellToMergeRow = table.Cell(1, 1);
            cellToMergeRow.Merge(table.Cell(2, 1)); // Сначала объединяем ячейки 1 и 2
            cellToMergeRow.Range.Text = "Показатель";



            foreach (var kvp in resultsPercentage)
            {
                for (var i = 3; i <= rowCount; i++)
                {
                    if (String.Equals(table.Cell(i, 1).Range.Text.TrimEnd('\r', '\a').ToString(), kvp.Key.ToString()))
                    {
                        table.Cell(i, resultIndex[kvp.Key]).Range.Text = kvp.Value.Positive.ToString("F2") + " %";
                    }
                    
                }
            }

            wordApp.Visible = true;





            // Освобождение ресурсов
            Marshal.ReleaseComObject(table);
            Marshal.ReleaseComObject(doc);
            Marshal.ReleaseComObject(wordApp);
        }

       

        private void CountResponses()
        {
            // Проходим по всем вкладкам в TabControl
            foreach (TabPage tabPage in tabControl1.TabPages)
            {
                // Проходим по всем GroupBox на каждой TabPage
                foreach (GroupBox groupBox in tabPage.Controls.OfType<GroupBox>())
                {
                    int positiveCount = 0; // Количество "Да" для текущего GroupBox
                    int negativeCount = 0; // Количество "Нет" для текущего GroupBox

                    // Извлекаем TableLayoutPanel из GroupBox
                    foreach (TableLayoutPanel tableLayoutPanel in groupBox.Controls.OfType<TableLayoutPanel>())
                    {
                        // Проходим по каждому Panel в TableLayoutPanel
                        foreach (Panel panel in tableLayoutPanel.Controls.OfType<Panel>())
                        {
                            // Ищем RadioButton с "Да" и "Нет"
                            RadioButton yesRadioButton = panel.Controls.OfType<RadioButton>().FirstOrDefault(rb => rb.Text == "Да");
                            RadioButton noRadioButton = panel.Controls.OfType<RadioButton>().FirstOrDefault(rb => rb.Text == "Нет");

                            // Увеличиваем счетчики на основе выбора
                            if (yesRadioButton != null && yesRadioButton.Checked)
                            {
                                positiveCount++;
                            }
                            else if (noRadioButton != null && noRadioButton.Checked)
                            {
                                negativeCount++;
                            }
                        }
                    }

                    // Сохраняем результаты для текущего GroupBox
                    results[groupBox.Text] = (positiveCount, negativeCount);


                    // Вычисляем процент позитивных и негативных ответов
                    int totalQuestions = positiveCount + negativeCount;
                    double positivePercentage = (totalQuestions > 0) ? (positiveCount / (double)totalQuestions) * 100 : 0;
                    double negativePercentage = (totalQuestions > 0) ? ((totalQuestions - positiveCount) / (double)totalQuestions) * 100 : 0;
                    resultsPercentage[groupBox.Text] = (positivePercentage, negativePercentage);

                    if (resultsPercentage[groupBox.Text].Positive >= 80)
                    {
                        resultIndex[groupBox.Text] = 2;
                    }
                    else if (resultsPercentage[groupBox.Text].Positive > 75 && resultsPercentage[groupBox.Text].Positive < 80)
                    {
                        resultIndex[groupBox.Text] = 3;
                    }
                    else
                    {
                        resultIndex[groupBox.Text] = 4;
                    }

                }
            }

            


        }
        private bool CheckRadioButton()
        {
            foreach (TabPage tabPage in tabControl1.TabPages) // Предполагаем, что ваш TabControl называется tabControl1
            {
                var groupBox = tabPage.Controls.OfType<GroupBox>().FirstOrDefault();
                if (groupBox != null)
                {
                    var tableLayoutPanel = groupBox.Controls.OfType<TableLayoutPanel>().FirstOrDefault();
                    if (tableLayoutPanel != null)
                    {
                        bool isAnyRadioButtonChecked = false;

                        // Проверяем все панели в TableLayoutPanel
                        foreach (Panel panel in tableLayoutPanel.Controls.OfType<Panel>())
                        {
                            var radioButtons = panel.Controls.OfType<RadioButton>();
                            if (radioButtons.Any(rb => rb.Checked))
                            {
                                isAnyRadioButtonChecked = true;
                            }
                        }

                        // Если ни одна радиокнопка не выбрана в данной вкладке, возвращаем false
                        if (!isAnyRadioButtonChecked)
                        {
                            return false;
                        }
                    }
                }
            }

            return true; // Если все радиокнопки выбраны
        
        }


        private void сформироватьОтчетWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
               
                if (!CheckRadioButton())
                {
                    MessageBox.Show("Похоже что не все кнопки отмечены!\nПроверьте, что вы ответили на все вопросы и повторите попытку :)");
                    return;
                }
                CountResponses();
                ExportExcel();
            }
            catch
            {
                MessageBox.Show("Ой, что-то пошло не так!\n" +
                                $"Но мы все равно сохранили данные для вас.\n" +
                                $"Можете с ними ознакомиться на вкладке 'Результаты'");

                setDataGridView();
                tabControl1.Selecting -= tabControl1_Selecting;
                tabControl1.SelectedTab = tabControl1.TabPages[10];
            }
            
        }
        private void setDataGridView()
        {
            dataGridView1.Rows.Clear(); // Очищает строки, если они есть

            dataGridView1.Rows.Add("№ 2 Активное выявление, учет и регистрация, анализ ИСМП среди пациентов и персонала"); // Пример данных для первого столбца
            dataGridView1.Rows.Add("№ 3 Организация микробиологических исследований (включая случай подозрения ИСМП)");
            dataGridView1.Rows.Add("№ 5 Организация стерилизации"); // Пример данных для первого столбца
            dataGridView1.Rows.Add("№ 6 Обеспечение эпидемиологической безопасности среды"); // Пример данных для второго столбца
            dataGridView1.Rows.Add("№ 7 Обеспечение эпидемиологической безопасности медицинских технологий"); // Пример данных для третьего столбца
            dataGridView1.Rows.Add("№ 8 Порядок оказания помощи пациентам, требующим изоляции (с инфекциями, передающимися воздушно-капельным путем, опасными инфекциями)"); // Пример данных для четвертого столбца
            dataGridView1.Rows.Add("№ 9 Наличие полностью оборудованных мест для мытья и обработки рук"); // Пример данных для четвертого столбца
            dataGridView1.Rows.Add("№ 10 Соблюдение правил гигиены рук персоналам"); // Пример данных для четвертого столбца
            dataGridView1.Rows.Add("№ 11 Соблюдение персоналом алгоритма использования индивидуальных средств защиты"); // Пример данных для четвертого столбца
            dataGridView1.Rows.Add("№ 12 Профилактика ИСМП у медицинского персонала"); // Пример данных для четвертого столбца


            for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                // Разрешаем перенос текста в столбце "Описание"
                dataGridView1.Rows[i].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            }

            foreach (var kvp in resultsPercentage)
            {
                for (var i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    

                    if (String.Equals(dataGridView1.Rows[i].Cells[0].Value.ToString(), kvp.Key.ToString()))
                    {
                        dataGridView1.Rows[i].Cells[resultIndex[kvp.Key]-1].Value = kvp.Value.Positive.ToString("F2") + " %";
                    }

                }
            }

            // Центрируем заголовки столбцов
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter; // Центрируем заголовок
                column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter; // Центрируем содержимое ячеек
            }




        }
        private void ExportDataGridViewToExcel()
        {

            // Создаем экземпляр Excel
            var excelApp = new Excel.Application();
            

            // Создаем рабочую книгу и рабочий лист
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Заголовки столбцов
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1].ColumnWidth = 30;
                
                worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText; // Заголовок
                worksheet.Cells[1, i + 1].WrapText = true;
                


            }

            // Данные
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {

                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    

                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value; // Данные
                   
                    worksheet.Cells[i + 2, j + 1].WrapText = true;
                   
                }
            }

            // Центрируем данные в ячейках
            Excel.Range usedRange = worksheet.UsedRange;
            usedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            usedRange.VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;



            // Определение диапазона для обводки
            Excel.Range range = worksheet.Range["A1",
                worksheet.Cells[dataGridView1.Rows.Count + 1, dataGridView1.Columns.Count]];

            // Установка стиля обводки
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.Weight = Excel.XlBorderWeight.xlThin;

            // Сохранение книги (при необходимости)
            // workbook.SaveAs("C:\\path\\to\\your\\file.xlsx");
            excelApp.Visible = true; // Сделаем приложение видимым


            // Освобождение ресурсов
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        private void groupBox7_Enter(object sender, EventArgs e)
        {

        }

        private void показатьРезульатыВПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {


            try
            {
                if (!CheckRadioButton())
                {
                    MessageBox.Show("Похоже что не все кнопки отмечены!\nПроверьте, что вы ответили на все вопросы и повторите попытку :)");
                    return;
                }


                tabControl1.Selecting -= tabControl1_Selecting;

                CountResponses();
                setDataGridView();
                // Переход к TabPage с индексом 1 (второй TabPage)
                tabControl1.SelectedTab = tabControl1.TabPages[10];
            }
            catch
            {

            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void button1_Click_3(object sender, EventArgs e)
        {
            try
            {
                ExportDataGridViewToExcel();
            }
            catch {
                MessageBox.Show("Ой, что-то пошло не так!\n" +
                    "Но вы всегда можете ознакомиться с результатами на вкладке 'Результаты'.");
            }
            
            
            
        }
        //private void CheckIndicator()
        //{
        //    // Проход по всем TabPages в TabControl
        //    foreach (TabPage tabPage in tabControl1.TabPages)
        //    {
        //        // Проход по всем GroupBox в текущем TabPage
        //        foreach (GroupBox groupBox in tabPage.Controls.OfType<GroupBox>())
        //        {
        //            // Получаем Panel, в которой находятся RadioButton
        //            TableLayoutPanel layoutPanel = groupBox.Controls.OfType<TableLayoutPanel>().FirstOrDefault();
        //            if (layoutPanel != null)
        //            {
        //                // Получаем все панели с RadioButton
        //                bool allSelected = true; // Переменная для проверки

        //                foreach (Panel panel in layoutPanel.Controls.OfType<Panel>())
        //                {
        //                    // Проверяем, есть ли выбранный RadioButton в текущей панели
        //                    var selectedRadioButton = panel.Controls.OfType<RadioButton>().FirstOrDefault(r => r.Checked);

        //                    if (selectedRadioButton == null)
        //                    {
        //                        allSelected = false; // Если ни один не выбран, то false
        //                        break; // Выход из цикла
        //                    }
        //                }

        //                // Добавляем результат в словарь
        //                radioButtonResults[groupBox.Text] = allSelected;
        //            }
        //        }
        //    }


        //}

        //private void поПоказателю5ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        CheckIndicator();
        //        if (!radioButtonResults["№ 5 Организация стерилизации"])
        //        {
        //            MessageBox.Show("Похоже вы ответили не на все вопросы показателя № 5!\n Проверьте, что вы ответили на все вопросы и повторите попытку :)");
        //            return;
        //        }
        //        CountResponses();

        //        MessageBox.Show("% позитивных ответов по показателю № 5: " + resultsPercentage["№ 5 Организация стерилизации"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
            
        //}

        //private void поПоказателю6ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

               
        //        CountResponses();
        //        MessageBox.Show("% позитивных ответов по показателю № 6: " + resultsPercentage["№ 6 Обеспечение эпидемиологической безопасности среды"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
            

        //}

        //private void поПоказателю7ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        CountResponses();
        //        MessageBox.Show("% позитивных ответов по показателю № 7: " + resultsPercentage["№ 7 Обеспечение эпидемиологической безопасности медицинских технологий"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
        //}

        //private void поПоказателю8ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        CountResponses();
        //        MessageBox.Show("% позитивных ответов по показателю № 8: " + resultsPercentage["№ 8 Порядок оказания помощи пациентам, требующим изоляции (с инфекциями, передающимися воздушно-капельным путем, опасными инфекциями)"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
        //}

        //private void поПоказателю9ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        CountResponses();
        //        MessageBox.Show("% позитивных ответов по показателю № 9: " + resultsPercentage["№ 9 Наличие полностью оборудованных мест для мытья и обработки рук"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
        //}

        //private void поПоказателю10ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        CountResponses();
        //        MessageBox.Show("% позитивных ответов по показателю № 10: " + resultsPercentage["№ 10 Соблюдение правил гигиены рук персоналам"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
        //}

        //private void поПоказателю11ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        CountResponses();
        //        MessageBox.Show("% позитивных ответов по показателю № 11: " + resultsPercentage["№ 11 Соблюдение персоналом алгоритма использования индивидуальных средств защиты"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
        //}

        //private void поПоказателю12ToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //    try
        //    {

        //        CountResponses();
        //        MessageBox.Show("% позитивных ответов по показателю № 12: " + resultsPercentage["№ 12 Профилактика ИСМП у медицинского персонала"].Positive.ToString() + " %");
        //    }
        //    catch
        //    {

        //    }
        //}
    }
    
    
}
