using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;  //чтобы не возникало конфликта, создаем псевдоним
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public string fileName = "";
        public string selectedName="";

        public Form1()
        {
            InitializeComponent();
            processing init = new processing();
            init.initGrid(dataGridView1);
            init.initData(dataGridView1, label1, comboBox1, ref fileName);
        }

        private void button1_Click(object sender, EventArgs e)
        {
                timer1.Enabled = true;
                timer1.Start();
        }  //Начать

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Enabled = false;
            Close();
        }  //Закрыть

        private void timer1_Tick(object sender, EventArgs e)
        {
            processing trafTake = new processing();
            trafTake.outputScreenCurrentTraffic(dataGridView1, label1, fileName);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Enabled = false; 
            processing initPath = new processing();
            fileName = initPath.loadPath(label1, fileName);
            if (fileName!="")
            {
                initPath.initData(dataGridView1, label1, comboBox1, ref fileName);
            }
        }  //Файл

        private void button4_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            timer1.Enabled = false;
        }  //Стоп

        private void button6_Click(object sender, EventArgs e)
        {
            processing outputRezName = new processing();
            outputRezName.sortName(fileName);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedName = comboBox1.SelectedItem.ToString();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (selectedName == "") MessageBox.Show("Выберите процесс!", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Question);
            else
            {
                processing outputRezTime = new processing();
                outputRezTime.sortTime(selectedName, fileName);
            }
        }  //по названию

    }

//---------------------------------------------------------------------------------------------------------------------------------------------------------------
    class processing
    {
        protected string procName="";
        protected string procType="";
        protected string procsTime="";
        protected string proccTime="";
        protected string procSize="";

        protected XmlDocument xmlDoc = new XmlDocument();

        public string checkFile(System.Windows.Forms.Label lab1, string fileName)
        {
            string filePath = "path.txt";
            if (System.IO.File.Exists(filePath))
            {
                fileName = System.IO.File.ReadAllText(filePath);
                if (fileName != "") { lab1.Text = "Выбран файл:  " + fileName; return fileName; }
                else return "";
            }
            else return "";
        }
        public void writePath(string fileName)
        {
            File.WriteAllText("path.txt", fileName);
        }
        public string loadPath(System.Windows.Forms.Label lab1, string fileName)
        {
            OpenFileDialog OPF = new OpenFileDialog();
            OPF.Filter = "xml-файлы(*.xml)|*.xml";
            DialogResult dlg = OPF.ShowDialog();
            if (dlg == DialogResult.OK)
            {
                fileName = OPF.FileName;
                lab1.Text = "Выбран файл:  " + fileName;
                File.WriteAllText("path.txt", fileName);
                return fileName;
            }
            else return "";
        }
        public void initGrid(System.Windows.Forms.DataGridView Gr1)
        {
            var column1 = new DataGridViewColumn();
            column1.HeaderText = "Process name"; 
            column1.Width = 150; 
            column1.ReadOnly = true; 
            column1.Name = "name"; //текстовое имя колонки, его можно использовать вместо обращений по индексу
            column1.Frozen = true; //флаг, что данная колонка всегда отображается на своем месте
            column1.CellTemplate = new DataGridViewTextBoxCell(); //тип нашей колонки
            
            var column2 = new DataGridViewColumn();
            column2.HeaderText = "Type"; 
            column2.Width = 75;
            column2.Name = "type";
            column2.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            column2.CellTemplate = new DataGridViewTextBoxCell();
            
            var column3 = new DataGridViewColumn();
            column3.HeaderText = "Start time";
            column3.Width = 150;
            column3.Name = "startDate";
            column3.CellTemplate = new DataGridViewTextBoxCell();

            var column4 = new DataGridViewColumn();
            column4.HeaderText = "Current time";
            column4.Width = 150;
            column4.Name = "currentDate";
            column4.CellTemplate = new DataGridViewTextBoxCell(); 
            
            var column5 = new DataGridViewColumn();
            column5.HeaderText = "Size";
            column5.Width = 150;
            column5.Name = "size";
            column5.CellTemplate = new DataGridViewTextBoxCell();

            Gr1.Columns.Add(column1);
            Gr1.Columns.Add(column2);
            Gr1.Columns.Add(column3);
            Gr1.Columns.Add(column4);
            Gr1.Columns.Add(column5);
            Gr1.AllowUserToAddRows = false;   
        }  //создание таблицы
        public void initData(System.Windows.Forms.DataGridView dataGridView1, System.Windows.Forms.Label label1, System.Windows.Forms.ComboBox comboBox1, ref string fileName)
        {
            fileName = checkFile(label1, fileName);
            if (fileName == "") fileName = loadPath(label1, fileName);
            if (fileName != "")
            {
                for (int i = 0; i < dataGridView1.ColumnCount; i++) 
                    for (int j = 0; j < dataGridView1.RowCount; j++)
                        if (dataGridView1[i, j].Value != null)
                        {
                            dataGridView1.Rows.Clear();
                        }
                comboBox1.Items.Clear();
                outputFromXML(dataGridView1, fileName);
                sortChoice(comboBox1, fileName);
            }
            label1.Text = "Выбран файл:  " + fileName;

        }

        public void outputScreenCurrentTraffic(System.Windows.Forms.DataGridView Gr, System.Windows.Forms.Label label1, string fileName)
        {
            PerformanceCounterCategory pCC = new PerformanceCounterCategory("Network Interface");
            string instance = pCC.GetInstanceNames()[0];
            PerformanceCounter pCSent = new PerformanceCounter("Network Interface", "Bytes Sent/sec", instance);
            PerformanceCounter pCReceived = new PerformanceCounter("Network Interface", "Bytes Received/sec", instance);

            System.Diagnostics.Process[] proc;
            proc = System.Diagnostics.Process.GetProcesses();
            foreach (System.Diagnostics.Process i in proc)
            {
                double sizeRec = pCReceived.NextValue();    
                double sizeSent = pCSent.NextValue();    
                if (sizeRec != 0)
                {
                    procsTime = Convert.ToString(DateTime.Now + "." + DateTime.Now.Millisecond);
                    procType = "0";
                    procName = i.ProcessName;
                    procSize = Convert.ToString(sizeRec+" B");
                    proccTime = Convert.ToString(DateTime.Now + "." + DateTime.Now.Millisecond);
                    Gr.Rows.Add(procName, procType, procsTime, proccTime, procSize);
                    outputInXml(fileName);
                }
                else if (sizeSent != 0)
                {
                    procType = "1";
                    procName = i.ProcessName;
                    procsTime = Convert.ToString(DateTime.Now + "." + DateTime.Now.Millisecond);
                    procSize = Convert.ToString(sizeSent+" B");
                    proccTime = Convert.ToString(DateTime.Now + "." + DateTime.Now.Millisecond);
                    Gr.Rows.Add(procName, procType, procsTime, proccTime, procSize); 
                    outputInXml(fileName);
                }
            }
        }
        public void outputInXml(string fileName)
        {
            xmlDoc.Load(fileName);
            XmlElement xmlRoot = xmlDoc.DocumentElement; //корневой элемент
            XmlElement elem_portion = xmlDoc.CreateElement("portion");
            XmlElement elem_processName = xmlDoc.CreateElement("processName");
            XmlElement elem_type = xmlDoc.CreateElement("type");
            XmlElement elem_sTime = xmlDoc.CreateElement("sTime");
            XmlElement elem_cTime = xmlDoc.CreateElement("cTime");
            XmlElement elem_size = xmlDoc.CreateElement("size");
            XmlText text_processName = xmlDoc.CreateTextNode(procName);
            XmlText text_type = xmlDoc.CreateTextNode(procType);
            XmlText text_sTime = xmlDoc.CreateTextNode(procsTime);
            XmlText text_cTime = xmlDoc.CreateTextNode(proccTime);
            XmlText text_size = xmlDoc.CreateTextNode(procSize);
            xmlRoot.AppendChild(elem_portion);
            elem_portion.AppendChild(elem_processName);
            elem_portion.AppendChild(elem_type);
            elem_portion.AppendChild(elem_sTime);
            elem_portion.AppendChild(elem_cTime);
            elem_portion.AppendChild(elem_size);
            elem_processName.AppendChild(text_processName);
            elem_type.AppendChild(text_type);
            elem_sTime.AppendChild(text_sTime);
            elem_cTime.AppendChild(text_cTime);
            elem_size.AppendChild(text_size);
            elem_processName.AppendChild(text_processName);
            xmlDoc.Save(fileName);
        }  //вывод в XML


        protected DateTime sTimeMax, sTimeMin;  // Период наблюдения
        public struct printList
        {
            public string pName;
            public int pType;
            public string psTime;
            public string pcTime;
            public Int64 pSize;
            public int mnt;
            public printList(string pName2, int pType2, string psTime2, string pcTime2, Int64 pSize2, int mnt2)
            {
                pName = pName2;
                pType = pType2;
                psTime = psTime2;
                pcTime = pcTime2;
                pSize = pSize2;
                mnt = mnt2;
            }
        }
        // Список порций для отчета Итоги. Формируется в LoadXML()
        protected List<printList> printListXML = new List<printList>();

        public bool LoadXML(string fileName)
        {
            DataSet portionsDS = new DataSet();
            try
            {
                printListXML.Clear();
                printList printCreat;
                string processName; // Имя процеса
                int pTp; // 0 - входящие данные; 1 - исходящие
                DateTime sTime, cTime; // Текущее время
                string sTime2, cTime2;
                Int64 size; // Размер порции в битах
                string size2; // Символьное представление размера порции
                portionsDS.ReadXml(fileName); // Формируем DataSet portionsDS
                DataTableCollection tbls = portionsDS.Tables; // Коллекция таблиц (у нас одна таблица traffic), содержит отдельные элементы коллекции
                System.Data.DataTable tbl = tbls[0]; // создается таблица с трафиком
                DataRowCollection portions = tbl.Rows; // Коллекция строк таблицы
                // Период наблюдения
                sTimeMax = DateTime.MinValue;
                sTimeMin = DateTime.MaxValue;
                foreach (DataRow portion in portions)
                {
                    processName = Convert.ToString(portion[0]);
                    pTp = Convert.ToInt32(portion["type"]);
                    sTime = Convert.ToDateTime(portion["sTime"]);
                    // Корректируем период наблюдения
                    if (sTime > sTimeMax) sTimeMax = sTime;
                    if (sTime < sTimeMin) sTimeMin = sTime;
                    cTime = Convert.ToDateTime(portion["cTime"]);
                    size2 = Convert.ToString(portion["size"]);
                    size2 = size2.Replace("B", "");
                    size2 = size2.Replace(",", "").Trim();
                    size = Convert.ToInt64(size2);
                    sTime2 = Convert.ToString(sTime);
                    cTime2 = Convert.ToString(cTime);
                    printCreat.pName = processName;
                    printCreat.pType = pTp;
                    printCreat.psTime = sTime2;
                    printCreat.pcTime = cTime2;
                    printCreat.pSize = size;
                    printCreat.mnt = 1;
                    printListXML.Add(printCreat);
                }
            }
            catch (System.IO.IOException e)
            {
                MessageBox.Show("Плохой файл " + fileName + ". Сообщение " + e.Message);
                return false;
            }
            finally
            {
                portionsDS.Dispose();
            }
            return true;
        }

        public void outputFromXML(System.Windows.Forms.DataGridView Gr, string fileName)
        {
            xmlDoc.Load(fileName);
            XmlElement xmlRoot = xmlDoc.DocumentElement;
            foreach (XmlNode xmlNode in xmlRoot)
            {
                foreach (XmlNode childnode in xmlNode)
                {
                        if (childnode.Name == "processName")
                            procName = childnode.InnerText;
                        if (childnode.Name == "type")
                            procType = childnode.InnerText;
                        if (childnode.Name == "sTime")
                            procsTime = childnode.InnerText;
                        if (childnode.Name == "cTime")
                            proccTime = childnode.InnerText;
                        if (childnode.Name == "size")
                            procSize = childnode.InnerText;
                }
                if (procSize != "") Gr.Rows.Add(procName, procType, procsTime, proccTime, procSize);
            }
        }

        protected Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
        public void sortName(string fileName)
        {
             if (!LoadXML(fileName)) return;
            // Имена всех процессы без повторений в алфавитном порядке
            var printListName = printListXML.GroupBy(p => p.pName).Select(
                grp => new printList { pName = grp.Key }).OrderBy(p => p.pName);
            // Итоги по входящим данным
            var printListRec = printListXML.Where(p => p.pType == 0).GroupBy(p => p.pName).Select(
              grp => new printList { pName = grp.Key, pSize = grp.Sum(p => p.pSize), mnt = grp.Sum(p => p.mnt) }).OrderBy(p => p.pName);
            // Итоги по исходящим данным
            var printListSent = printListXML.Where(p => p.pType == 1).GroupBy(p => p.pName).Select(
              grp => new printList { pName = grp.Key, pSize = grp.Sum(p => p.pSize), mnt = grp.Sum(p => p.mnt) }).OrderBy(p => p.pName);
            // Имена всех процессов
            string[] arrPrcAll = printListName.Select(p => p.pName).ToArray();
            // Входящие порции
            // Имена процессов
            string[] arrPrcIn = printListRec.Select(p => p.pName).ToArray();
            // Суммарные порции входящих данных соответственно процессов "Proc1", "Proc2", "Proc3", "Proc4
            Int64[] prtnInTotal = printListRec.Select(p => p.pSize).ToArray();
            // Число порций входящих данных
            int[] mntInTotal = printListRec.Select(p => p.mnt).ToArray();
            // Исходящие порции
            // Имена процессов
            string[] arrPrcOut = printListSent.Select(p => p.pName).ToArray();
            // Суммарные порции исходящих данных
            Int64[] prtnOutTotal = printListSent.Select(p => p.pSize).ToArray();
            // Число порций исходящих данных соответственно процессов "Proc2", "Proc4", "Proc5"
            int[] mntOutTotal = printListSent.Select(p => p.mnt).ToArray();
            if (xlApp == null)
            {
                MessageBox.Show("Excel не найден");
                return;
            }
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            // Заголовки
            xlWorkSheet.Cells[1, 2] = "Интернет-трафик за период с " + sTimeMin + " по " + sTimeMax;
            xlWorkSheet.Cells[2, 2] = "№";
            xlWorkSheet.Cells[2, 3] = "Процесс";
            xlWorkSheet.Cells[2, 4] = "Производительность приема";
            xlWorkSheet.Cells[2, 6] = "Производительность передачи";
            xlWorkSheet.Cells[3, 4] = "Байт";
            xlWorkSheet.Cells[3, 5] = "Число порций";
            xlWorkSheet.Cells[3, 6] = "Байт";
            xlWorkSheet.Cells[3, 7] = "Число порций";
            int lstClmn = 8; // Следующий после крайнего справа столбца
            // Выравнивание
            for (int i = 2; i < 4; i++)
                for (int j = 2; j < lstClmn; j++)
                    xlWorkSheet.Cells[i, j].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            // Ширина столбцов
            Excel.Range range = xlWorkSheet.Range["B1:B1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 7;
            range = xlWorkSheet.Range["C1:C1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 30;
            range = xlWorkSheet.Range["D1:G1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 15;
            // Границы ячеек с заголовками
            xlWorkSheet.Cells[2, 2].Borders(1).ColorIndex = 1; // Граница слева
            xlWorkSheet.Cells[2, 3].Borders(1).ColorIndex = 1; // Граница слева
            xlWorkSheet.Cells[2, 3].Borders(2).ColorIndex = 1; // Граница справа
            xlWorkSheet.Cells[2, 5].Borders(2).ColorIndex = 1; // Граница справа
            xlWorkSheet.Cells[2, 7].Borders(2).ColorIndex = 1; // Граница справа
            for (int i = 2; i < lstClmn; i++) xlWorkSheet.Cells[2, i].Borders(3).ColorIndex = 1; // Граница сверху
            for (int i = 4; i < lstClmn; i++) xlWorkSheet.Cells[3, i].Borders(3).ColorIndex = 1; // Граница сверху
            for (int i = 2; i < lstClmn; i++) xlWorkSheet.Cells[3, i].Borders(4).ColorIndex = 1; // Граница снизу
            for (int i = 2; i < lstClmn; i++)
                for (int k = 1; k < 3; k++)
                    xlWorkSheet.Cells[3, i].Borders(k).ColorIndex = 1; // Границы слева и справа
            // Объединение ячеек
            range = xlWorkSheet.get_Range("D2:E2");
            range.Merge(Type.Missing);
            range = xlWorkSheet.get_Range("F2:G2");
            range.Merge(Type.Missing);
            // Всего имен процессов
            int prcMnt = arrPrcAll.Length;
            // Границы ячеек с данными
            for (int i = 4; i < 4 + prcMnt; i++)
                for (int j = 2; j < lstClmn; j++)
                    for (int k = 1; k < 5; k++)
                        xlWorkSheet.Cells[i, j].Borders(k).ColorIndex = 1; // Границы слева, справа, сверху и снизу
            // Выводим номера строк и имена процессов
            int n = 0;
            for (int i = 0; i < prcMnt; i++)
            {
                n++;
                xlWorkSheet.Cells[i + 4, 2] = n;
                xlWorkSheet.Cells[i + 4, 3] = arrPrcAll[i];
            }
            // Число входящих и исходящих данных в итоговых массивах
            int inMnt = arrPrcIn.Length;
            int outMnt = arrPrcOut.Length;
            // Получено
            for (int i = 0; i < prcMnt; i++)
            {
                string prcNm = arrPrcAll[i];
                for (int j = 0; j < inMnt; j++)
                {
                    string prcNm2 = arrPrcIn[j];
                    if (prcNm2 == prcNm)
                    {
                        xlWorkSheet.Cells[i + 4, 4] = prtnInTotal[j] / mntInTotal[j];
                        xlWorkSheet.Cells[i + 4, 5] = mntInTotal[j];
                    }
                }
            }
            // Передано
            for (int i = 0; i < prcMnt; i++)
            {
                string prcNm = arrPrcAll[i];
                for (int j = 0; j < outMnt; j++)
                {
                    string prcNm2 = arrPrcOut[j];
                    if (prcNm2 == prcNm)
                    {
                        xlWorkSheet.Cells[i + 4, 6] = prtnOutTotal[j] / mntOutTotal[j];
                        xlWorkSheet.Cells[i + 4, 7] = mntOutTotal[j];
                    }
                }
            }
            // Показываем отчет, выведенный в Excel
            xlApp.Visible = true;

        }

        public void sortChoice(System.Windows.Forms.ComboBox cBox, string fileName)
        {
            if (!LoadXML(fileName)) return;
            var printListName = printListXML.GroupBy(p => p.pName).Select(
                grp => new printList { pName = grp.Key }).OrderBy(p => p.pName);
            string[] arrPrcAll = printListName.Select(p => p.pName).ToArray();
            cBox.Items.AddRange(arrPrcAll);
        }
        public void sortTime(string selectedName, string fileName)
        {
            if (!LoadXML(fileName)) return;
            
            var printListcTimeRec = printListXML.Where(p => p.pName == selectedName && p.pType == 0).Select(p => p.pcTime);
            var printListcTimeSent = printListXML.Where(p => p.pName == selectedName && p.pType == 1).Select(p => p.pcTime);
            var printListsTimeRec = printListXML.Where(p => p.pName == selectedName && p.pType == 0).Select(p => p.psTime);
            var printListsTimeSent = printListXML.Where(p => p.pName == selectedName && p.pType == 1).Select(p => p.psTime);
            var printListsSizeRec = printListXML.Where(p => p.pName == selectedName && p.pType == 0).Select(p => p.pSize);
            var printListSizeSent = printListXML.Where(p => p.pName == selectedName && p.pType == 1).Select(p => p.pSize);

            bool flRec=true, flSent=true;
            if (printListsSizeRec.Count()<1) flRec = false;
            if (printListSizeSent.Count()<1) flSent = false;

            
                string[] arrsTimeRec = printListsTimeRec.ToArray();
                string[] arrcTimeRec = printListcTimeRec.ToArray();
                Int64[] arrSizeRec = printListsSizeRec.ToArray();
            
            
                string[] arrcTimeSent = printListcTimeSent.ToArray();
                string[] arrsTimeSent = printListsTimeSent.ToArray();
                Int64[] arrSizeSent = printListSizeSent.ToArray();
            

            sTimeMax = DateTime.MinValue;
            sTimeMin = DateTime.MaxValue;
            if (flRec == true && flSent == true)  
            {
                DateTime t_min0 = Convert.ToDateTime(printListsTimeRec.Min());
                DateTime t_max0 = Convert.ToDateTime(printListsTimeRec.Max());
                DateTime t_min1 = Convert.ToDateTime(printListsTimeSent.Min());
                DateTime t_max1 = Convert.ToDateTime(printListsTimeSent.Max()); 
                if (t_min0 < t_min1) sTimeMin = t_min0;
                else sTimeMin = t_min1;
                if (t_max0 > t_max1) sTimeMax = t_max0;
                else sTimeMax = t_max1;

            }
            else
                if (flRec==true)
                {
                    sTimeMin = Convert.ToDateTime(printListsTimeRec.Min());
                    sTimeMax = Convert.ToDateTime(printListcTimeRec.Max());
                }
                else
                {
                    sTimeMin = Convert.ToDateTime(printListsTimeSent.Min());
                    sTimeMax = Convert.ToDateTime(printListcTimeSent.Max());
                }

            if (xlApp == null)
            {
                MessageBox.Show("Excel не найден");
                return;
            }
            int inMnt = arrsTimeRec.Length;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            xlWorkBook = xlApp.Workbooks.Add();
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            if (flRec == true) creatExcelRec(selectedName, arrcTimeRec, arrsTimeRec, arrSizeRec, xlWorkSheet);
            if (flSent == true) creatExcelSent(arrsTimeRec.Length, selectedName, arrcTimeSent, arrsTimeSent, arrSizeSent, xlWorkSheet);            
            // Показываем отчет, выведенный в Excel
            xlApp.Visible = true;
        }
        public void creatExcelRec(string selectedName, string[] arrcTimeRec, string[] arrsTimeRec, Int64[] arrSizeRec, Excel.Worksheet xlWorkSheet)
        {

            int inMnt = arrsTimeRec.Length;
            // Заголовки
            xlWorkSheet.Cells[1, 2] = "История процесса " + selectedName + " за период с " + sTimeMin + " по " + sTimeMax;
            xlWorkSheet.Cells[2, 2] = "Текущее время";
            xlWorkSheet.Cells[2, 3] = "Время порции";
            xlWorkSheet.Cells[2, 4] = "Производительность, байт/с";
            xlWorkSheet.Cells[3, 2] = "Входящие данные";
            // Выравнивание
            for (int i = 2; i < 5; i++) xlWorkSheet.Cells[2, i].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
            // Границы ячеек
            for (int i = 2; i < 5 + inMnt; i++)
                for (int j = 2; j < 5; j++)
                    for (int k = 1; k < 5; k++)
                    {
                        if (i == 3 || i == 4 + inMnt) continue;
                        xlWorkSheet.Cells[i, j].Borders(k).ColorIndex = 1; // Границы слева, справа, сверху и снизу
                    }
            // Границы ячеек строках 3 и 4 + inMnt с текстом Входящие данные и Исходящие данные
            xlWorkSheet.Cells[3, 2].Borders(1).ColorIndex = 1; // Граница слева
            xlWorkSheet.Cells[3, 4].Borders(2).ColorIndex = 1; // Граница справа
            // Ширина столбцов
            Excel.Range range = xlWorkSheet.Range["B1:D1", System.Type.Missing];
            range.EntireColumn.ColumnWidth = 25;
            // Вывод входящих данных
            for (int i = 0; i < inMnt; i++)
            {
                xlWorkSheet.Cells[i + 4, 2] = arrcTimeRec[i];
                xlWorkSheet.Cells[i + 4, 3] = arrsTimeRec[i];
                xlWorkSheet.Cells[i + 4, 4] = arrSizeRec[i];
            }
            
        }
        public void creatExcelSent(int count, string selectedName, string[] arrcTimeSent, string[] arrsTimeSent, Int64[] arrSizeSent, Excel.Worksheet xlWorkSheet)
        {
            // Число входящих и исходящих данных

            int outMnt = arrsTimeSent.Length;
            if (count == 0)
            {
                // Заголовки
                xlWorkSheet.Cells[1, 2] = "История процесса " + selectedName + " за период с " + sTimeMin + " по " + sTimeMax;
                xlWorkSheet.Cells[2, 2] = "Текущее время";
                xlWorkSheet.Cells[2, 3] = "Время порции";
                xlWorkSheet.Cells[2, 4] = "Размер порции, байт";
                xlWorkSheet.Cells[3, 2] = "Исходящие данные";
                // Выравнивание
                for (int i = 2; i < 5; i++) xlWorkSheet.Cells[2, i].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                // Границы ячеек
                for (int i = 2; i < 5 + outMnt; i++)
                    for (int j = 2; j < 5; j++)
                        for (int k = 1; k < 5; k++)
                        {
                            if (i == 3 || i == 4 + outMnt) continue;
                            xlWorkSheet.Cells[i, j].Borders(k).ColorIndex = 1; // Границы слева, справа, сверху и снизу
                        }
                // Границы ячеек строка 3 
                xlWorkSheet.Cells[3, 2].Borders(1).ColorIndex = 1; // Граница слева
                xlWorkSheet.Cells[3, 4].Borders(2).ColorIndex = 1; // Граница справа
                // Ширина столбцов
                Excel.Range range = xlWorkSheet.Range["B1:D1", System.Type.Missing];
                range.EntireColumn.ColumnWidth = 25;
                // Вывод входящих данных
                for (int i = 0; i < outMnt; i++)
                {
                    xlWorkSheet.Cells[i + 4, 2] = arrcTimeSent[i];
                    xlWorkSheet.Cells[i + 4, 3] = arrsTimeSent[i];
                    xlWorkSheet.Cells[i + 4, 4] = arrSizeSent[i];
                }
            }
            else 
            {
                xlWorkSheet.Cells[count + 4, 2] = "Исходящие данные";
                // Выравнивание
                for (int i = 2; i < 5; i++) xlWorkSheet.Cells[2, i].HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter;
                // Границы ячеек
                for (int i = 2; i < 5 + count + outMnt; i++)
                    for (int j = 2; j < 5; j++)
                        for (int k = 1; k < 5; k++)
                        {
                            if (i == 3 || i == 4 + count) continue;
                            xlWorkSheet.Cells[i, j].Borders(k).ColorIndex = 1; // Границы слева, справа, сверху и снизу
                        }
                // Границы ячеек 
                xlWorkSheet.Cells[3, 2].Borders(1).ColorIndex = 1; // Граница слева
                xlWorkSheet.Cells[3, 4].Borders(2).ColorIndex = 1; // Граница справа
                xlWorkSheet.Cells[4 + count, 2].Borders(1).ColorIndex = 1; // Граница слева
                xlWorkSheet.Cells[4 + count, 4].Borders(2).ColorIndex = 1; // Граница справа
                // Ширина столбцов
                Excel.Range range = xlWorkSheet.Range["B1:D1", System.Type.Missing];
                range.EntireColumn.ColumnWidth = 25;
                // Вывод исходящих данных
                for (int i = 0; i < outMnt; i++)
                {
                    xlWorkSheet.Cells[i + count + 5, 2] = arrcTimeSent[i];
                    xlWorkSheet.Cells[i + count + 5, 3] = arrsTimeSent[i];
                    xlWorkSheet.Cells[i + count + 5, 4] = arrSizeSent[i];
                }
            }
        }
    }
}

