using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BeosztasGenerator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public class AssignClass
        {
            public string Name;
            public string Year;
            public bool Monday;
            public bool Tuesday;
            public bool Wednesday;
            public bool Thursday;
            public bool Friday;
            public bool Saturday;
            public bool Sunday;
            public bool NewYear;
            public bool GoodFriday;
            public bool EasterMonday;
            public bool PentecostMonday;
            public bool Christmas1st;
            public bool Christmas2nd;
            public bool NewYearsEve;

            public List<List<string>> GroupList = new List<List<string>>();

            public string FirstGroup;

            public AssignClass()
            {
                ResetAssignment();
            }

            public void ResetAssignment()
            {
                Name = "";
                Year = "";
                Monday = false;
                Tuesday = false;
                Wednesday = false;
                Thursday = false;
                Friday = false;
                Saturday = false;
                Sunday = false;
                NewYear = false;
                GoodFriday = false;
                EasterMonday = false;
                PentecostMonday = false;
                Christmas1st = false;
                Christmas2nd = false;
                NewYearsEve = false;

                GroupList.Clear();

                FirstGroup = "";
            }
        }

        public AssignClass MyAssignment = new AssignClass();
        public List<DateTime> DateList = new List<DateTime>();

        public void RefreshGUI()
        {
            textBox_AssignmentName.Text = MyAssignment.Name;
            comboYear.SelectedIndex = comboYear.FindStringExact(MyAssignment.Year);
            checkBox_Monday.Checked = MyAssignment.Monday;
            checkBox_Tuesday.Checked = MyAssignment.Tuesday;
            checkBox_Wednesday.Checked = MyAssignment.Wednesday;
            checkBox_Thursday.Checked = MyAssignment.Thursday;
            checkBox_Friday.Checked = MyAssignment.Friday;
            checkBox_Saturday.Checked = MyAssignment.Saturday;
            checkBox_Sunday.Checked = MyAssignment.Sunday;
            checkBox_NewYear.Checked = MyAssignment.NewYear;
            checkBox_GoodFriday.Checked = MyAssignment.GoodFriday;
            checkBox_EasterMonday.Checked = MyAssignment.EasterMonday;
            checkBox_PentecostMonday.Checked = MyAssignment.PentecostMonday;
            checkBox_Christmas1st.Checked = MyAssignment.Christmas1st;
            checkBox_Christmas2nd.Checked = MyAssignment.Christmas2nd;
            checkBox_NewYearsEve.Checked = MyAssignment.NewYearsEve;

            InitDataGridView();

            int GroupCounter = 0;
            foreach (List<string> Group in MyAssignment.GroupList)
            {
                int PersonCounter = 0;
                foreach (string Person in Group)
                {
                    if (!String.IsNullOrEmpty(Person))
                    {
                        dataGridView1.Rows[GroupCounter].Cells[PersonCounter].Value = Person;
                    }
                    PersonCounter++;
                }
                GroupCounter++;
            }

            if (!StartGroupComboBox.Items.Contains(MyAssignment.FirstGroup))
            {
                StartGroupComboBox.Items.Add(MyAssignment.FirstGroup);
            }

            StartGroupComboBox.Text = MyAssignment.FirstGroup;
        }

        public void RefreshAssignment()
        {
            MyAssignment.Name = textBox_AssignmentName.Text;

            if (!String.IsNullOrEmpty(comboYear.Text))
            {
                MyAssignment.Year = comboYear.Text;
            }
            else
            {
                MyAssignment.Year = "";
            }

            MyAssignment.Monday = checkBox_Monday.Checked;
            MyAssignment.Tuesday = checkBox_Tuesday.Checked;
            MyAssignment.Wednesday = checkBox_Wednesday.Checked;
            MyAssignment.Thursday = checkBox_Thursday.Checked;
            MyAssignment.Friday = checkBox_Friday.Checked;
            MyAssignment.Saturday = checkBox_Saturday.Checked;
            MyAssignment.Sunday = checkBox_Sunday.Checked;
            MyAssignment.NewYear = checkBox_NewYear.Checked;
            MyAssignment.GoodFriday = checkBox_GoodFriday.Checked;
            MyAssignment.EasterMonday = checkBox_EasterMonday.Checked;
            MyAssignment.PentecostMonday = checkBox_PentecostMonday.Checked;
            MyAssignment.Christmas1st = checkBox_Christmas1st.Checked;
            MyAssignment.Christmas2nd = checkBox_Christmas2nd.Checked;
            MyAssignment.NewYearsEve = checkBox_NewYearsEve.Checked;

            MyAssignment.GroupList.Clear();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                List<string> rowList = new List<string>();

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null)
                    {
                        string cellValue = cell.Value.ToString();

                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            rowList.Add(cellValue);
                        }
                    }
                }

                MyAssignment.GroupList.Add(rowList);
            }

            MyAssignment.FirstGroup = StartGroupComboBox.Text;
        }

        public List<DateTime> GetDatesForDay(DayOfWeek Day)
        {
            List<DateTime> LocalDateList = new List<DateTime>();

            var start = new DateTime(Int32.Parse(MyAssignment.Year), 1, 1);
            var end = start.AddYears(1);

            while (start < end)
            {
                if (start.DayOfWeek == Day)
                {
                    LocalDateList.Add(start);
                    start = start.AddDays(7);
                }
                else
                    start = start.AddDays(1);
            }

            return LocalDateList;
        }

        public static DateTime GetEasterSunday(int year)
        {
            int day = 0;
            int month = 0;

            int g = year % 19;
            int c = year / 100;
            int h = (c - (int)(c / 4) - (int)((8 * c + 13) / 25) + 19 * g + 15) % 30;
            int i = h - (int)(h / 28) * (1 - (int)(h / 28) * (int)(29 / (h + 1)) * (int)((21 - g) / 11));

            day = i - ((year + (int)(year / 4) + i + 2 - c + (int)(c / 4)) % 7) + 28;
            month = 3;

            if (day > 31)
            {
                month++;
                day -= 31;
            }

            return new DateTime(year, month, day);
        }

        public void CollectDates()
        {
            DateList.Clear();

            if (MyAssignment.Monday)
            {
                DateList.AddRange(GetDatesForDay(DayOfWeek.Monday));
            }
            if (MyAssignment.Tuesday)
            {
                DateList.AddRange(GetDatesForDay(DayOfWeek.Tuesday));
            }
            if (MyAssignment.Wednesday)
            {
                DateList.AddRange(GetDatesForDay(DayOfWeek.Wednesday));
            }
            if (MyAssignment.Thursday)
            {
                DateList.AddRange(GetDatesForDay(DayOfWeek.Thursday));
            }
            if (MyAssignment.Friday)
            {
                DateList.AddRange(GetDatesForDay(DayOfWeek.Friday));
            }
            if (MyAssignment.Saturday)
            {
                DateList.AddRange(GetDatesForDay(DayOfWeek.Saturday));
            }
            if (MyAssignment.Sunday)
            {
                DateList.AddRange(GetDatesForDay(DayOfWeek.Sunday));
            }
            if (MyAssignment.NewYear)
            {
                DateList.Add(new DateTime(Int32.Parse(MyAssignment.Year), 1, 1));
            }
            if(MyAssignment.GoodFriday)
            {                
                DateList.Add(GetEasterSunday(Int32.Parse(MyAssignment.Year)).AddDays(-2));
            }
            if(MyAssignment.EasterMonday)
            {
                DateList.Add(GetEasterSunday(Int32.Parse(MyAssignment.Year)).AddDays(1));
            }
            if(MyAssignment.PentecostMonday)
            {
                DateList.Add(GetEasterSunday(Int32.Parse(MyAssignment.Year)).AddDays(50));
            }
            if (MyAssignment.Christmas1st)
            {
                DateList.Add(new DateTime(Int32.Parse(MyAssignment.Year), 12, 25));
            }
            if (MyAssignment.Christmas2nd)
            {
                DateList.Add(new DateTime(Int32.Parse(MyAssignment.Year), 12, 26));
            }
            if (MyAssignment.NewYearsEve)
            {
                DateList.Add(new DateTime(Int32.Parse(MyAssignment.Year), 12, 31));
            }

            /* Remove duplicates from DateList. */
            List<DateTime> LocalDateList = new List<DateTime>();
            foreach (var date in DateList)
            {
                if (!LocalDateList.Contains(date))
                {
                    LocalDateList.Add(date);
                }
            }
            DateList = LocalDateList;

            /* Sort the DateList. */            
            DateList.Sort((a, b) => a.CompareTo(b));
        }

        public void InitDataGridView()
        {
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            dataGridView1.Rows[0].Cells[0].Value = "";
            dataGridView1.Rows.AddCopies(0, 15);
            dataGridView1.AllowUserToAddRows = false;
        }

        public bool SetupIsValid()
        {
            bool Result = true;

            if (string.IsNullOrWhiteSpace(MyAssignment.Name))
            {
                Result &= false;
                MessageBox.Show("A beosztás generálása nem sikerült!\n\nNem adtál nevet a beosztásodnak!", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return Result;
            }

            if (MyAssignment.Year == string.Empty)
            {
                Result &= false;
                MessageBox.Show("A beosztás generálása nem sikerült!\n\nNem adtál meg évet!", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return Result;
            }

            bool YearWrong = false;
            foreach (char c in MyAssignment.Year)
            {
                if (!char.IsDigit(c))
                {
                    YearWrong = true;
                }
            }

            if (YearWrong)
            {
                Result &= false;
                MessageBox.Show("A beosztás generálása nem sikerült!\n\nHelytelenül adtad meg az évet!", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return Result;
            }

            if ((!MyAssignment.Monday) &&
                (!MyAssignment.Tuesday) &&
                (!MyAssignment.Wednesday) &&
                (!MyAssignment.Thursday) &&
                (!MyAssignment.Friday) &&
                (!MyAssignment.Saturday) &&
                (!MyAssignment.Sunday) &&
                (!MyAssignment.NewYear) &&
                (!MyAssignment.GoodFriday) &&
                (!MyAssignment.EasterMonday) &&
                (!MyAssignment.PentecostMonday) &&
                (!MyAssignment.Christmas1st) &&
                (!MyAssignment.Christmas2nd) &&
                (!MyAssignment.NewYearsEve))
            {
                Result &= false;
                MessageBox.Show("A beosztás generálása nem sikerült!\n\nNem választottál ki semmilyen napot!", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return Result;
            }

            bool GroupListEmpty = true;
            foreach (var Group in MyAssignment.GroupList)
            {
                foreach (var Person in Group)
                {
                    if (!string.IsNullOrWhiteSpace(Person))
                    {
                        GroupListEmpty = false;
                    }
                }
            }

            if (GroupListEmpty)
            {
                Result &= false;
                MessageBox.Show("A beosztás generálása nem sikerült!\n\nNem adtál meg személyeket.", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return Result;           
        }

        private void button_NewAssignment_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Ezzel minden mezőt törölsz. Biztos, hogy akarod?", "Figyelmeztetés", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                MyAssignment.ResetAssignment();
                RefreshGUI();
            }          
        }

        private void button_LoadAssignment_Click(object sender, EventArgs e)
        {
            OpenFileDialog Dialog = new OpenFileDialog();

            Dialog.Filter = "Excel fájl (*.xlsx)|*.xlsx";

            if (Dialog.ShowDialog() == DialogResult.OK)
            {
                bool SettingsFound = false;

                try
                {
                    byte[] bin = File.ReadAllBytes(Dialog.FileName);

                    using (MemoryStream stream = new MemoryStream(bin))
                    using (ExcelPackage excelPackage = new ExcelPackage(stream))
                    {
                        foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                        {
                            if (worksheet.Name == "Beállítások (ne módosítsd!)")
                            {
                                SettingsFound = true;                             
                                MyAssignment.Name = worksheet.Cells["B2"].Value.ToString();
                                MyAssignment.Year = worksheet.Cells["B3"].Value.ToString();
                                MyAssignment.Monday = bool.Parse(worksheet.Cells["B4"].Value.ToString());
                                MyAssignment.Tuesday = bool.Parse(worksheet.Cells["B5"].Value.ToString());
                                MyAssignment.Wednesday = bool.Parse(worksheet.Cells["B6"].Value.ToString());
                                MyAssignment.Thursday = bool.Parse(worksheet.Cells["B7"].Value.ToString());
                                MyAssignment.Friday = bool.Parse(worksheet.Cells["B8"].Value.ToString());
                                MyAssignment.Saturday = bool.Parse(worksheet.Cells["B9"].Value.ToString());
                                MyAssignment.Sunday = bool.Parse(worksheet.Cells["B10"].Value.ToString());
                                MyAssignment.NewYear = bool.Parse(worksheet.Cells["B11"].Value.ToString());
                                MyAssignment.GoodFriday = bool.Parse(worksheet.Cells["B12"].Value.ToString());
                                MyAssignment.EasterMonday = bool.Parse(worksheet.Cells["B13"].Value.ToString());
                                MyAssignment.PentecostMonday = bool.Parse(worksheet.Cells["B14"].Value.ToString());
                                MyAssignment.Christmas1st = bool.Parse(worksheet.Cells["B15"].Value.ToString());
                                MyAssignment.Christmas2nd = bool.Parse(worksheet.Cells["B16"].Value.ToString());
                                MyAssignment.NewYearsEve = bool.Parse(worksheet.Cells["B17"].Value.ToString());
                                MyAssignment.FirstGroup = worksheet.Cells["B18"].Value.ToString();

                                MyAssignment.GroupList.Clear();
                                for (int i = 19; i < 34; i++)
                                {
                                    List<String> Group = new List<String>();

                                    for (int j = 2; j < 14; j++)
                                    {
                                        if (worksheet.Cells[i, j].Value != null)
                                        {
                                            Group.Add(worksheet.Cells[i, j].Value.ToString());                                           
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }

                                    MyAssignment.GroupList.Add(Group);
                                }
                            }
                        }
                    }
                }
                catch
                {
                    SettingsFound = false;
                }

                if (SettingsFound == false)
                {
                    MessageBox.Show("A beosztás betöltése nem sikerült! Lehetséges hibaokok:\n\n   - Hibás/sérült a fájl.\n   - A fájl nem a Beosztás Generátorral lett készítve.\n   - A fájl nyitva van Excelben.", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    RefreshGUI();
                }
            }
        }        

        private void Form1_Load(object sender, EventArgs e)
        {
            InitDataGridView();
        }

        private void button_GenerateAssignment_Click(object sender, EventArgs e)
        {
            SaveFileDialog Dialog = new SaveFileDialog();

            Dialog.OverwritePrompt = true;
            Dialog.Filter = "Excel fájl (*.xlsx)|*.xlsx";
            Dialog.FileName = MyAssignment.Name + "_" + MyAssignment.Year + ".xlsx";

            RefreshAssignment();            

            if ((SetupIsValid()) && (Dialog.ShowDialog() == DialogResult.OK))
            {
                CollectDates();

                using (var p = new ExcelPackage())
                {
                    var AssignWorkSheet = p.Workbook.Worksheets.Add(MyAssignment.Name + " " + MyAssignment.Year);  
                    AssignWorkSheet.Cells["A1"].Value = MyAssignment.Name + " " + MyAssignment.Year;
                    AssignWorkSheet.Cells["A1:M1"].Merge = true;
                    AssignWorkSheet.Cells["A1:M1"].Style.Font.Size = 16;

                    AssignWorkSheet.Cells["B2"].Value = "Jan";
                    AssignWorkSheet.Cells["C2"].Value = "Feb";
                    AssignWorkSheet.Cells["D2"].Value = "Már";
                    AssignWorkSheet.Cells["E2"].Value = "Ápr";
                    AssignWorkSheet.Cells["F2"].Value = "Máj";
                    AssignWorkSheet.Cells["G2"].Value = "Jún";
                    AssignWorkSheet.Cells["H2"].Value = "Júl";
                    AssignWorkSheet.Cells["I2"].Value = "Aug";
                    AssignWorkSheet.Cells["J2"].Value = "Szept";
                    AssignWorkSheet.Cells["K2"].Value = "Okt";
                    AssignWorkSheet.Cells["L2"].Value = "Nov";
                    AssignWorkSheet.Cells["M2"].Value = "Dec";
                    AssignWorkSheet.Cells["B2:M2"].Style.Font.Size = 12;

                    AssignWorkSheet.Cells["A1:M2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    AssignWorkSheet.Cells["A1:M2"].Style.Font.Bold = true;

                    int RowCounter = 3;

                    foreach (var Group in MyAssignment.GroupList)
                    {
                        string tempString = "";

                        foreach (var Person in Group)
                        {
                            if (!String.IsNullOrEmpty(Person))
                            {
                                tempString = tempString + Person + "\n";
                            }                            
                        }

                        tempString = tempString.TrimEnd('\r', '\n');
                        AssignWorkSheet.Cells[RowCounter, 1].Style.WrapText = true;
                        AssignWorkSheet.Cells[RowCounter, 1].Style.Font.Bold = true;
                        AssignWorkSheet.Cells[RowCounter, 1].Value = tempString;

                        if (Group.Count != 0)
                        {
                            RowCounter++;
                        }
                    }

                    int index = 0;
                    int GroupCount = 0;
                    foreach (var Group in MyAssignment.GroupList)
                    {
                        if (Group.Count != 0)
                        {
                            if (Group[0] == MyAssignment.FirstGroup)
                            {
                                index = GroupCount;
                            }

                            GroupCount++;
                        }
                    }
                    
                    foreach (var date in DateList)
                    {
                        int row = index + 3;
                        int column = date.Month + 1;

                        if (AssignWorkSheet.Cells[row, column].Value != null)
                        {
                            AssignWorkSheet.Cells[row, column].Value = AssignWorkSheet.Cells[row, column].Value + ", " + date.Day.ToString();
                        }
                        else
                        {
                            AssignWorkSheet.Cells[row, column].Value = date.Day.ToString();
                        }

                        AssignWorkSheet.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        AssignWorkSheet.Cells[row, column].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        if (index == (GroupCount - 1))
                        {
                            index = 0;
                        }
                        else
                        {
                            index++;
                        }
                    }

                    string LastRow = (GroupCount + 2).ToString();

                    AssignWorkSheet.Cells["A1:M" + LastRow].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    AssignWorkSheet.Cells["A1:M" + LastRow].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    AssignWorkSheet.Cells["A1:M" + LastRow].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    AssignWorkSheet.Cells["A1:M" + LastRow].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    AssignWorkSheet.Cells["A1:M1"].Style.Border.Top.Style = ExcelBorderStyle.Thick;
                    AssignWorkSheet.Cells["A" + LastRow + ":M" + LastRow].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    AssignWorkSheet.Cells["A1:A" + LastRow].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                    AssignWorkSheet.Cells["M1:M" + LastRow].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    AssignWorkSheet.Cells["A2:M2"].Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    AssignWorkSheet.Cells["A3:A" + LastRow].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                    AssignWorkSheet.Cells["B2:M2"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    AssignWorkSheet.Cells["B2:M2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    AssignWorkSheet.Cells["A3:A" + LastRow].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    AssignWorkSheet.Cells["A3:A" + LastRow].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    AssignWorkSheet.Column(1).Width = 21;
                    AssignWorkSheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    AssignWorkSheet.PrinterSettings.TopMargin = (decimal).2 / 2.54M;
                    AssignWorkSheet.PrinterSettings.LeftMargin = (decimal).2 / 2.54M;
                    AssignWorkSheet.PrinterSettings.RightMargin = (decimal).2 / 2.54M;
                    AssignWorkSheet.PrinterSettings.BottomMargin = (decimal).2 / 2.54M;

                    /* Save the settings to another worksheet. */
                    var SettingsWorksheet = p.Workbook.Worksheets.Add("Beállítások (ne módosítsd!)");

                    SettingsWorksheet.Cells["A1"].Value = "Beállítások (ezt a lapot ne módosítsd!)";
                    SettingsWorksheet.Cells["A1:K1"].Merge = true;
                    SettingsWorksheet.Cells["A1"].Style.Font.Size = 16;
                    SettingsWorksheet.Cells["A1"].Style.Font.Bold = true;

                    SettingsWorksheet.Cells["A2"].Value = "Name";
                    SettingsWorksheet.Cells["B2"].Value = MyAssignment.Name;
                    SettingsWorksheet.Cells["A3"].Value = "Year";
                    SettingsWorksheet.Cells["B3"].Value = MyAssignment.Year;
                    SettingsWorksheet.Cells["A4"].Value = "Monday";
                    SettingsWorksheet.Cells["B4"].Value = MyAssignment.Monday.ToString();
                    SettingsWorksheet.Cells["A5"].Value = "Tuesday";
                    SettingsWorksheet.Cells["B5"].Value = MyAssignment.Tuesday.ToString();
                    SettingsWorksheet.Cells["A6"].Value = "Wednesday";
                    SettingsWorksheet.Cells["B6"].Value = MyAssignment.Wednesday.ToString();
                    SettingsWorksheet.Cells["A7"].Value = "Thursday";
                    SettingsWorksheet.Cells["B7"].Value = MyAssignment.Thursday.ToString();
                    SettingsWorksheet.Cells["A8"].Value = "Friday";
                    SettingsWorksheet.Cells["B8"].Value = MyAssignment.Friday.ToString();
                    SettingsWorksheet.Cells["A9"].Value = "Saturday";
                    SettingsWorksheet.Cells["B9"].Value = MyAssignment.Saturday.ToString();
                    SettingsWorksheet.Cells["A10"].Value = "Sunday";
                    SettingsWorksheet.Cells["B10"].Value = MyAssignment.Sunday.ToString();
                    SettingsWorksheet.Cells["A11"].Value = "NewYear";
                    SettingsWorksheet.Cells["B11"].Value = MyAssignment.NewYear.ToString();
                    SettingsWorksheet.Cells["A12"].Value = "GoodFriday";
                    SettingsWorksheet.Cells["B12"].Value = MyAssignment.GoodFriday.ToString();
                    SettingsWorksheet.Cells["A13"].Value = "EasterMonday";
                    SettingsWorksheet.Cells["B13"].Value = MyAssignment.EasterMonday.ToString();
                    SettingsWorksheet.Cells["A14"].Value = "PentecostMonday";
                    SettingsWorksheet.Cells["B14"].Value = MyAssignment.PentecostMonday.ToString();
                    SettingsWorksheet.Cells["A15"].Value = "Christmas1st";
                    SettingsWorksheet.Cells["B15"].Value = MyAssignment.Christmas1st.ToString();
                    SettingsWorksheet.Cells["A16"].Value = "Christmas2nd";
                    SettingsWorksheet.Cells["B16"].Value = MyAssignment.Christmas2nd.ToString();
                    SettingsWorksheet.Cells["A17"].Value = "NewYearsEve";
                    SettingsWorksheet.Cells["B17"].Value = MyAssignment.NewYearsEve.ToString();
                    SettingsWorksheet.Cells["A18"].Value = "FirstGroup";
                    SettingsWorksheet.Cells["B18"].Value = MyAssignment.FirstGroup;

                    RowCounter = 19;
                    foreach (var Group in MyAssignment.GroupList)
                    {
                        SettingsWorksheet.Cells[RowCounter, 1].Value = "Group" + (RowCounter-18).ToString();

                        int ColumnCounter = 2;
                        foreach (var Person in Group)
                        {
                            SettingsWorksheet.Cells[RowCounter, ColumnCounter].Value = Person;
                            ColumnCounter++;
                        }                        

                        RowCounter++;
                    }

                    try
                    {
                        p.SaveAs(new FileInfo(Dialog.FileName));
                        MessageBox.Show("A beosztás generálása befejeződött.", "Információ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch
                    {
                        MessageBox.Show("A beosztás generálása nem sikerült!\n\nNincs véletlenül megnyitva Excelben?", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    try
                    {
                        System.Diagnostics.Process.Start(Dialog.FileName);
                    }
                    catch
                    {
                        MessageBox.Show("Nem sikerült megnyitnom a fájlt Excelben.\n\nPróbáld meg te légy szíves!", "Hiba", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void StartGroupComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            RefreshAssignment();

            StartGroupComboBox.Items.Clear();
            StartGroupComboBox.Text = "";

            foreach (var Group in MyAssignment.GroupList)
            {
                if ((Group.Count != 0) && (!string.IsNullOrWhiteSpace(Group[0])))
                {
                    StartGroupComboBox.Items.Add(Group[0]);
                }
            }            
        }
    }
}
