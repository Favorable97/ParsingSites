using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ParsingSites {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e) {

            string country = textBox1.Text.ToString();
            string ans;
            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/"+ country + "/ВВП-по-ППС-на-душу-населения");
            }
            List<string> year = new List<string>();
            List<string> value = new List<string>();
            string table = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));
            string ch = "&#160;";
            table = table.Replace(ch, " "); 
            
            Regex regex1 = new Regex(@"<td>(\w*)</td>");
            Regex regex2 = new Regex(@"<td>(\w*)\s(\w*)</td>");
            Regex regex3 = new Regex(@"<td>(\W*)(\w*),(\w*)</td>");
            Regex regex4 = new Regex(@"<td>(\w*)(\s*)(\w*),(\w*)</td>");
            MatchCollection matches = regex1.Matches(table);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
                    
            }
            matches = regex2.Matches(table);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
                    
            }

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/ИПЦ");
            }
            List<string> year1 = new List<string>();
            List<string> value1 = new List<string>();
            string table1 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));

            matches = regex1.Matches(table1);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year1.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex3.Matches(table1);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value1.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/topics/Демография/Смертность/Коэффициент-младенческой-смертности");
            }
            List<string> year2 = new List<string>();
            List<string> value2 = new List<string>();
            string table2 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));

            matches = regex1.Matches(table2);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year2.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex3.Matches(table2);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value2.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/Коэффициент-смертности");
            }
            List<string> year3 = new List<string>();
            List<string> value3 = new List<string>();
            string table3 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));

            matches = regex1.Matches(table3);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year3.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex3.Matches(table3);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value3.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/topics/Демография/Возраст/Население-в-возрасте-15-24");
            }
            List<string> year4 = new List<string>();
            List<string> value4 = new List<string>();
            string table4 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));
            table4 = table4.Replace(ch, " ");

            matches = regex1.Matches(table4);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year4.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex4.Matches(table4);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value4.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            } else {
                matches = regex3.Matches(table4);
                if (matches.Count > 0) {
                    foreach (Match match in matches) {
                        richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                        value4.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                    }
                }
            }

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/Коэффициент-рождаемости");
            }
            List<string> year5 = new List<string>();
            List<string> value5 = new List<string>();
            string table5 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));

            matches = regex1.Matches(table5);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year5.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex3.Matches(table5);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value5.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/topics/Образование/Высшее-образование/Валовой-показатель-охвата-Высшее-образование-МСКО-5-6");
            }
            List<string> year6 = new List<string>();
            List<string> value6 = new List<string>();
            string table6 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));

            matches = regex1.Matches(table6);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year6.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex3.Matches(table6);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value6.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/topics/Образование/Высшее-образование/Валовой-показатель-завершения-обучения-первая-ступень-высшего-образования-МСКО-5");
            }
            List<string> year7 = new List<string>();
            List<string> value7 = new List<string>();
            /*string table7 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));

            matches = regex1.Matches(table7);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year7.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex3.Matches(table7);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value7.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }*/

            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/" + country + "/Доля-городского-насления");
            }
            List<string> year8 = new List<string>();
            List<string> value8 = new List<string>();
            string table8 = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));

            matches = regex1.Matches(table8);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    year8.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }
            matches = regex3.Matches(table8);
            if (matches.Count > 0) {
                foreach (Match match in matches) {
                    richTextBox1.AppendText(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1) + ", ");
                    value8.Add(match.Value.Substring(match.Value.IndexOf('>') + 1, match.Value.LastIndexOf('<') - match.Value.IndexOf('>') - 1));
                }
            }

            Microsoft.Office.Interop.Excel.Application excel;
            Workbook excelworkBook;
            Worksheet excelSheet;
            Range excelCellrange;


            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;
            excel.DisplayAlerts = false;

            excelworkBook = excel.Workbooks.Add(Type.Missing);

            excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Rows[1].Columns[1] = "Год";
            excelSheet.Rows[1].Columns[2] = "ВВП по ППС на душу населения";
            excelSheet.Rows[1].Columns[3] = "Год";
            excelSheet.Rows[1].Columns[4] = "Инфляция(ИПЦ)";
            excelSheet.Rows[1].Columns[5] = "Год";
            excelSheet.Rows[1].Columns[6] = "Коэффициент младенческой смертности";
            excelSheet.Rows[1].Columns[7] = "Год";
            excelSheet.Rows[1].Columns[8] = "Коэффициент смертности";
            excelSheet.Rows[1].Columns[9] = "Год";
            excelSheet.Rows[1].Columns[10] = "Население в возрасте 15-24";
            excelSheet.Rows[1].Columns[11] = "Год";
            excelSheet.Rows[1].Columns[12] = "Коэффициент рождаемости";
            excelSheet.Rows[1].Columns[13] = "Год";
            excelSheet.Rows[1].Columns[14] = "Валовой показатель охвата Высшее образование МСКО-5,-6";
            excelSheet.Rows[1].Columns[15] = "Год";
            excelSheet.Rows[1].Columns[16] = "Валовой-показатель-завершения-обучения-первая-ступень-высшего-образования-МСКО-5";
            excelSheet.Rows[1].Columns[17] = "Год";
            excelSheet.Rows[1].Columns[18] = "Доля городского насления";

            for (int i = 0; i < year.Count; i++) {
                excelSheet.Rows[i + 2].Columns[1] = year[i];
                excelSheet.Rows[i + 2].Columns[2] = value[i];
            }

            for (int i = 0; i < year1.Count; i++) {
                excelSheet.Rows[i + 2].Columns[3] = year1[i];
                excelSheet.Rows[i + 2].Columns[4] = value1[i];
            }

            for (int i = 0; i < year2.Count; i++) {
                excelSheet.Rows[i + 2].Columns[5] = year2[i];
                excelSheet.Rows[i + 2].Columns[6] = value2[i];
            }

            for (int i = 0; i < year3.Count; i++) {
                excelSheet.Rows[i + 2].Columns[7] = year3[i];
                excelSheet.Rows[i + 2].Columns[8] = value3[i];
            }

            for (int i = 0; i < year4.Count; i++) {
                excelSheet.Rows[i + 2].Columns[9] = year4[i];
                excelSheet.Rows[i + 2].Columns[10] = value4[i];
            }

            for (int i = 0; i < year5.Count; i++) {
                excelSheet.Rows[i + 2].Columns[11] = year5[i];
                excelSheet.Rows[i + 2].Columns[12] = value5[i];
            }

            for (int i = 0; i < year.Count; i++) {
                excelSheet.Rows[i + 2].Columns[13] = year6[i];
                excelSheet.Rows[i + 2].Columns[14] = value6[i];
            }

            for (int i = 0; i < year7.Count; i++) {
                excelSheet.Rows[i + 2].Columns[15] = year7[i];
                excelSheet.Rows[i + 2].Columns[16] = value7[i];
            }

            for (int i = 0; i < year8.Count; i++) {
                excelSheet.Rows[i + 2].Columns[17] = year8[i];
                excelSheet.Rows[i + 2].Columns[18] = value8[i];
            }

            excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[year.Count + 1, 18]];
            excelCellrange.EntireColumn.AutoFit();
            Borders border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            excel.AlertBeforeOverwriting = false;
            excelworkBook.SaveAs(@"C:\Users\melikyan\Desktop\" + country + ".xlsx");
            excel.Quit();
            //richTextBox1.AppendText("\n" + table);
        }
    }
}
