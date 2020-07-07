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
            string ans;
            using (WebClient wc = new WebClient { Encoding = Encoding.UTF8 }) {
                ans = wc.DownloadString("https://knoema.ru/atlas/Венгрия/ВВП-на-душу-населения");
            }
            List<string> year = new List<string>();
            List<string> value = new List<string>();
            string table = ans.Substring(ans.IndexOf("<table"), ans.IndexOf("</table>") - ans.IndexOf("<table"));
            string ch = "&#160;";
            table = table.Replace(ch, " "); 
            
            Regex regex1 = new Regex(@"<td>(\w*)</td>");
            Regex regex2 = new Regex(@"<td>(\w*)\s(\w*)</td>");
            Regex regex3 = new Regex(@"<td>(\w*),(\w*)</td>");
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
                ans = wc.DownloadString("https://knoema.ru/atlas/Венгрия/Уровень-бедности");
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
                ans = wc.DownloadString("https://knoema.ru/atlas/Венгрия/Коэффициент-рождаемости");
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
                ans = wc.DownloadString("https://knoema.ru/atlas/Венгрия/Коэффициент-смертности");
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
            excelSheet.Rows[1].Columns[2] = "ВВП на душу населения";
            excelSheet.Rows[1].Columns[3] = "Уровень бедности";
            excelSheet.Rows[1].Columns[4] = "Коэффициент-рождаемости";
            excelSheet.Rows[1].Columns[5] = "Уровень бедности";

            for (int i = 0; i < year.Count; i++) {
                excelSheet.Rows[i + 2].Columns[1] = year[i];
                excelSheet.Rows[i + 2].Columns[2] = value[i];
            }

            excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[year.Count + 1, 2]];
            excelCellrange.EntireColumn.AutoFit();
            Borders border = excelCellrange.Borders;
            border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            excel.AlertBeforeOverwriting = false;
            excelworkBook.SaveAs(@"C:\Users\melikyan\Desktop\Венгрия.xlsx");
            excel.Quit();
            //richTextBox1.AppendText("\n" + table);
        }
    }
}
