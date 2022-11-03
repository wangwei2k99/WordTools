using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
namespace WordTools
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Application wapp = new Word.Application();
            try
            {
                PublicTools.FileNames = PublicTools.SelectWordFile();
                if (PublicTools.FileNames is null)
                {
                    MessageBox.Show($"请选择正确的Word文件!");
                    return;
                }
                string log = "";
                foreach (var fileName in PublicTools.FileNames)
                {
                    Word.Document docx = wapp.Documents.Open(fileName);
                    string text = "", grzf = "";
                    text = docx.Content.Text;
                    string ysk;
                    if (Regex.IsMatch(text, @"预收款:*\d+(\.\d+)\s*起付线:*\d+(\.\d+)?"))
                    {
                        ysk = Regex.Match(text, @"预收款:*\d+(\.\d+)\s*起付线:*\d+(\.\d+)?").ToString();
                    }
                    else
                    {
                        ysk = Regex.Match(text, @"预收款:*\d+(\.\d+)").ToString();
                    }
                    if (ysk == "")
                    {
                        log = $"{log}\r失败：{fileName}";
                        continue;
                    }
                    else
                    {
                        log = $"{log}\r成功：{fileName}";
                    }
                    grzf = Regex.Match(text, @"个人支付:*\d+(\.\d+)").ToString();
                    double fk = Convert.ToDouble(Regex.Match(grzf, @"\d+(\.\d+)").ToString());
                    SearchReplace(ref docx, ysk, "预收款:0.00 起付线:600.00");
                    string text1;
                    if (fk - 600 >= 0)
                    {
                        text1 = $"【补交】:{fk - 600:f2}（现金）";
                    }
                    else
                    {
                        text1 = $"退款:{Math.Abs(fk - 600):f2}";
                    }
                    string tkxx = Regex.Match(text, @"(【补交】:*|退款:*)\d+(\.\d+)(（现金）)*").ToString();
                    SearchReplace(ref docx, tkxx, text1);
                    docx.Save();
                    docx.Close();

                }
                MessageBox.Show(log);
                wapp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                wapp.Quit();
            }
        }
        private void SearchReplace(ref Word.Document docx, string find, string replace)
        {

            Word.Find findObject = docx.Content.Find;
            findObject.ClearFormatting();
            findObject.Text = find;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replace;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            findObject.Execute(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                ref replaceAll, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        }
    }
}
