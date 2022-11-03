using Microsoft.Win32;
using System;
using System.Windows.Forms;

namespace WordTools
{
    public static class PublicTools
    {
        /// <param 文件列表="FileNames"></param>
        private static string[] fileNames;
        public static string[] FileNames
        {
            get
            {
                if (fileNames == null || fileNames[0] == "")
                {
                    MessageBox.Show("请先选择文件再进行其他操作！");
                    return null;
                }
                else
                {
                    return fileNames;
                }
            }
            set { fileNames = value; }
        }
        #region 文件选择框
        public static string[] SelectWordFile()
        {
            OpenFileDialog dlg = new OpenFileDialog()
            {
                Multiselect = true,
                RestoreDirectory = true,
                Filter = "Word文件|*.docx"
            };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                return dlg.FileNames;
            }
            return null;
        }
        #endregion
        #region 保存文件
        public static string SaveFile()
        {
            SaveFileDialog dlg = new SaveFileDialog()
            {
                Filter = "Word 文件(*.docx)|*.docx|Word 文件(*.doc)|*.doc",
                FilterIndex = 0,
                RestoreDirectory = true, //保存对话框是否记忆上次打开的目录
                                         //saveFileDialog.CreatePrompt = true;
                Title = "导出Word文件到"
            };
            DateTime now = DateTime.Now;
            dlg.FileName = "Word" + now.Year.ToString().PadLeft(2) + now.Month.ToString().PadLeft(2, '0') + now.Day.ToString().PadLeft(2, '0') + "-" + now.Hour.ToString().PadLeft(2, '0') + now.Minute.ToString().PadLeft(2, '0') + now.Second.ToString().PadLeft(2, '0');
            return dlg.FileName;
        }
        #endregion
    }
}