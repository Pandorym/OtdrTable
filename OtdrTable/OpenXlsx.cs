using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Xlsx2OtdrTable
{
    public partial class OpenXlsx : Form
    {
        public OpenXlsx()
        {
            InitializeComponent();
        }

        private void OpenXlsx_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择结算表";
            fileDialog.Filter = "所有文件（*.*）|*.*";
            if (fileDialog.ShowDialog() == DialogResult.OK) fileNames = fileDialog.FileNames;
            textBox1.Text = String.Join(", ", fileNames);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            readXlsx();
            Getparameter getParameter = new Getparameter(xlsxInfo);
            getParameter.ShowDialog();
            for (Int32 i = 0; i < getParameter.xlsxInfo.Length; i++)
                Xlsx.ExportingXlsx(getParameter.xlsxInfo[i]);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        static private String[] fileNames;
        static private String[,] xlsxInfo;
        

        static private void readXlsx()
        {
            String[,] xlsxnames = new String[500, 3];
            Int32 i = 0;
            foreach (String fileName in fileNames)
            {
                String[,] names = Xlsx.ReadXlsx(fileName);
                for (Int32 a = 0; a < names.GetLength(0); a++)
                {
                    xlsxnames[i, 0] = names[a, 0];
                    xlsxnames[i, 1] = names[a, 1];
                    xlsxnames[i, 2] = names[a, 2];
                    i++;
                }
            }
            xlsxInfo = new String[i, 3];
            for (Int32 a = 0; a < i; a++)
            {
                xlsxInfo[a, 0] = xlsxnames[a, 0];
                xlsxInfo[a, 1] = xlsxnames[a, 1];
                xlsxInfo[a, 2] = xlsxnames[a, 2];
            }
        }

    }
}

