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
    public partial class Getparameter : Form
    {

        private System.Windows.Forms.TextBox[] imgName,chainLength,overllLength,By,gyts;
        private System.Windows.Forms.Label[] projectName;
        private System.Windows.Forms.Button button;
        private TableLayoutPanel TLP;

        public XlsxInfo[] xlsxInfo;

        public Getparameter(String[,] xlsxInfo)
        {
            InitializeComponent();
            this.Resize += new System.EventHandler(this.Form_Resize);

            // 0 = xlsxName; 1 = imgName; 2 = chainLength; 3 = overallLength; 4 = Dot B y; 
            this.xlsxInfo = new XlsxInfo[xlsxInfo.GetLength(0)];
            TLP = new TableLayoutPanel();
            TLP.Location = new Point(0, 0);
            TLP.Size = this.ClientSize;
            TLP.AutoScroll = true;
            TLP.ColumnCount = 6; TLP.RowCount = this.xlsxInfo.Length;
            this.Controls.Add(TLP);


            projectName = new Label[this.xlsxInfo.Length];
            gyts = new TextBox[this.xlsxInfo.Length];
            imgName = new TextBox[this.xlsxInfo.Length];
            chainLength = new TextBox[this.xlsxInfo.Length];
            overllLength = new TextBox[this.xlsxInfo.Length];
            By = new TextBox[this.xlsxInfo.Length];
            Int32 s = 1;
            // Processing incoming
            for (Int32 i = 0; i < this.xlsxInfo.Length; i++)
            {
                Int32 x = i;
                this.xlsxInfo[i].projectName = xlsxInfo[i, 0];
                this.xlsxInfo[i].gyts = Int32.Parse(xlsxInfo[i, 2].Remove(xlsxInfo[i, 2].Length-1,1).Remove(0,5));
                if (i > 0 && xlsxInfo[i - s, 0] == xlsxInfo[i, 0])
                {
                    s++;
                    this.xlsxInfo[i].xlsxName = xlsxInfo[i, 0] + s.ToString("D3") + "曲线图.xlsx";
                }
                else
                {
                    this.xlsxInfo[i].xlsxName = xlsxInfo[i, 0] + "曲线图.xlsx";
                    if (s != 1)
                    {
                        this.xlsxInfo[i - s].xlsxName = xlsxInfo[i - s, 0] + "001曲线图.xlsx";
                        this.projectName[i - s].Text = this.xlsxInfo[i - s].xlsxName;
                        s = 1;
                    }
                }
                this.xlsxInfo[i].chainLength = Double.Parse(xlsxInfo[i, 1]) * 1000;

                // hypothesis
                Random random = new Random();
                this.xlsxInfo[i].imgName = xlsxInfo[i, 0];
                this.xlsxInfo[i].overallLength = this.xlsxInfo[i].chainLength * random.Next(200,250) / 100;
                this.xlsxInfo[i].overallLength = this.xlsxInfo[i].overallLength < 1020 ? 1020 : this.xlsxInfo[i].overallLength;
                this.xlsxInfo[i].By = Convert.ToDouble(new Random().Next(12000, 17000)) / 1000;

                this.projectName[i] = new Label();
                this.projectName[i].AutoSize = true;
                this.projectName[i].Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
                this.projectName[i].Text = this.xlsxInfo[i].xlsxName;
                TLP.Controls.Add(this.projectName[i],0,i);

                
                this.imgName[i] = new TextBox();
                this.imgName[i].Width = 300;
                this.imgName[i].SelectedText = this.xlsxInfo[i].imgName;
                this.imgName[i].TextChanged += (o, e) => this.xlsxInfo[x].imgName = 
                    imgName[x].Text ?? this.xlsxInfo[x].imgName;
                this.imgName[i].TabIndex = i;
                TLP.Controls.Add(this.imgName[i], 1, i);

                this.chainLength[i] = new TextBox();
                this.chainLength[i].Width = 80;
                this.chainLength[i].SelectedText = this.xlsxInfo[i].chainLength.ToString("0.00");
                this.chainLength[i].TextChanged += (o, e) => this.xlsxInfo[x].chainLength = chainLength[x].Text != "" ? 
                    Double.Parse(chainLength[x].Text): 
                    this.xlsxInfo[x].chainLength;
                this.chainLength[i].TabIndex = i;
                TLP.Controls.Add(this.chainLength[i], 2, i);

                this.overllLength[i] = new TextBox();
                this.overllLength[i].Width = 80;
                this.overllLength[i].SelectedText = this.xlsxInfo[i].overallLength.ToString("0.00");
                this.overllLength[i].TextChanged += (o, e) => this.xlsxInfo[x].overallLength = overllLength[x].Text != "" ?
                    Double.Parse(overllLength[x].Text): 
                    this.xlsxInfo[x].overallLength;
                this.overllLength[i].TabIndex = i;
                TLP.Controls.Add(this.overllLength[i], 3, i);

                this.By[i] = new TextBox();
                this.By[i].Width = 80;
                this.By[i].SelectedText = this.xlsxInfo[i].By.ToString("0.00");
                this.By[i].TextChanged += (o, e) => this.xlsxInfo[x].By = By[x].Text != "" ?
                    Double.Parse(By[x].Text) :
                    this.xlsxInfo[x].By;
                this.By[i].TabIndex = i;
                TLP.Controls.Add(this.By[i], 4, i);

                this.gyts[i] = new TextBox();
                this.gyts[i].Width = 50;
                this.gyts[i].SelectedText = this.xlsxInfo[i].gyts.ToString();
                this.gyts[i].TextChanged += (o, e) => this.xlsxInfo[x].gyts = gyts[x].Text != "" ?
                    Int32.Parse(gyts[x].Text):
                    this.xlsxInfo[x].gyts;
                this.gyts[i].TabIndex = i;
                TLP.Controls.Add(this.gyts[i], 5, i);
            }


            this.Visible = false;
            


            this.button = new Button();
            this.button.Location = new System.Drawing.Point((this.Width - 90) /2,30 + this.xlsxInfo.Length * 20);
            this.button.Name = "button";
            this.button.Size = new System.Drawing.Size(90, 23);
            this.button.TabIndex = 3;
            this.button.Text = "確認無誤";
            this.button.UseVisualStyleBackColor = true;
            this.Anchor = AnchorStyles.Bottom;
            this.button.Click += new System.EventHandler(this.button_Click);
            TLP.Controls.Add(this.button);
        }

        private void button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form_Resize(object sender, System.EventArgs e)
        {
            TLP.Size = ClientSize;
        }
    }
}
