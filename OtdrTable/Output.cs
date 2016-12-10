using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing.Imaging;

using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using OtdrTable.NPOI_Ex;
using System.Drawing;

namespace OtdrTable {
    class Output {
        private IWorkbook Workbook;
        private ISheet ActiveSheet;
        ICellStyle DefaultStyle;
        private XlsxInfo Info;

        //-點信息[x,y]
        // coordinate {
        //   x : 點序號 {
        //     0 : 該點對應的縱坐標/dB,
        //     1 : 該點對應的實際行坐標/m
        //-}
        private Double[,] coordinate;
        private Int32 BKey; // B點在coordinate中的索引
        private Double TLR; // 鏈衰減係數

        private Random random = new Random();
        static private Object ConsoleLock = new Object();

        static public void ShowInfo(XlsxInfo[] xlsxInfo) {
            Int32 gytsSum = 0;
            for (Int32 i = 0; i < xlsxInfo.Length; i++)
                gytsSum += xlsxInfo[i].gyts;
            Conex.DebugLine("Project sum: {0}; Gyts sum: {1}.", xlsxInfo.Length, gytsSum);
            for (Int32 i = 0; i < xlsxInfo.Length; i++)
                Conex.WriteLine("{0,2}. Wait   0/{1,-3} {2,10:F2} {3,10:F2} {4,6:F3} {5}",
                    (i + 1).ToString(),
                    xlsxInfo[i].gyts,
                    xlsxInfo[i].chainLength,
                    xlsxInfo[i].overallLength,
                    xlsxInfo[i].By,
                    xlsxInfo[i].imgName);
        }

        public void ExportingXlsx(XlsxInfo info, Int32 LineNum) {
            lock (ConsoleLock) {
                Conex.SetCursorPosition(4, LineNum);
                Conex.Warn("EXEC");
            }
            try {
                this.Workbook = new XSSFWorkbook();
                this.Info = info;
                CreateFile(LineNum);
                Workbook.Save(info.xlsxName);
                lock (ConsoleLock) {
                    Conex.SetCursorPosition(4, LineNum);
                    Conex.Info("Done");
                }
            }
            catch {
                lock (ConsoleLock) {
                    Conex.SetCursorPosition(4, LineNum);
                    Conex.Error("Fail");
                }
            }
            return;
        }

        public void CreateFile(Int32 LineNum) {

            TLR = 0.25 + Convert.ToDouble(random.Next(5, 95)) / 1000; // total loss ratio of overal

            CurveGraph CG = this.markGraphBase();

            // GOTO: ColFmt.Font.Family = 3;
            // GOTO: ColFmt.Font.Scheme = TFontScheme.None;

            IFont font_8 = Workbook.CreateFont();
            font_8.FontName = "宋体";
            font_8.FontHeight = 8;

            IFont font_10 = Workbook.CreateFont();
            font_10.FontName = "宋体";
            font_10.FontHeight = 10;

            ICellStyle DefaultStyle = Workbook.CreateCellStyle();
            DefaultStyle.Alignment = HorizontalAlignment.Left;
            DefaultStyle.SetFont(font_8);

            ICellStyle style_centen_8 = Workbook.CreateCellStyle();
            style_centen_8.Alignment = HorizontalAlignment.Center;
            style_centen_8.SetFont(font_8);

            ICellStyle style_centen_10 = Workbook.CreateCellStyle();
            style_centen_10.Alignment = HorizontalAlignment.Center;
            style_centen_10.ShrinkToFit = false;
            style_centen_10.SetFont(font_10);

            this.DefaultStyle = DefaultStyle;

            Int32 i = 1;
            for (Int32 s = 0; s < Math.Max(1, (Info.gyts) / 6); s++) {

                ActiveSheet = Workbook.CreateSheet((1 + 6 * s).ToString() + "-" + (6 + s * 6).ToString());

                // GOTO: xls.OptionsCheckCompatibility = false;
                // GOTO: xls.PrintOptions = TPrintOptions.Orientation | TPrintOptions.NoPls;

                // Fix P#0; Sheet.DefaultColumnWidth = 2720;
                ActiveSheet.DefaultRowHeight = 270;

                for (Int32 colnum = 0; colnum < 9; colnum++)
                    ActiveSheet.SetColumnWidth(colnum, 2720);
                ActiveSheet.SetColumnWidth(4, 672);

                Int32 gyts = (s+1) * 6 > Info.gyts ? Info.gyts % 6 : 6;
                for (Int32 j = 0; j < gyts; j++, i++) {

                    Int32 AKey;
                    Double A2B, A2BTLP;


                    Int32 OffsetX = j / 2 * 8,
                          OffsetY = j % 2 * 5;
                    AKey = (Int32)(BKey / Convert.ToDouble(random.Next(250, 400) / 100.00));
                    A2B = coordinate[BKey, 0] - coordinate[AKey, 0];
                    A2BTLP = (coordinate[AKey, 1] - coordinate[BKey, 1]) / A2B * 1000;
                    A2BTLP = TLR - random.Next(1, 99) / 1000.00;
                    if (j % 2 == 0) {
                        SetRowHeight(0 + OffsetX, 540);
                        SetRowHeight(2 + OffsetX, 2700);
                    }
                    
                    MergeCells(0 + OffsetX, 0 + OffsetX, 0 + OffsetY, 3 + OffsetY);
                    MergeCells(1 + OffsetX, 1 + OffsetX, 0 + OffsetY, 3 + OffsetY);
                    MergeCells(2 + OffsetX, 2 + OffsetX, 0 + OffsetY, 3 + OffsetY);
                    MergeCells(3 + OffsetX, 3 + OffsetX, 0 + OffsetY, 3 + OffsetY);


                    SetCellStyle(1 + OffsetX, 0 + OffsetY, style_centen_10);
                    SetCellValue(1 + OffsetX, 0 + OffsetY, Info.imgName + i.ToString("D3"));
                    
                    AddPicture(CG.GetImg(coordinate[AKey, 0]), 0, 0, -10000, 0, 0 + OffsetY, 2 + OffsetX, 4 + OffsetY, 3 + OffsetX);                    

                    SetCellStyle(3 + OffsetX, 0 + OffsetY, style_centen_8);
                    SetCellValue(3 + OffsetX, 0 + OffsetY, "曲线信息");

                    SetCellValue(4 + OffsetX, 0 + OffsetY, "A-B 距离:");
                    SetCellValue(4 + OffsetX, 1 + OffsetY, ((Int32)A2B).ToString() + " m");

                    SetCellValue(4 + OffsetX, 2 + OffsetY, "A-B 衰减;");
                    SetCellValue(4 + OffsetX, 3 + OffsetY, (A2BTLP * A2B / 1000).ToString("0.000") + " dB");

                    SetCellValue(5 + OffsetX, 0 + OffsetY, "A-B 衰减系数;");
                    SetCellValue(5 + OffsetX, 1 + OffsetY, A2BTLP.ToString("0.000") + " dB/km");

                    SetCellValue(5 + OffsetX, 2 + OffsetY, "A 纵坐标;");
                    SetCellValue(5 + OffsetX, 3 + OffsetY, coordinate[AKey, 1].ToString("0.000") + " dB");

                    SetCellValue(6 + OffsetX, 0 + OffsetY, "链长;");
                    SetCellValue(6 + OffsetX, 1 + OffsetY, Info.chainLength.ToString() + " m");

                    SetCellValue(6 + OffsetX, 2 + OffsetY, "链衰减;");
                    SetCellValue(6 + OffsetX, 3 + OffsetY, (TLR * Info.chainLength / 1000).ToString("0.000") + " dB");

                    SetCellValue(7 + OffsetX, 0 + OffsetY, "链衰减系数;");
                    SetCellValue(7 + OffsetX, 1 + OffsetY, TLR.ToString("0.000") + " dB/km");

                    SetCellValue(7 + OffsetX, 2 + OffsetY, "事件数量;");
                    SetCellValue(7 + OffsetX, 3 + OffsetY, 2);

                    // GOTO: xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Microsoft");
                    lock (ConsoleLock) {
                        Conex.SetCursorPosition(9, LineNum);
                        if (i == Info.gyts) Conex.Info("{0,3}", i.ToString());
                        else Conex.Warn("{0,3}", i.ToString());
                    }
                }
            }
        }

        private IPicture AddPicture(Bitmap pic, Int32 dx1, Int32 dy1, Int32 dx2, Int32 dy2, Int32 col1, Int32 row1, Int32 col2, Int32 row2) {
            return ActiveSheet.AddPicture(pic, dx1, dy1, dx2, dy2, col1, row1, col2, row2);
        }

        private void SetRowHeight(Int32 row, Int16 height) {
            ActiveSheet.SetRowHeight(row, height);
        }

        private void SetCellStyle(Int32 row, Int32 col, ICellStyle style) {
            ActiveSheet.SetCellStyle(row, col, style);
        }

        private ICell SetCellValue(Int32 row, Int32 col, String value) {
            return ActiveSheet.SetCellValue(row, col, value, DefaultStyle);
        }

        private ICell SetCellValue(Int32 row, Int32 col, Int32 value) {
            return ActiveSheet.SetCellValue(row, col, value, DefaultStyle);
        }

        private Int32 MergeCells(Int32 firstRow, Int32 lastRow, Int32 firstCol, Int32 lastCol) {
            return ActiveSheet.MergeCells(firstRow, lastRow, firstCol, lastCol);
        }

        // 生成底圖。
        // De: 即對初始化CurveGraph。獲得比成圖缺少A點標記的對象。
        public CurveGraph markGraphBase() {
            CurveGraph CG = new CurveGraph();
            makeCoordinate(Info.chainLength, Info.overallLength, TLR, Info.By);
            CG.InitImg(Info.chainLength, Info.overallLength, coordinate);
            return CG;
        }

        private void makeCoordinate(Double chainLength, Double overallLength, Double TLR, Double By) {
            Int32 coordinateL = random.Next((Int32)overallLength * 2, (Int32)overallLength * 3);
            if (coordinateL > 10000) coordinateL = 10000;
            coordinate = new Double[coordinateL, 2];
            Boolean peak = true;
            Double Ay = By + (TLR * chainLength / 1000);
            coordinate[1, 0] = 0; coordinate[1, 1] = Ay; // A
            coordinate[0, 0] = chainLength; coordinate[0, 1] = By; // B
            Double x = 0, y = Ay;
            Int32 i = 2;
            do {
                if (i >= coordinate.GetLength(0)) i = 2;
                Double a = ((Double)random.Next(1, 100)) / 100;
                coordinate[i, 0] += a;
                x += a;
                i++;
            } while (x < overallLength);
            coordinate[coordinate.GetLength(0) - 1, 0] = overallLength;
            x = 0;
            // Double keepLength = ((Double)random.Next(13, 17))/10;
            Int32 px = coordinate.Length / 10 / 60;
            for (i = 2; i < coordinate.GetLength(0) - 1; i++) {
                if (peak && x > coordinate[0, 0]) {
                    Double N = coordinate[0, 1] * random.Next(1400, 1600) / 1000.00, //波峰高度
                           M = N - (coordinate[0, 1] / 5),
                           O1 = random.Next(10, 30) / 10,
                           O2 = random.Next(10, 30) / 10,
                           O3 = random.Next(10, 30) / 10,
                           O4 = random.Next(10, 30) / 10,
                           O5 = random.Next(10, 30) / 10,
                           O6 = random.Next(10, 30) / 10;

                    for (Int32 j = 0; j < px * 16; j++, i++) // 用一个像素登上波峰，然后七个像素跌落到后续值
                    {
                        coordinate[i, 0] += x;
                        x = coordinate[i, 0];
                        if (j < px) coordinate[i, 1] = N + O1;
                        else if (j < px * 4) coordinate[i, 1] = N - M * 1 / 2 + O2;
                        else if (j < px * 6) coordinate[i, 1] = N - M * 3 / 4 + O3;
                        else if (j < px * 8) coordinate[i, 1] = N - M * 7 / 8 + O4;
                        else if (j < px * 10) coordinate[i, 1] = N - M * 15 / 16 + O5;
                        else if (j < px * 12) coordinate[i, 1] = N - M * 31 / 32 + O6;
                        else if (j < px * 14) coordinate[i, 1] = coordinate[0, 1] / random.Next(1500, 2000) * 1000;
                        else coordinate[i, 1] = coordinate[0, 1] / random.Next(1500, 4000) * 1000;
                    }
                    peak = false;
                }
                coordinate[i, 0] += x;
                x = coordinate[i, 0];
                if (x < coordinate[0, 0] + px * 2)//* keepLength)
                {
                    // if (random.Next(1,3) == 1) TLR -= Convert.ToDouble(random.Next(4, 10)) / 1000;
                    // else TLR += Convert.ToDouble(random.Next(4, 10)) / 1000;
                    // coordinate[i, 1] = coordinate[0, 1] + TLR * (coordinate[0, 0] - coordinate[i, 0]) / 1000;
                    coordinate[i, 1] = By;
                }
                else coordinate[i, 1] = coordinate[0, 1] / random.Next(1500, 7000) * 1000;
                if (x < coordinate[0, 0]) BKey = i;
                coordinate[i, 1] = coordinate[i, 1] < 0 ? 0 : coordinate[i, 1];
            }
        }

    }
}
