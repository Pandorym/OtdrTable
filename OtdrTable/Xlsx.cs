using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing.Imaging;
using FlexCel.Core;
using FlexCel.XlsAdapter;

namespace OtdrTable {
    static class Xlsx {
        static Random random = new Random();
        static private Double[,] coordinate;
        static private Int32 BKey;

        static public void ExportingXlsx(XlsxInfo info, Int32 LineNum) {
            Console.SetCursorPosition(4, LineNum);
            Conex.Warn("EXEC");
            try {
                XlsFile xls = new XlsFile(true);
                CreateFile(xls, info, LineNum);
                xls.Save(info.xlsxName);
                Console.SetCursorPosition(4, LineNum);
                Conex.Info("Done");
            }
            catch {
                Console.SetCursorPosition(4, LineNum);
                Conex.Error("Fail");
            }
            return;
        }

        static void CreateFile(XlsFile xls, XlsxInfo info, Int32 LineNum) {
            xls.NewFile(1, TExcelFileFormat.v2016);

            for (Int32 s = 0; s < Math.Max(1, (info.gyts) / 6); s++) {

                xls.ActiveSheet = s + 1;
                xls.SheetName = (1 + 6 * s).ToString() + "-" + (6 + s * 6).ToString();
                if (s != (info.gyts / 6) - 1) xls.AddSheet();

                xls.OptionsCheckCompatibility = false;
                xls.PrintOptions = TPrintOptions.Orientation | TPrintOptions.NoPls;

                xls.DefaultColWidth = 2720;
                xls.DefaultRowHeight = 270;
                xls.DefaultRowHeightAutomatic = false;

                TFlxFormat ColFmt;
                ColFmt = xls.GetFormat(xls.GetColFormat(1));
                ColFmt.Font.Name = "宋体";
                ColFmt.Font.Family = 3;
                ColFmt.Font.Size20 = 160;
                ColFmt.Font.Scheme = TFontScheme.None;
                ColFmt.HAlignment = THFlxAlignment.left;
                xls.SetFormat(0, ColFmt);


                xls.SetColWidth(5, 672);

                Int32 i = 0;

                Int32 gyts = (s+1) * 6 > info.gyts ? info.gyts % 6 : 6;
                for (Int32 j = 0; j < gyts; j++, i++) {

                    Double TLR = 0.25 + Convert.ToDouble(random.Next(5, 95)) / 1000; // total loss ratio
                    makeCoordinate(info.chainLength, info.overallLength, TLR, info.By);
                    CurveGraph.InitImg(info.chainLength, info.overallLength, coordinate);
                    Int32 AKey;
                    Double A2B, A2BTLP;


                    Int32 OffsetX = i / 2 * 8,
                          OffsetY = i % 2 * 5;
                    AKey = (Int32)(BKey / Convert.ToDouble(random.Next(250, 400) / 100.00));
                    A2B = coordinate[BKey, 0] - coordinate[AKey, 0];
                    A2BTLP = (coordinate[AKey, 1] - coordinate[BKey, 1]) / A2B * 1000;
                    A2BTLP = TLR - random.Next(1, 99) / 1000.00;
                    if (i % 2 == 0) {
                        xls.SetRowHeight(1 + OffsetX, 540);
                        xls.SetAutoRowHeight(2 + OffsetX, false);
                        xls.SetRowHeight(3 + OffsetX, 2700);
                        xls.SetAutoRowHeight(4 + OffsetX, false);
                        xls.SetAutoRowHeight(5 + OffsetX, false);
                        xls.SetAutoRowHeight(6 + OffsetX, false);
                        xls.SetAutoRowHeight(7 + OffsetX, false);
                        xls.SetAutoRowHeight(8 + OffsetX, false);
                    }

                    xls.MergeCells(1 + OffsetX, 1 + OffsetY, 1 + OffsetX, 4 + OffsetY);
                    xls.MergeCells(2 + OffsetX, 1 + OffsetY, 2 + OffsetX, 4 + OffsetY);
                    xls.MergeCells(3 + OffsetX, 1 + OffsetY, 3 + OffsetX, 4 + OffsetY);
                    xls.MergeCells(4 + OffsetX, 1 + OffsetY, 4 + OffsetX, 4 + OffsetY);

                    TFlxFormat fmt;

                    fmt = xls.GetCellVisibleFormatDef(2 + OffsetX, 1 + OffsetY);
                    fmt.HAlignment = THFlxAlignment.center;
                    fmt.Font.Size20 = 200;
                    xls.SetCellFormat(2 + OffsetX, 1 + OffsetY, xls.AddFormat(fmt));
                    xls.SetCellValue(2 + OffsetX, 1 + OffsetY, info.imgName + (s * 6 + i + 1).ToString("D3"));


                    using (MemoryStream stream = new MemoryStream()) {
                        CurveGraph.GetImg(coordinate[AKey, 0]).Save(stream, ImageFormat.Jpeg);
                        TImageProperties ImgProps = new TImageProperties();
                        ImgProps.Anchor = new TClientAnchor(TFlxAnchorType.MoveAndDontResize, 3 + OffsetX, 0, 1 + OffsetY, 0, 4 + OffsetX, 0, 5 + OffsetY, 0);
                        ImgProps.ShapeName = "Img" + i.ToString("D3");
                        xls.AddImage(stream, ImgProps);
                    }


                    fmt = xls.GetCellVisibleFormatDef(4 + OffsetX, 1 + OffsetY);
                    fmt.HAlignment = THFlxAlignment.center;
                    xls.SetCellFormat(4 + OffsetX, 1 + OffsetY, xls.AddFormat(fmt));
                    xls.SetCellValue(4 + OffsetX, 1 + OffsetY, "曲线信息");

                    xls.SetCellValue(5 + OffsetX, 1 + OffsetY, "A-B 距离:");
                    xls.SetCellValue(5 + OffsetX, 2 + OffsetY, ((Int32)A2B).ToString() + " m");

                    xls.SetCellValue(5 + OffsetX, 3 + OffsetY, "A-B 衰减;");
                    xls.SetCellValue(5 + OffsetX, 4 + OffsetY, (A2BTLP * A2B / 1000).ToString("0.000") + " dB");

                    xls.SetCellValue(6 + OffsetX, 1 + OffsetY, "A-B 衰减系数;");
                    xls.SetCellValue(6 + OffsetX, 2 + OffsetY, A2BTLP.ToString("0.000") + " dB/km");

                    xls.SetCellValue(6 + OffsetX, 3 + OffsetY, "A 纵坐标;");
                    xls.SetCellValue(6 + OffsetX, 4 + OffsetY, coordinate[AKey, 1].ToString("0.000") + " dB");

                    xls.SetCellValue(7 + OffsetX, 1 + OffsetY, "链长;");
                    xls.SetCellValue(7 + OffsetX, 2 + OffsetY, info.chainLength.ToString() + " m");

                    xls.SetCellValue(7 + OffsetX, 3 + OffsetY, "链衰减;");
                    xls.SetCellValue(7 + OffsetX, 4 + OffsetY, (TLR * info.chainLength / 1000).ToString("0.000") + " dB");

                    xls.SetCellValue(8 + OffsetX, 1 + OffsetY, "链衰减系数;");
                    xls.SetCellValue(8 + OffsetX, 2 + OffsetY, TLR.ToString("0.000") + " dB/km");

                    xls.SetCellValue(8 + OffsetX, 3 + OffsetY, "事件数量;");
                    xls.SetCellValue(8 + OffsetX, 4 + OffsetY, 2);

                    xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Microsoft");


                    Console.SetCursorPosition(9, LineNum);
                    if (i + 1 == gyts) Conex.Info("{0,3}", gyts);
                    else Conex.Warn("{0,3}", i + 1);
                }
            }
        }

        static private void makeCoordinate(Double chainLength, Double overallLength, Double TLR, Double By) {
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

        static public String[,] ReadXlsx(String filePath) {
            XlsFile xls = new XlsFile(filePath);
            xls.ActiveSheetByName = "工程量明细表";
            Int32 XF = -1, rowCount = 0; ;
            Int32 c = 0;

            // 獲得項目數，即有效行數
            for (Int32 row = 1; row < xls.RowCount; row++) {
                if (xls.GetCellValueIndexed(row, 1, ref XF) == null) c++;
                else c = 0;
                // 往下十行為空 視為文件結束
                if (c.Equals(10)) {
                    rowCount = row - 10;
                    break;
                }
            }
            rowCount = rowCount == 0 ? xls.RowCount - 5 : rowCount - 5; // 有效行數 = 總行數 - 頭三行 - 尾二行

            String[,] str = new String[rowCount, 3];
            String notEmpty = String.Empty;
            for (Int32 row = 4; row < rowCount + 4; row++) {
                str[row - 4, 0] = xls.GetCellValueIndexed(row, 3, ref XF) != null ?
                    xls.GetCellValueIndexed(row, 3, ref XF).ToString() : notEmpty;
                str[row - 4, 1] = xls.GetCellValueIndexed(row, 5, ref XF).ToString();
                str[row - 4, 2] = xls.GetCellValueIndexed(row, 4, ref XF).ToString();
                notEmpty = str[row - 4, 0];
            }
            return str;
        }

        static public String[,] ReadTxt(String filePath) {
            StreamReader sr = new StreamReader(filePath, Encoding.UTF8);
            String[] Line = sr.ReadToEnd().Split('\n');
            String[,] str = new String[Line.Length,3];
            for (Int32 i = 0; i < Line.Length; i++) {
                if (Line[i][Line[i].Length - 1] == '\r') Line[i] = Line[i].Substring(0, Line[i].Length - 1);
                str[i, 0] = Line[i].Split(' ')[0];
                str[i, 2] = Line[i].Split(' ')[1];
                str[i, 1] = Line[i].Split(' ')[2];
            }
            return str;
        }
    }
}
