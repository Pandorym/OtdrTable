using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace OtdrTable
{
    class CurveGraph
    {
        private Int32[,] dot;
        private Double yDiv;
        private Graphics g;
        private Bitmap Img;

        public Bitmap InitImg(Double chainLength, Double overallLength, Double[,] coordinate)
        {
            Img = new Bitmap(601, 276);
            g = Graphics.FromImage(Img);
            yDiv = Convert.ToDouble(overallLength) / 10.00;
            allWhite();
            addFont();
            addDottedline();
            linkCoordinate(coordinate);
            addLimit((Int32)chainLength);
            addBFromBx(chainLength);
            return Img;
        }

        public Bitmap GetImg(Double Ax)
        {
            Bitmap ImgI = (Bitmap)Img.Clone(),
                ImgRet;
            addAFromAx(Ax);
            ImgRet = (Bitmap)Img.Clone();
            Img = (Bitmap)ImgI.Clone();
            g = Graphics.FromImage(Img);
            return ImgRet;
        }

        private void addBFromBx(Double Bx)
        {
            Int32 BxPiexl = (Int32)Math.Floor(Bx / (yDiv / 60.00));
            Int32 BoxX = BxPiexl - (32 + ((Int32)Bx).ToString().Length * 6) - 2;
            
            Pen p = new Pen(Color.Black, 1);
            g.DrawLine(p, BxPiexl, 50, BxPiexl, 275);

            SolidBrush b = new SolidBrush(Color.White);
            g.FillRectangle(b, BoxX, 50 , 32 + ((Int32)Bx).ToString().Length * 6, 15);


            Font font = new Font("宋体", 10);
            SolidBrush fontB = new SolidBrush(Color.Black);
            g.DrawString("B" + Bx.ToString("0.00"), font, fontB, new Point(BoxX, 50 + 2));

        }

        private void addAFromAx(Double Ax)
        {
            Int32 AxPiexl = (Int32)Math.Floor(Ax / (yDiv / 60.00));

            Pen p = new Pen(Color.Black, 1);
            g.DrawLine(p, AxPiexl, 25, AxPiexl, 275);

            SolidBrush b = new SolidBrush(Color.Black);
            g.FillRectangle(b, AxPiexl + 1, 25, 32 + ((Int32)Ax).ToString().Length * 6, 15);

            Font font = new Font("宋体", 10);
            SolidBrush fontB = new SolidBrush(Color.White);
            g.DrawString("A" + Ax.ToString("0.00"), font, fontB, new Point(AxPiexl, 25 + 2));
        }

        public void addFont()
        {
            Font font = new Font("宋体", 12, FontStyle.Bold);
            SolidBrush sbrush = new SolidBrush(Color.Black);
            g.DrawString("迹线图", font, sbrush, new PointF(2, 2));
            font = new Font("宋体", 10);
            g.DrawString(yDiv + " m/Div 4.000dB/Div", font, sbrush, new PointF(62, 6));
        }

        public void addLimit(Int32 chainLength)
        {
            Pen p = new Pen(Color.Black, 1);
            g.DrawLine(p, 0, 264, 0, 274);
            g.DrawLine(p, 1, 264, 1, 274);
            g.DrawLine(p, 2, 268, 7, 268);
            g.DrawLine(p, 2, 269, 7, 269);

            Int32 Bx = (Int32)Math.Floor(chainLength / (yDiv / 60));
            g.DrawLine(p, Bx, 264, Bx, 274);
            g.DrawLine(p, Bx - 1, 264, Bx - 1, 274);
            g.DrawLine(p, Bx - 2, 268, Bx - 7, 268);
            g.DrawLine(p, Bx - 2, 269, Bx - 7, 269);
        }

        public void linkCoordinate(Double[,] coordinate)
        {
            Pen p = new Pen(Color.Black,1);
            dot = new Int32[coordinate.GetLength(0), coordinate.GetLength(1)];
            for (Int32 i = 0; i < coordinate.GetLength(0); i++)
            {
                for (Int32 j = i; j < coordinate.GetLength(0); j++)
                    if (coordinate[i, 0] > coordinate[j, 0])
                    {
                        coordinate[i, 0] = coordinate[i, 0] + coordinate[j, 0];
                        coordinate[j, 0] = coordinate[i, 0] - coordinate[j, 0];
                        coordinate[i, 0] = coordinate[i, 0] - coordinate[j, 0];
                        coordinate[i, 1] = coordinate[i, 1] + coordinate[j, 1];
                        coordinate[j, 1] = coordinate[i, 1] - coordinate[j, 1];
                        coordinate[i, 1] = coordinate[i, 1] - coordinate[j, 1];
                    }
                dot[i, 0] = (Int32)Math.Floor(coordinate[i, 0] / (yDiv / 60.00));
                dot[i, 1] = (Int32)Math.Floor(275 - (coordinate[i, 1] / (4.00 / 25.00)));
                dot[i, 1] = dot[i, 1] < 25 ? 25 : dot[i, 1];
                if (i == 0) continue;
                if (dot[i - 1, 0] != dot[i, 0] || dot[i - 1, 1] != dot[i, 1]) g.DrawLine(p, dot[i - 1, 0], dot[i - 1, 1], dot[i, 0], dot[i, 1]);
                else Img.SetPixel(dot[i, 0], dot[i, 1], Color.Black);
            }
        }

        public void allWhite()
        {
            for (Int16 x = 0; x < 601; x++)
                for (Int16 y = 0; y < 276; y++)
                    Img.SetPixel(x, y, Color.White);
        }

        public void addDottedline()
        {
            for (Int16 i = 0; i < 11; i++)
                dottedLine(i * 60, 25, 0, 276, Img);
            for (Int16 i = 0; i < 11; i++)
                dottedLine(0, i * 25 + 25, 601, i * 25 + 25, Img);
        }

        public void dottedLine(Int32 Ax, Int32 Ay, Int32 Bx, Int32 By, Bitmap Img)
        {
            Int16 d = 0;  // 0 = Up | 1 = down | 2 = left | 3 = right
            Int32 Length = 1;
            if (Ax != Bx) { Length = Bx - Ax; d = 3; }
            if (Ay != By) { Length = By - Ay; d = 1; }
            if (Length < 0) { Length *= -1; d -= 1; }
            for (Int32 i = 0; i < Length; i++)
                if (d == 0) Img.SetPixel(Ax, Ay - i, i / 3 % 2 == 1 ? Color.White : Color.Black);
                else if (d == 1) Img.SetPixel(Ax, Ay + i, i / 3 % 2 == 1 ? Color.White : Color.Black);
                else if (d == 2) Img.SetPixel(Ax - i, Ay, i / 3 % 2 == 1 ? Color.White : Color.Black);
                else if (d == 3) Img.SetPixel(Ax + i, Ay, i / 3 % 2 == 1 ? Color.White : Color.Black);
        }
    }
}
