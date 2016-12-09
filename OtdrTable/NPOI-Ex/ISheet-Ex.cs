using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using NPOI.XSSF.UserModel;

namespace OtdrTable.NPOI_Ex {
    static class ISheet_Ex {

        static public IPicture AddPicture(this ISheet t,Bitmap pic, Int32 dx1,Int32 dy1, Int32 dx2, Int32 dy2, Int32 col1, Int32 row1, Int32 col2, Int32 row2) {
            Int32 PicIndex;
            using (MemoryStream stream = new MemoryStream()) {
                pic.Save(stream, ImageFormat.Jpeg);
                PicIndex = t.Workbook.AddPicture(stream.ToArray(), PictureType.JPEG);
            }

            IDrawing patriarch = t.CreateDrawingPatriarch();
            IClientAnchor anchor = t.Workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Dx1 = dx1;
            anchor.Dy1 = dy1;
            anchor.Dx2 = dx2;
            anchor.Dy2 = dy2;
            anchor.Col1 = col1;
            anchor.Row1 = row1;
            anchor.Col2 = col2;
            anchor.Row2 = row2;
            
            return patriarch.CreatePicture(anchor, PicIndex);
        }

        static public ICell SetCellValue(this ISheet t, Int32 row, Int32 col, String value, ICellStyle style = null) {
            IRow Row = t.GetRow(row) ?? t.CreateRow(row);
            ICell Cell = Row.GetCell(col) ?? Row.CreateCell(col);
            Cell.SetCellValue(value);
            if (Cell.CellStyle.Index == 0 && style != null) Cell.CellStyle = style;
            return Cell;
        }

        static public ICell SetCellValue(this ISheet t, Int32 row, Int32 col, Int32 value, ICellStyle style = null) {
            IRow Row = t.GetRow(row) ?? t.CreateRow(row);
            ICell Cell = Row.GetCell(col) ?? Row.CreateCell(col);
            Cell.SetCellValue(value);
            if (Cell.CellStyle.Index == 0 && style != null) Cell.CellStyle = style;
            return Cell;
        }

        static public void SetRowHeight(this ISheet t, Int32 row, Int16 height) {
            IRow Row = t.GetRow(row) ?? t.CreateRow(row);
            Row.Height = height;
        }

        static public void SetCellStyle(this ISheet t, Int32 row, Int32 col, ICellStyle style) {
            IRow Row = t.GetRow(row) ?? t.CreateRow(row);
            ICell Cell = Row.GetCell(col) ?? Row.CreateCell(col);
            Cell.CellStyle = style;
        }

        static public String GetCellValue(this ISheet t, Int32 row, Int32 col) {
            IRow Row = t.GetRow(row);
            if (Row == null) return null;
            ICell Cell = Row.GetCell(col);
            if (Cell == null) return null;
            return Cell.ToString();
        }

        static public void RemoveCol(this ISheet t, Int32 col) {
            for (Int32 row = t.FirstRowNum; row < t.LastRowNum + 1; row++) {
                IRow Row = t.GetRow(row);
                if (Row == null) continue;
                ICell Cell = Row.GetCell(col);
                if (Cell != null) Row.RemoveCell(Cell);
            }
        }

        static public Int32 MergeCells(this ISheet t, Int32 firstRow, Int32 lastRow, Int32 firstCol, Int32 lastCol) {
            CellRangeAddress Range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
            return t.AddMergedRegion(Range);
        }
    }
}
