using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NPOI.SS.UserModel;

namespace OtdrTable.NPOI_Ex {
    static class IWorkbook_Ex {
        static public void Save(this IWorkbook t, String filepath) {
            using (FileStream fs = File.Open(filepath, FileMode.Create)) // 打開文件
                t.Write(fs);
        }

        static public void Read(this IWorkbook t, String filepath) {
            using (FileStream fs = File.OpenRead(filepath))
                t = WorkbookFactory.Create(fs);
        }
    }
}
