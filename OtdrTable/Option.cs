using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace OtdrTable {
    class Option {
        public MakeGytsMode MarkGytsMode = MakeGytsMode.All;
        public Boolean Once = false;
        public List<String> args =  new List<String>();
        public Int32 WaveLength = 1310;
    }
}
