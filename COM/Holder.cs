using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace AudioSwitch.COM
{
    internal class COMHolder<T> : IDisposable
    {
        private T obj;
        public T Value { get { return obj; } }

        public COMHolder(T o) {
            this.obj = o;
        }

        public void Dispose()
        {
            Marshal.ReleaseComObject(this.obj);
        }
    }
}
