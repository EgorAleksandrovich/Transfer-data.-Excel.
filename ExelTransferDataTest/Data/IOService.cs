using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExelTransferDataTest.Data
{
    public interface IOService
    {
        string OpenFileDialog(string defaultPath);

        //Other similar untestable IO operations
        Stream OpenFile(string path);
    }
}
