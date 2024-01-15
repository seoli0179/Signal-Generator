using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Serial2
{
    class SerialControl
    {
        public String[] serchPortList()
        {
            return System.IO.Ports.SerialPort.GetPortNames();
        }

    }
}
