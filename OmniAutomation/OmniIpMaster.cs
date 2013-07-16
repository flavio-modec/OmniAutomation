using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net.Sockets;
using Modbus;
using Modbus.Device;
using Modbus.Message;
using Modbus.Utility;
using Modbus.Data;
using Modbus.IntegrationTests.CustomMessages;
using Modbus.IO;


namespace OmniAutomation
{
    class OmniIpMaster
    {
        ModbusIpMaster master;

        public OmniIpMaster(ModbusIpMaster _master)
        {
            master = _master;
        }

        public ushort[] ReadInt16(ushort address, ushort points)
        {
            return master.ReadHoldingRegisters(1, address, points);
        }

        public int[] ReadInt32(ushort address, ushort points)
        {
            int[] longval = new int[points];
            CustomReadHoldingRegistersRequest readLong = new CustomReadHoldingRegistersRequest(3, 1, address, points);
            CustomReadHoldingRegistersResponse response = master.ExecuteCustomMessage<CustomReadHoldingRegistersResponse>(readLong);
            for (ushort i = 0; i < points; i++)
            {
                byte[] longOr = new byte[] { response.Data[3 + 4 * i], response.Data[2 + 4 * i], response.Data[1 + 2 * i], response.Data[0 + 2 * i] };
                longval[i] = BitConverter.ToInt32(longOr, 0);
            }
            return longval;
        }

        public float[] ReadFloat (ushort address, ushort points)
        {
            float[] floatval = new float[points];
            CustomReadHoldingRegistersRequest readFloat = new CustomReadHoldingRegistersRequest(3, 1, address, points);
            CustomReadHoldingRegistersResponse response = master.ExecuteCustomMessage<CustomReadHoldingRegistersResponse>(readFloat);
            for (ushort i = 0; i < points; i++)
            {
                byte[] floatOr = new byte[] { response.Data[3 + 4 * i], response.Data[2 + 4 * i], response.Data[1 + 4 * i], response.Data[0 + 4 * i] };
                floatval[i] = BitConverter.ToSingle(floatOr, 0);
            }
            return floatval;
        }

        public string ReadString(ushort address, ushort points)
        {
            CustomReadHoldingRegistersRequest readAscii = new CustomReadHoldingRegistersRequest(3, 1, address, points);
            CustomReadHoldingRegistersResponse response = master.ExecuteCustomMessage<CustomReadHoldingRegistersResponse>(readAscii);
            return System.Text.Encoding.ASCII.GetString(response.Data);
        }
    }
}
