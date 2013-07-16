using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using Modbus.Data;
using Modbus.Message;
using Unme.Common;

namespace Modbus.IntegrationTests.CustomMessages
{
	public class CustomReadBufferResponse : IModbusMessage
	{
		private byte _functionCode;
		private byte _slaveAddress;
		private ushort _addressPoint;
        private ushort _packetNum;
		private ushort _transactionId;
		private byte[] _data;
        private string _strData;

		public byte[] Data
		{
			get
			{return  _data;}
		}

        public string StrData
        {
            get
            {
                return _strData ;
            }
        }

		public byte[] MessageFrame
		{
			get
			{
				List<byte> frame = new List<byte>();
				frame.Add(SlaveAddress);
				frame.AddRange(ProtocolDataUnit);

				return frame.ToArray();
			}
		}

		public byte[] ProtocolDataUnit
		{
			get
			{
				List<byte> pdu = new List<byte>();

				pdu.Add(_functionCode);
                pdu.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)AddressPoint)));
                pdu.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short)PacketNum)));
				pdu.AddRange(_data);

				return pdu.ToArray();
			}
		}

		public ushort TransactionId
		{
			get { return _transactionId; }
			set { _transactionId = value; }
		}

		public byte FunctionCode
		{
			get { return _functionCode; }
			set { _functionCode = value; }
		}

		public byte SlaveAddress
		{
			get { return _slaveAddress; }
			set { _slaveAddress = value; }
		}

        public ushort AddressPoint
        {
            get { return _addressPoint; }
            set { _addressPoint = value; }
        }

        public ushort PacketNum
        {
            get { return _packetNum; }
            set { _packetNum = value; }
        }

		public void Initialize(byte[] frame)
		{
			if (frame == null)
				throw new ArgumentNullException("frame");

			if (frame.Length < 3 || frame.Length < 3 + frame[2])
				throw new ArgumentException("Message frame does not contain enough bytes.", "frame");

			SlaveAddress = frame[0];
			FunctionCode = frame[1];
            AddressPoint = (ushort) IPAddress.NetworkToHostOrder(BitConverter.ToInt16(frame, 2));
            PacketNum = (ushort)IPAddress.NetworkToHostOrder(BitConverter.ToInt16(frame, 4));
			
            _data = frame.Slice(6,128).ToArray();

            _strData = System.Text.Encoding.ASCII.GetString(_data); 

		}
	}
}
