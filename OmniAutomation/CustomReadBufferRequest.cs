using System;
using System.Collections.Generic;
using System.Net;
using Modbus.Message;

namespace Modbus.IntegrationTests.CustomMessages
{
	public class CustomReadBufferRequest : IModbusMessage
	{
		private byte _functionCode;
		private byte _slaveAddress;
		private ushort _startAddress;
		private ushort _packetNum;
		private ushort _transactionId;

		public CustomReadBufferRequest(byte functionCode, byte slaveAddress, ushort startAddress, ushort packetNum)
		{
			_functionCode = functionCode;
			_slaveAddress = slaveAddress;
			_startAddress = startAddress;
			_packetNum = packetNum;
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

				pdu.Add(FunctionCode);
				pdu.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short) StartAddress)));
				pdu.AddRange(BitConverter.GetBytes(IPAddress.HostToNetworkOrder((short) PacketNum)));

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

		public ushort StartAddress
		{
			get { return _startAddress; }
			set { _startAddress = value; }
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

			if (frame.Length != 6)
				throw new ArgumentException("Invalid frame.", "frame");

			SlaveAddress = frame[0];
			FunctionCode = frame[1];
			StartAddress = (ushort) IPAddress.NetworkToHostOrder(BitConverter.ToInt16(frame, 2));
			PacketNum = (ushort) IPAddress.NetworkToHostOrder(BitConverter.ToInt16(frame, 4));
		}
	}
}