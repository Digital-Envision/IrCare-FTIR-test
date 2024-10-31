using System;
using System.Collections.Generic;
using System.Text;

namespace MicroLabData
{
	public struct _progress
	{
		public int nStructSize;	// size of _progress structure
		public FTIR_STATE state;
		public int currentUnits;
		public int totalUnits;
		public int recentRejected;
		public int rejectReason;       // reason, or Good if the last scan was good
		public int numRejectsSame;     // num consecutive rejects with same rejectReason
	}

	public class MLProgress
	{
		public MLProgress() { m_progress.nStructSize = 28; } // System.Runtime.InteropServices.Marshal.SizeOf(_progress); }

		public _progress m_progress;
	}
}
