using System;
using System.Collections.Generic;
using System.Text;
using CommonShared;
using System.IO;
using System.Runtime.InteropServices;

namespace MicroLabData
{
    // for use in call to GetStatusEx
	public enum VERSIONSIZES { SIZEOFSERIALNO = 17 };

	[StructLayout(LayoutKind.Sequential)]
	public struct _instrumentMLDiag
	{
		public int nVersion;        // initial version is 100
		public int nEnergyStatus;   // Height of Center burst             
		public int nLaserStatus;    // (long in C/C++)             
		public int numTemps;        // (long in C/C++) number of fTemp items below
		public int nBatteryMinutes;
		public int nBatteryPct;
		public int nBatteryState;   // bits: 1=connected, 2=ac connected, 4= charging, 16=fullycharged
		public float fSourceCurrentStatus;
		public float fSourceVoltageStatus;
		public float fSpare;        // was Detector Status (actually reserved for Temp)

		public float fTempCPU;      // Cpu board
		public float fTempPower;    // Power board
		public float fTempIR;       // IR board
		public float fTempDetector; // Detector 

		// Version 102
		public Int32 nSystemStatus;
		public Int32 nShutdownReason;
	};

	public class MLDiag
	{
		public MLDiag() { m_diag.nVersion = 102; m_diag.numTemps = 4; }

		public _instrumentMLDiag m_diag;
		public _instrumentMLDiag InstrumentDiag
		{
			set { m_diag = value; }
			get { return m_diag; }
		}

		static public MLDiag LoadDiagFrom(BinaryReader fi, ref bool bRet)
		{
			bRet = false; // init as fail

			MLDiag diag = new MLDiag();

			if (fi != null)
			{
				try
				{
					// read version first from file...
					int fileVersion = fi.ReadInt32();

					// then read file Type and verify
					bool oktoLoad = true;
					if (fileVersion < 100)
						oktoLoad = false;

					if (oktoLoad)
					{
						// base case 
						int numT = fi.ReadInt32();

						diag.m_diag.nEnergyStatus = fi.ReadInt32();
						diag.m_diag.nLaserStatus = fi.ReadInt32();
						if (fileVersion >= 101)
						{
							diag.m_diag.nBatteryMinutes = fi.ReadInt32();
							diag.m_diag.nBatteryPct = fi.ReadInt32();
							diag.m_diag.nBatteryState = fi.ReadInt32();
						}
						else
						{
							diag.m_diag.nBatteryMinutes = (int)fi.ReadSingle();
							diag.m_diag.nBatteryPct = (int)fi.ReadSingle();
							diag.m_diag.nBatteryState = 0;
						}

						// CML 7-21-2007-  During the most recent formal Code Review, Jon Frattaroli noticed these variables were being loaded in a different order than in the SaveDiagTo method.
						diag.m_diag.fSpare = fi.ReadSingle();
						diag.m_diag.fSourceCurrentStatus = fi.ReadSingle();
						diag.m_diag.fSourceVoltageStatus = fi.ReadSingle();


						if (numT > 0)
							diag.m_diag.fTempCPU = fi.ReadSingle();
						if (numT > 1)
							diag.m_diag.fTempPower = fi.ReadSingle();
						if (numT > 2)
							diag.m_diag.fTempIR = fi.ReadSingle();
						if (numT > 3)
							diag.m_diag.fTempDetector = fi.ReadSingle();
					}
					if (fileVersion >= 102)
					{
						diag.m_diag.nSystemStatus = fi.ReadInt32();
						diag.m_diag.nShutdownReason = fi.ReadInt32();
					}
					bRet = true;
				}
				catch
				{
				}
			}
			return diag;
		}

		public bool SaveDiagTo(BinaryWriter fo)
		{
			bool ret = false;
			try
			{
				fo.Write(m_diag.nVersion);
				fo.Write(m_diag.numTemps);
				fo.Write(m_diag.nEnergyStatus);
				fo.Write(m_diag.nLaserStatus);
				fo.Write(m_diag.nBatteryMinutes);
				fo.Write(m_diag.nBatteryPct);
				fo.Write(m_diag.nBatteryState);
				fo.Write(m_diag.fSpare);
				fo.Write(m_diag.fSourceCurrentStatus);
				fo.Write(m_diag.fSourceVoltageStatus);
				fo.Write(m_diag.fTempCPU);
				fo.Write(m_diag.fTempPower);
				fo.Write(m_diag.fTempIR);
				fo.Write(m_diag.fTempDetector);
				fo.Write(m_diag.nSystemStatus);
				fo.Write(m_diag.nShutdownReason);

				ret = true;
			}
			catch
			{
			}
			return ret;
		}
	};
    

    // for use in call to GetVersionEx
	public struct _instrumentMLVersionEx
	{
		public int nVersion;        // initial version is 100
		public int fwRev;
		public int dllRev;
		public int spareRev;

		public int instrType;
		public int sampleTechType;	// negative ==> ATR
		public int atrType;         // 1, 3, 9 (for ATR type sampleTechs)
		public int spare;

		public double dLaserWN;
		public double dBasePathLength;   // for Transmission/gas cell sampleTechs
		public double dAdjustPathLength;   // for Transmission/gas cell sampleTechs

		public short serialNo01;
		public short serialNo02;
		public short serialNo03;
		public short serialNo04;
		public short serialNo05;
		public short serialNo06;
		public short serialNo07;
		public short serialNo08;
		public short serialNo09;
		public short serialNo10;
		public short serialNo11;
		public short serialNo12;
		public short serialNo13;
		public short serialNo14;
		public short serialNo15;
		public short serialNo16;
		public short serialNo17;

		public int nCpuBrdRev;
		public int nPwrBrdRev;
		public int nIrBrdRev;
		public int nLasBrdRev;
	};

    public class MLVersion
    {
		public MLVersion() { 
		m_version.nVersion = 102;
		m_version.serialNo01 = 0;
		m_version.serialNo02 = 0;
		m_version.serialNo03 = 0;
		m_version.serialNo04 = 0;
		m_version.serialNo05 = 0;
		m_version.serialNo06 = 0;
		m_version.serialNo07 = 0;
		m_version.serialNo08 = 0;
		m_version.serialNo09 = 0;
		m_version.serialNo10 = 0;
		m_version.serialNo11 = 0;
		m_version.serialNo12 = 0;
		m_version.serialNo13 = 0;
		m_version.serialNo14 = 0;
		m_version.serialNo15 = 0;
		m_version.serialNo16 = 0;
		m_version.serialNo17 = 0;
	} 

        public _instrumentMLVersionEx m_version;
        public _instrumentMLVersionEx InstrumentVersion
        {
            set { m_version = value; }
            get { return m_version; }
        }

		public string SerialNoString
		{
			get
			{
				char[] ser = new char[17];
				ser[16] = (char)'\0';
				ser[15] = (char)m_version.serialNo16;
				ser[14] = (char)m_version.serialNo15;
				ser[13] = (char)m_version.serialNo14;
				ser[12] = (char)m_version.serialNo13;
				ser[11] = (char)m_version.serialNo12;
				ser[10] = (char)m_version.serialNo11;
				ser[9] = (char)m_version.serialNo10;
				ser[8] = (char)m_version.serialNo09;
				ser[7] = (char)m_version.serialNo08;
				ser[6] = (char)m_version.serialNo07;
				ser[5] = (char)m_version.serialNo06;
				ser[4] = (char)m_version.serialNo05;
				ser[3] = (char)m_version.serialNo04;
				ser[2] = (char)m_version.serialNo03;
				ser[1] = (char)m_version.serialNo02;
				ser[0] = (char)m_version.serialNo01;

				string serstr = "";
				for (int ii = 0; ii < 15 && ser[ii] != '\0'; ii++)
					serstr += ser[ii];
				ser = null;

				return serstr; }
		}

        static public MLVersion LoadInstrumentVersionFrom(BinaryReader fi, ref bool bRet)
        {
            bRet = false; // init as fail

            MLVersion vers = new MLVersion();

            if (fi != null)
            {
                try
                {
                    // read version first from file...
                    int fileVersion = fi.ReadInt32();

                    // then read file Type and verify
                    bool oktoLoad = true;
                    if (fileVersion < 100)
                        oktoLoad = false;

                    if (oktoLoad)
                    {
                        vers.m_version.fwRev = fi.ReadInt32();
                        vers.m_version.dllRev = fi.ReadInt32();
                        vers.m_version.spareRev = fi.ReadInt32();
                        vers.m_version.instrType = fi.ReadInt32();
                        vers.m_version.sampleTechType = fi.ReadInt32();
                        vers.m_version.atrType = fi.ReadInt32();
                        vers.m_version.spare = fi.ReadInt32();
                        vers.m_version.dLaserWN = fi.ReadDouble();
                        vers.m_version.dBasePathLength = fi.ReadDouble();
                        vers.m_version.dAdjustPathLength = fi.ReadDouble();

						if (fileVersion > 100)
						{
	                        int siz = fi.ReadInt32();
							vers.m_version.serialNo01 = (short)fi.ReadChar();
							vers.m_version.serialNo02 = (short)fi.ReadChar();
							vers.m_version.serialNo03 = (short)fi.ReadChar();
							vers.m_version.serialNo04 = (short)fi.ReadChar();
							vers.m_version.serialNo05 = (short)fi.ReadChar();
							vers.m_version.serialNo06 = (short)fi.ReadChar();
							vers.m_version.serialNo07 = (short)fi.ReadChar();
							vers.m_version.serialNo08 = (short)fi.ReadChar();
							vers.m_version.serialNo09 = (short)fi.ReadChar();
							vers.m_version.serialNo10 = (short)fi.ReadChar();
							vers.m_version.serialNo11 = (short)fi.ReadChar();
							vers.m_version.serialNo12 = (short)fi.ReadChar();
							vers.m_version.serialNo13 = (short)fi.ReadChar();
							vers.m_version.serialNo14 = (short)fi.ReadChar();
							vers.m_version.serialNo15 = (short)fi.ReadChar();
							vers.m_version.serialNo16 = (short)fi.ReadChar();
							vers.m_version.serialNo17 = (short)fi.ReadChar();
							for (int jj = (int)VERSIONSIZES.SIZEOFSERIALNO; jj < siz; jj++)
								fi.ReadChar();	// bit bucket if no place for it to go
						}
                    }
                    bRet = true;
                }
                catch
                {
                }
            }
            return vers;
        }

        public bool SaveInstrumentVersionTo(BinaryWriter fo)
        {
            bool ret = false;
            try
            {
                fo.Write(m_version.nVersion);
                fo.Write(m_version.fwRev);
                fo.Write(m_version.dllRev);
                fo.Write(m_version.spareRev);
                fo.Write(m_version.instrType);
                fo.Write(m_version.sampleTechType);
                fo.Write(m_version.atrType);
                fo.Write(m_version.spare);
                fo.Write(m_version.dLaserWN);
                fo.Write(m_version.dBasePathLength);
                fo.Write(m_version.dAdjustPathLength);

				int siz = (int)VERSIONSIZES.SIZEOFSERIALNO;
				fo.Write(siz);

				fo.Write((char)m_version.serialNo01);
				fo.Write((char)m_version.serialNo02);
				fo.Write((char)m_version.serialNo03);
				fo.Write((char)m_version.serialNo04);
				fo.Write((char)m_version.serialNo05);
				fo.Write((char)m_version.serialNo06);
				fo.Write((char)m_version.serialNo07);
				fo.Write((char)m_version.serialNo08);
				fo.Write((char)m_version.serialNo09);
				fo.Write((char)m_version.serialNo10);
				fo.Write((char)m_version.serialNo11);
				fo.Write((char)m_version.serialNo12);
				fo.Write((char)m_version.serialNo13);
				fo.Write((char)m_version.serialNo14);
				fo.Write((char)m_version.serialNo15);
				fo.Write((char)m_version.serialNo16);
				fo.Write((char)m_version.serialNo17);

                ret = true;
            }
            catch
            {
            }
            return ret;
        }
    };

  


}
