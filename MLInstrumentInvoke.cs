using CommonShared;
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using MicroLabData;

namespace MicroLabPC
{
    public class InstrumentFailedException : System.Exception
    {
        public InstrumentFailedException(int code) { nInitCode = code; }
        public int nInitCode;
    }

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
    };


    // for use in call to GetVersionEx
    public struct _instrumentMLVersionEx
    {
        public int nVersion;        // initial version is 100, current 101
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
    };

	public class MLProgress
	{
		public MLProgress() { m_progress.nStructSize = 28; } // System.Runtime.InteropServices.Marshal.SizeOf(_progress); }

		public _progress m_progress;
	}

    /// <summary>
    /// P/Invoke wrapper for the instrument dll's interface
    /// </summary>
	public class MLInstrumentInvoke
    {
		#region public methods
		#endregion

		#region properties
		#endregion

		#region constants

		public const double MLInstrumentRangeMaxFrom = 4000.0;
		public const double MLInstrumentRangeMaxTo = 650.0;
		public const int MLCleanRes = 4;

		#endregion

		protected static bool m_bInitialized = false;
		public static bool Initialized
		{
			get { return m_bInitialized; }
			set { m_bInitialized = value; }
		}
        protected static bool m_binitializing = false;
        public static bool initializing
        {
            get { return m_binitializing; }
            set { m_binitializing = value; }
        }

		#region protected/private methods

		/// <summary>
		/// Determines if instrument has been successfully initialized and if not, attempts initialization.
		/// </summary>
		/// Each instrument interface function should make a call to this method to verify that the instrument
		/// has been initialized.
		/// <returns>0 if success, non-zero error code otherwise</returns>
		protected static int InitializationCheck()
		{
			int nInitCode = 0; // init as success
			bool bExiting = false;

			while (!Initialized && !bExiting) // not yet initialized
			{
                if (!initializing)
                {
		            if (!Initialized)
                    {
                        initializing = true;
				        nInitCode = MLInstrumentInvoke.Init();
				        //nInitCode = -200; // diagnostic (injected error condition)
						if (nInitCode == 0)
						{
							Initialized = true; // successful initialization - set the property to true
							initializing = false;
                        }
				        else
				        {

/*					        string sTitle = StringHandler.GetString("InstrumentInitErrorTitle");
							string sMsg1 = StringHandler.GetString("InstrumentInitErrorText_Part1");
							string sMsg2 = StringHandler.GetString("InstrumentInitErrorText_Part2");
					        
					        DialogResult result = MessageBox.Show(sMsg1 + nInitCode.ToString() + sMsg2, sTitle, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
        
					        //if (result == DialogResult.Cancel)
					        //{
					        //	bExiting = Program.g_MainForm.ExitApp(); // allow user opportunity to exit the app
					        //}
					        // else we will try calling Init again
*/
                            bExiting = true;
//					        Application.Exit();
                            initializing = false;

                            InstrumentFailedException anError = new InstrumentFailedException(nInitCode);
                            throw anError;
                        }
                    }
				}
			}
			return nInitCode;
		}

		#endregion

		#region event handlers
		#endregion

		#region instrument interface calls

		public static int SetComputeParams(PHASEPOINTS ppoints, PHASETYPE ptype, APODTYPE papod, APODTYPE iapod, ZFFTYPE zff, OFFSETCORRECTTYPE offset)
		{
			InitializationCheck();
			return FTIRInst_SetComputeParams(ppoints, ptype, papod, iapod, zff, offset);
		}
        [DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_SetComputeParams(PHASEPOINTS ppoints, PHASETYPE ptype, APODTYPE papod, APODTYPE iapod, ZFFTYPE zff, OFFSETCORRECTTYPE offset);

		public static int StartSingleBeam(int numScans, double from, double to, int res, bool bAutoSetBkg, bool bAutoSetClean)
		{
			InitializationCheck();
			int ret = 0;
			try
			{
#if DEBUG
				//string sss = "Parameters: numScans=" + numScans.ToString() +
				//							" from=" + from.ToString() +
				//							" to=" + to.ToString() +
				//							" res=" + res.ToString() +
				//							" SetBKG=" + bAutoSetBkg.ToString() +
				//							" SetCLN=" + bAutoSetClean.ToString() +
				//							"\n";
				//DebugOut.Write("About to StartSingleBeam\n");
				//DebugOut.Write(sss);
#endif
				ret = FTIRInst_dptrStartSingleBeam(numScans, ref from, ref to, res, (bAutoSetBkg != false) ? 1 : 0, (bAutoSetClean != false) ? 1 : 0);
			}
			catch (Exception ce)
			{
				MessageBox.Show(ce.Message, "Exception on StartSingleBeam");
#if DEBUG
				//DebugOut.Write("Exception Caught\n");
#endif
			}
			return ret;

		}
        [DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrStartSingleBeam(int numScans, ref double from, ref double to, int res, int bAutoSetBkg, int bAutoSetClean);

		public static FTIR_STATE CheckProgress(ref int currentUnits, ref int totalUnits)
		{
			// Since this method is often called many times in succession, we don't do the
			// initialization check in this function so that the user is not barraged with
			// a mass of message boxes - if the instrument has not yet been initialized
			// when this is called, the user has no doubt already been notified.
			//InitializationCheck();
			return FTIRInst_CheckProgress(ref currentUnits, ref totalUnits);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern FTIR_STATE FTIRInst_CheckProgress(ref int currentUnits, ref int totalUnits);

		public static FTIR_STATE CheckProgressEx(ref int currentUnits, ref int totalUnits,ref int rejectedScans)
		{
			try
			{
				return  FTIRInst_CheckProgressEx(ref currentUnits, ref totalUnits,ref rejectedScans);
			}
			catch
			{
				rejectedScans = 0;
				return FTIRInst_CheckProgress(ref currentUnits, ref totalUnits);
			}
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern FTIR_STATE FTIRInst_CheckProgressEx(ref int currentUnits, ref int totalUnits,ref int rejectedScans);

		public static int CheckProgressStruct(ref MLProgress pProgress)
		{
			InitializationCheck();
			try
			{
				return FTIRInst_CheckProgressStruct(ref pProgress.m_progress);
			}
			catch
			{
				pProgress.m_progress.state = FTIRInst_CheckProgressEx(ref pProgress.m_progress.currentUnits, ref pProgress.m_progress.totalUnits, ref pProgress.m_progress.recentRejected);
				pProgress.m_progress.numRejectsSame = 0;
				pProgress.m_progress.rejectReason = 0;
				return 0;
			}
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_CheckProgressStruct(ref _progress pProgress);
		
		public static int GetSingleBeamSize()
        {
			double refF = 0;
            double refT = 0;
            int refR = 0;
			InitializationCheck();
			return FTIRInst_dptrGetSingleBeam(null, 0, ref refF, ref refT, ref refR);
        }
        public static int GetSingleBeam(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes)
		{
			InitializationCheck();
			return FTIRInst_dptrGetSingleBeam(array, size, ref actualFrom, ref actualTo, ref actualRes);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrGetSingleBeam(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes);

		// JSA 12/06/07 new functions to retrieve the instrument's stored background spectrum (or its size) - if one exists
		public static int GetBackgroundSize()
		{
			double refF = 0;
			double refT = 0;
			int refR = 0;
			InitializationCheck();
			return GetBackground(null, 0, ref refF, ref refT, ref refR);
		}
		public static int GetBackground(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes)
		{
			InitializationCheck();
			try
			{
				return FTIRInst_dptrGetBackground(array, size, ref actualFrom, ref actualTo, ref actualRes);
			}
			catch
			{
				return FTIRInst_dptrGetSingleBeam(array, size, ref actualFrom, ref actualTo, ref actualRes);
			}

		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrGetBackground(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes);

		// JSA 12/06/07 new functions to retrieve the instrument's stored clean spectrum (or its size) - if one exists
		public static int GetCleanSize()
		{
			double refF = 0;
			double refT = 0;
			int refR = 0;
			return GetClean(null, 0, ref refF, ref refT, ref refR);
		}
		public static int GetClean(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes)
		{
			InitializationCheck();
			try
			{
				return FTIRInst_dptrGetClean(array, size, ref actualFrom, ref actualTo, ref actualRes);
			}
			catch
			{
				return FTIRInst_dptrGetSingleBeam(array, size, ref actualFrom, ref actualTo, ref actualRes);
			}
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrGetClean(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes);

		public static int GetRatioSpectrum(double[] bkgarray, double[] smparray, double[] outarray, int size, DATAYTYPE ytype)
		{
			InitializationCheck();
			try
			{
				return FTIRInst_dptrGetRatioSpectrum(bkgarray, smparray, outarray, size, ytype);
			}
			catch
			{
				MessageBox.Show("Older FTIRInst.DLL interface found.  Need to update instrument driver.");
				return -1;
			}
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrGetRatioSpectrum(double[] bkgarray, double[] smparray, double[] outarray, int size, DATAYTYPE ytype);

		// JSA 12/06/07 FTIRInst_dptrSetBackground() is not yet fully implemented in the firmware - do not call it!
		// Currently, the only way to set the background is via FTIRInst_dptrStartSingleBeam() with the bAutoSetBkg set to TRUE
		/*
		public static int SetBackground(double[] array, int size, double from, double to, int res)
		{
			InitializationCheck();
			return FTIRInst_dptrSetBackground(array, size, ref from, ref to, res);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrSetBackground(double[] array, int size, ref double from, ref double to, int res);
		*/

		// JSA 12/06/07 FTIRInst_dptrSetClean() is not yet fully implemented in the firmware - do not call it!
		// Currently, the only way to set the background is via FTIRInst_dptrStartSingleBeam() with the bAutoSetClean set to TRUE
		/*
		public static int SetClean(double[] array, int size, double from, double to, int res)
		{
			InitializationCheck();
			return FTIRInst_dptrSetClean(array, size, ref from, ref to, res);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrSetClean(double[] array, int size, ref double from, ref double to, int res);
		*/

		public static int StartSpectrum(int numScans, double from, double to, int res, DATAXTYPE xtype, DATAYTYPE ytype, bool bAutoSetUnknown)
		{
			InitializationCheck();
			return FTIRInst_dptrStartSpectrum(numScans, ref from, ref to, res, xtype, ytype, (bAutoSetUnknown != false) ? 1 : 0);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrStartSpectrum(int numScans, ref double from, ref double to, int res, DATAXTYPE xtype, DATAYTYPE ytype, int bAutoSetUnknown);

        public static int GetSpectrumSize()
        {
			double refF = 0;
            double refT = 0;
            int refR = 0;
			InitializationCheck();
			return FTIRInst_dptrGetSpectrum(null, 0, ref refF, ref refT, ref refR);
        }
        public static int GetSpectrum(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes)
		{
			InitializationCheck();
			return FTIRInst_dptrGetSpectrum(array, size, ref actualFrom, ref actualTo, ref actualRes);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrGetSpectrum(double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes);

		public static int SetUnknown(double[] array, int size, double from, double to, int res, DATAXTYPE xtype, DATAYTYPE ytype, bool bIsATR)
		{
			InitializationCheck();
			return FTIRInst_dptrSetUnknown(array, size, ref from, ref to, res, xtype, ytype, (bIsATR != false) ? 1 : 0);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrSetUnknown(double[] array, int size, ref double from, ref double to, int res, DATAXTYPE xtype, DATAYTYPE ytype, int bIsATR);

		public static int KillCollection()
		{
			InitializationCheck();
			return FTIRInst_KillCollection();
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_KillCollection();

		public static int SoftReset()
		{
			InitializationCheck();
			return FTIRInst_SoftReset();
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_SoftReset();

		public static int GetVersion(ref int fwRev, ref int dllRev, ref int spareRev)
		{
			InitializationCheck();
			return FTIRInst_GetVersion(ref fwRev, ref dllRev, ref spareRev);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_GetVersion(ref int fwRev, ref int dllRev, ref int spareRev);

		public static int GetStatus(ref int nEnergyStatus, ref float fBatteryStatus, ref float fSourceCurrentStatus, ref float fSourceVoltageStatus, ref int nLaserStatus, ref float fDetectorStatus)
		{
			InitializationCheck();
			return FTIRInst_GetStatus(ref nEnergyStatus, ref fBatteryStatus, ref fSourceCurrentStatus, ref fSourceVoltageStatus, ref nLaserStatus, ref fDetectorStatus);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_GetStatus(ref int nEnergyStatus, ref float fBatteryStatus, ref float fSourceCurrentStatus, ref float fSourceVoltageStatus, ref int nLaserStatus, ref float fDetectorStatus);

        public static int GetVersionEx(MLVersion vInfo)
        {
            InitializationCheck();
            bool bCalled = false;
            int ret = 0;
            try
            {
                ExFunctionInvoke ex = new ExFunctionInvoke();
                ret = ex.GetVersionEx(ref vInfo.m_version);
                bCalled = true;
            }
            catch
            {
            }
            if (!bCalled)
            {
                ret = FTIRInst_GetVersion(ref vInfo.m_version.fwRev, ref vInfo.m_version.dllRev, ref vInfo.m_version.spareRev);
                vInfo.m_version.dLaserWN = 7633.0;
                vInfo.m_version.instrType = 0;
                vInfo.m_version.sampleTechType = (int)SAMPLINGTECHNOLOGYTYPE.ST_TRANSMISSIONCELL;
                vInfo.m_version.atrType = 1;
                vInfo.m_version.dBasePathLength = 0.1;
                vInfo.m_version.dAdjustPathLength = 0;
            }
            return ret;
        }

        public static int GetStatusEx(ref MLDiag dStatus)
        {
            InitializationCheck();
            bool bCalled = false;
            int ret = 0;
            try
            {
                ExFunctionInvoke ex = new ExFunctionInvoke();
                ret = ex.GetStatusEx(ref dStatus.m_diag);
                bCalled = true;
            }
            catch
            {
            }
            if (!bCalled)
            {
                float fBatStat = 0;
                ret = FTIRInst_GetStatus(ref dStatus.m_diag.nEnergyStatus, ref fBatStat, ref dStatus.m_diag.fSourceCurrentStatus, ref dStatus.m_diag.fSourceVoltageStatus, ref dStatus.m_diag.nLaserStatus, ref dStatus.m_diag.fTempDetector);
                dStatus.m_diag.nBatteryMinutes = (int)fBatStat;
                dStatus.m_diag.nBatteryPct = 0;
                dStatus.m_diag.nBatteryState = 0;
                dStatus.m_diag.fTempCPU = 0;
                dStatus.m_diag.fTempIR = 0;
                dStatus.m_diag.fTempPower = 0;
            }
            return ret;
        }
        
        public static int Init()
		{
			// do not put this here - this is one of the few instrument interface functions
			// where you should not do this (would be recursive and would blow the stack)
			//InitializationCheck();
			return FTIRInst_Init();
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_Init();

        public static int Deinit()
		{
			InitializationCheck();
			return FTIRInst_Deinit();
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_Deinit();

        public static int GetLiveSpectrum(double from, double to, int res, DATAXTYPE xtype, DATAYTYPE ytype, double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes)
        {
            InitializationCheck();
            return FTIRInst_dptrGetLiveSpectrum(ref from, ref to, res, xtype, ytype, array, size, ref actualFrom, ref actualTo, ref actualRes);
        }
        [DllImport("FTIRInst.dll", SetLastError = true)]
        private static extern int FTIRInst_dptrGetLiveSpectrum(ref double from, ref double to, int res, DATAXTYPE xtype, DATAYTYPE ytype, double[] array, int size, ref double actualFrom, ref double actualTo, ref int actualRes);

		public static int SetLaserWN(double newLaser)
		{
			InitializationCheck();
			float fltLaser = (float)newLaser;
			return FTIRInst_SetLaserWaveNumber(ref fltLaser);
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_SetLaserWaveNumber(ref float newLaser);

		public static double GetLaserWN()
		{
			InitializationCheck();
			float fltLaser = 0;
			FTIRInst_GetLaserWaveNumber(ref fltLaser);
			return (double)fltLaser;
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_GetLaserWaveNumber(ref float curLaser);

		public static int SetPathlength(double newPathlength, MLVersion _vInfo)
		{
			InitializationCheck();
			float fltPathlength = (float)newPathlength;
			int ret = FTIRInst_SetPathlen(ref fltPathlength);

			// Now force re-getting the vInfo
			GetVersionEx(_vInfo);
			return ret;
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_SetPathlen(ref float newPathlength);

		public static double GetPathlength(MLVersion vers)
		{
			InitializationCheck();
			double fltPathlength = FTIRInst_GetPathlen(ref vers.m_version);
			return (double)fltPathlength;
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern double FTIRInst_GetPathlen(ref MicroLabData._instrumentMLVersionEx _vInfo);

		public static int StartCoaddedIGram(int numScans, int nRes, int nPhasePts)
		{
			InitializationCheck();
			int ret = 0;
			try
			{
				ret = FTIRInst_StartCoaddedIgram(numScans, nRes, nPhasePts);
			}
			catch (Exception ce)
			{
				MessageBox.Show(ce.Message, "Exception on StartCoaddedIgram");
			}
			return ret;
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_StartCoaddedIgram(int numScans, int nRes, int nPhasePts);
		public static int GetCoaddedIgram(double[] pArray, int nArraySize)
		{
			InitializationCheck();
			int ret = 0;
			try
			{
				ret = FTIRInst_dptrGetCoaddedIgram(pArray, nArraySize);
			}
			catch (Exception ce)
			{
				MessageBox.Show(ce.Message, "Exception on GetCoaddedIgram");
			}
			return ret;
		}
		[DllImport("FTIRInst.dll", SetLastError = true)]
		private static extern int FTIRInst_dptrGetCoaddedIgram(double[] pArray,  int nArraySize);


		#endregion

		#region member variables
		#endregion

		#region enums and nested classes
		#endregion
	}

    public class ExFunctionInvoke
    {
        public int GetVersionEx(ref MicroLabData._instrumentMLVersionEx _vInfo)
        {
            return FTIRInst_GetVersionEx(ref _vInfo);
        }
        [DllImport("FTIRInst.dll", SetLastError = true)]
        private static extern int FTIRInst_GetVersionEx(ref MicroLabData._instrumentMLVersionEx _vInfo);

        public int GetStatusEx(ref MicroLabData._instrumentMLDiag _dStatus)
        {
            return FTIRInst_GetStatusEx(ref _dStatus);
        }
        [DllImport("FTIRInst.dll", SetLastError = true)]
        private static extern int FTIRInst_GetStatusEx(ref MicroLabData._instrumentMLDiag _dStatus);
    }
}
