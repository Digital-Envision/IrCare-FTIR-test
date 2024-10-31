using System;

namespace MicroLabData
{
    // Signature identification stored to ML files to make sure we are reading the correct file
	// JSA 12/06/07 added new MLCLEAN file type
	public enum MLFILETYPE { MLCONFIG = 0x6372a001, MLMETHOD = 0x2637a002, MLRESULTS = 0x2736a003, MLSYSTEMCHECK = 0x5a098876, MLCLEAN = 0x2112a007}

	// JSA 03-29-07 moved to SharedCommonEnums.cs
	// Common 'public' enums needed throughout namespace
    //public enum DATAXTYPE { XT_ARB = 0, XT_WN = 1, XT_uM = 2, XT_nM = 3, XT_Seconds = 4, XT_Minutes = 5, XT_MassCharge = 9, XT_RAMSHFT = 13, XT_Points = 22, XT_Hours = 30, XT_AMU = 50, XT_Custom = 51 };
    //public enum DATAYTYPE { YT_ARB = 0, YT_IGRAM = 1, YT_Abs = 2, YT_Percent = 11, YT_Intensity = 12, YT_RelAbundance = 13, YT_Trans = 128, YT_Refl = 129, YT_Custom = 51, YT_Abundance = 52, };

	// Library Spectrum Comparison types
    public enum QUALITYTYPE { QT_EUCLIDEAN = 1, QT_ABSVALUE = 2, QT_DERIV_ABS = 3, QT_LEASTSQ = 4, QT_DERIV_LS = 5, QT_CORR = 6, QT_DERIV_CORR = 7, QT_SIM = 8, QT_DERIV_SIM = 9, QT_DOT = 10 /* fast coarse check */  };
    public enum PHASETYPE { PT_MERTZ = 1, PT_FORMAN = 2, PT_FORMANRES = 3 };
	public enum APODTYPE
    {
        APOD_NONE = 0, APOD_BOXCAR = APOD_NONE, APOD_TRIANGULAR = 1, APOD_WEAKNORTONBEER = 2, APOD_MEDIUMNORTONBEER = 3,
        APOD_STRONGNORTONBEER = 4, APOD_HAPPGENZEL = 5, APOD_BESSEL = 6, APOD_COSINE = 7,
        APOD_HANNING = APOD_COSINE,
    };

    // JAL Added Trend type method
    // JSA Added Data Collection Only and Discriminate method types
	// JSA 03-07-07 Added Components method type
	public enum METHODTYPE { MTYPE_QUAL = 1, MTYPE_QUANT = 2, MTYPE_BOTH = 3, MTYPE_TREND = 4, MTYPE_DATACOLL = 5, MTYPE_DISCRIM = 6, MTYPE_COMP = 7 };
    public enum QUANTTYPE { QUANT_PLS = 1, QUANT_SAS = 2 };

	public enum FTIR_STATE { FTIR_Init = 0, FTIR_Collecting = 1, FTIR_DataReady = 2, FTIR_Aborting = 3, FTIR_Error = 4, FTIR_Transmitting = 5 };
	public enum PHASEPOINTS { PP_128 = 128, PP_256 = 256, PP_512 = 512, PP_1024 = 1024 };
	public enum OFFSETCORRECTTYPE { OT_NONE = 0, OT_ALL = 1, OT_ENDS = 2 };
	public enum ZFFTYPE { ZFF_NONE = 0, ZFF_2 = 1, ZFF_4 = 2, ZFF_8 = 3, ZFF_16 = 4 };

    public enum CSIDL
    {
        CSIDL_FLAG_CREATE = 0x8000, CSIDL_APPDATA = 0x001a, CSIDL_COMMON_APPDATA = 0x0023,
        CSIDL_COMMON_DOCUMENTS = 0x002e, CSIDL_COMMON_PROGRAMS = 0x0017, CSIDL_LOCAL_APPDATA = 0x001c,
        CSIDL_MYDOCUMENTS = 0x000c, CSIDL_PERSONAL = 0x0005, CSIDL_PROFILE = 0x0028,
        CSIDL_PROGRAM_FILES = 0x0026, CSIDL_PROGRAM_FILES_COMMON = 0x002b,
        CSIDL_SYSTEM = 0x0025, CSIDL_WINDOWS = 0x0024,
    };

    // DLSAlg curve fitting order
    public enum POLY_ORDER { ORDER_LINE = 2, ORDER_2nd = 3, ORDER_3rd = 4, ORDER_4th = 5, ORDER_5th = 6, ORDER_6th = 7 };

    // LIB and Spectrum information types
    public enum LIBBYTES { LB_1BYTE = 0, LB_2BYTES = 1, LB_4BYTES = 2, LB_DOUBLE = 3 };
    public enum LIBISATR { ATR_UNDEFINED = -1, ATR_FALSE = 0, ATR_TRUE = 1 };
    public enum LIBDATATYPE { LIBDT_FTIR = 0, LIBDT_RAMAN = 1, LIBDT_MASS_SPEC = 2, LIBDT_MAXTYPE = 15 } ;
    public enum LIBENCTYPE { LIBENC_NONE = 0, LIBENC_SD1 = 1, /* more types can be added here */ };

    // LIB SEARCH
    public enum GETLIBENTRYITEMS { GLE_DATA = 0x0001, GLE_TEXT = 0x0002, GLE_UCID = 0x0004, GLE_STRUCT = 0x0008 /* future - not yet implemented!*/ };
    public enum INTERPTYPE { IT_UNKTOLIBSpline = 1, IT_LIBTOUNKSpline = 2, IT_UNKTOLIBLinear = 3, IT_LIBTOUNKLinear = 4 };
    public enum SEARCHSTATECODE { SSC_INIT = 0, SSC_RUNNING = 1, SSC_COMPLETED = 2, SSC_ABORTED = 3, };

	// JSA added residual mode enum
	public enum RESIDUALMODETYPE { RM_NONE = 0, RM_TOPHIT = 1, RM_DEFSPEC = 2 };

	// JSA 03-07-07 added sampling technology type enum
    public enum SAMPLINGTECHNOLOGYTYPE 
	{ 
		ST_NONE = 0, 
		ST_ATRSINGLE = -1, 
		ST_ATRTRIPLE = -3, 
		ST_ATRNINEBOUNCE = -9, 
		ST_TRANSMISSIONCELL = 1, 
		ST_GASCELL = 2, 
		ST_REFLECTANCE = 3,
                ST_LASTVALID = 3,
	};

	public enum MLPASSFAILSTATE
	{
		eInvalidPassFailState      = -1,
		eFirstPassFailState        =  1,
		eUndeterminedPassFailState =  1, // N/A
		eGreenPassFailState        =  2, // PASS
		eYellowPassFailState       =  3, // PASS (marginal)
		eRedPassFailState          =  4, // FAIL
		eLastPassFailState         =  4,
	};

	public enum MLPANELTYPES
	{
		eMLInvalidPanel                 = -1,
		eMLLoginPanel                   =  1,
		eMLMainMenuPanel                =  2,
		eMLCleanCrystalPanel            =  3,
		eMLCheckingCrystalProgressPanel =  4,
		eMLCollectBkgdProgressPanel     =  5,
		eMLUserPrepMsgSamplePanel       =  6,
		eMLPrepareSamplePanel           =  7,
		eMLSamplingProgressPanel        =  8,
		eMLResultsFileListPanel         =  9,
		eMLResultsPanel                 = 10,
		eMLDetailsPanel                 = 11,
		eMLDataHandlingPanel            = 12,
		eMLMethodsFileListPanel         = 13,
		eMLMethodEditPanel              = 14,
		eMLTroubleshootingPanel         = 15,
		eMLSystemCheckPanel             = 16,
		eMLTextFindPanel                = 17,
		eMLAdvancedFeaturesPanel        = 18,
		eMLCustomFieldsPanel            = 19,
		eMLDetailsSampleOnlyPanel       = 20,
		eMLLibraryMaintenancePanel		= 21,
		eMLUserManagementPanel			= 22,
		eMLSystemCheckResultsPanel      = 23,
		eMLAboutBoxPanel                = 24,
		eMLReportConfigPanel			= 25,
		eMLGainAdjustPanel				= 26,
	};

	public enum MLSTATUSTYPES
	{
		eMLInvalidStatus        = -1,
		eMLStandbyStatus        =  1,
		eMLReadyStatus          =  2,
		eMLCheckingStatus       =  3,
		eMLCollectingBkgdStatus =  4,
		eMLSamplingStatus       =  5,
		eMLResultsStatus        =  6,
		eMLSavingStatus         =  7,
		eMLPassedStatus         =  8,
		eMLFailedStatus         =  9,
	};

	public enum MLOPERATIONTYPE
	{
		eMLInvalidOperation     = -1,
		eMLNone                 =  1,
		eMLRunCleanOperation    =  2,
		eMLRunBkgdOperation     =  3,
		eMLRunSpectrumOperation =  4,
		eMLGainAdjust			=  5, // only runs interferogram data
	};

	public enum MLMETHODLISTHIGHLIGHTTYPE
	{
		eMLInvalid      = -1,
		eMLNorm         =  1,
		eMLActiveOrNone =  2,
		eMLSaved        =  3,
		eMLAbortedEdit  =  4,
	};

	// JSA 04-24-07 added custom field type
	public enum CUSTOMFIELDTYPE
	{
		eInvalid = -1, // invalid
		eString  =  1, // text
		eInt     =  2, // integer
		eDouble  =  3, // floating point
		eDate    =  4, // date
	};

	// JSA 07-10-07 added instrument type enumerators
	public enum ML_INSTRUMENT_TYPE
	{
		eInstrumentType_Undefined = 0,
		eInstrumentType_First     = 1,
		eInstrumentType_ML        = 1, // ML(PC)
		eInstrumentType_MLP       = 2, // MLP
		eInstrumentType_MLX       = 3, // MLX
		eInstrumentType_Exoscan   = 4, // aka MLs or Exoscan
		eInstrumentType_Last      = 4, // aka MLs or Exoscan
	}

	public enum ML_PATH_LENGTH
	{
		// (these values may change in the future)
		ePathLength_Undefined = 0,
		// Gas Path Values
		ePathLength_10_CM     = 1,
		ePathLength_15_CM     = 2,
		ePathLength_20_CM     = 3,
		// Transmission Path Values
		ePathLength_100_UM    = 4,
		ePathLength_Last      = 4,
	}

	// JSA 07-23-07 added run mode type enumerators (refers to any current run)
	public enum ML_RUN_MODE
	{
		eRunMode_Undefined         = 0,
		eRunMode_First             = 1,
		eRunMode_Normal            = 1,
		eRunMode_PerformanceTest   = 2, // performance test (signal-to-noise)
		eRunMode_StabilityTest     = 3, // stability test
		eRunMode_LaserFreqCalTest  = 4, // laser frequency calibration check
		eRunMode_PathlengthCalTest = 5, // pathlength calibration check
		eRunMode_GainAdjust		   = 6,
		eRunMode_Last              = 6,
	}

	// JSA 07-23-07 added system check type enumerators (refers to any currently selected system checks)
	[Flags]
	public enum ML_SYSTEM_CHECK_TYPE
	{
		SCT_NONE        = 0x00000000,
		SCT_PERFORMANCE = 0x00000001, // performance test (signal-to-noise)
		SCT_STABILITY   = 0x00000010, // stability test
		SCT_LASER       = 0x00000100, // laser frequency calibration check
		SCT_PATHLENGTH  = 0x00001000, // pathlength calibration check
	}

	// JSA 07-31-07 added UI image type enumerators (refers to a UI image like a photo or diagram)
	public enum ML_UI_IMAGE_TYPE
	{
		eImageType_Undefined             = 0,
		eImageType_First                 = 1,
		eImageType_MainMenu              = 1,
		eImageType_CleanCrystal1         = 2,
		eImageType_CleanCrystal2         = 3,
		eImageType_PrepSample1           = 4,
		eImageType_PrepSample2           = 5,
		eImageType_ClosedMechanism       = 6,
		eImageType_ClosedMechanism_small = 7,
		eImageType_ClosedPolystyrene     = 8,
		eImageType_Last                  = 8,
	}

	// JAL added Method Trend information type (bit fields)
    // typically can be (RATIO or SUBTRACT) and (PEAK or AREA)
    [Flags]
    public enum METH_TRENDTYPE { METH_TT_NONE = 0, METH_TT_RATIO = 0x00000001, METH_TT_SUBTRACT = 0x00000002, METH_TT_PEAK = 0x00010000, METH_TT_AREA = 0x00020000, };

	// JAL added enums for storing results
    public enum CURVE_DATA_TYPE { CDT_NONE = 0, 
                            CDT_RAWDATA     = 0x00000001, CDT_DARKDATA     = 0x00000002, CDT_BKGREFDATA    = 0x00000004, CDT_SPECTRUMDATA = 0x00000008,
                            CDT_RAMANDATA   = 0x00000010, CDT_MSPROFILEDATA= 0x00000020, CDT_MSCENTROIDDATA= 0x00000040, CDT_TICDATA      = 0x00000080,  
                            CDT_LCGCDATA    = 0x00000100, CDT_CALIBDATA    = 0x00000200, };
    
    // JAL had to change TREND from 4 to 16 and lose 4 altogether because this is a bit field and 
    // COMPONENT results was set to 5 (QWUAL | TREND)
    // Now when reading in from older version results files -- will look for 5 (QUAL|Illegal) and will reset to COMPONENT.
    [Flags]
    public enum RESULTS_DATA_TYPE { RDT_NONE = 0,
    RDT_QUALRESULTS = 0x00000001, RDT_QUANTRESULTS = 0x00000002, IllegalValue = 0x00000004, RDT_COMPONENTRESULTS = 0x00000008, RDT_TRENDRESULTS = 0x00000010, };

	public enum PLUGIN_ID { PLUGIN_GALACTIC = 0x5ac3, PLUGIN_SMITHSDETECTION = 0x5d91, PLUGIN_A2TECHNOLOGIES = 0x5ff5, };

    public enum EXTERNAL_FILE_TYPE { EFTYPE_A2R = 0, EFTYPE_ASP = 1, EFTYPE_SPC = 2 };

    public enum BKGVALID_DISPLAYTYPES { BV_MINUTES = 0, BV_HOURS = 1, BV_DAYS = 2 };
}