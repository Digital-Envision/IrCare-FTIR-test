using System;
using System.Collections.Generic;
using System.Text;

namespace CommonShared
{
	#region enumerated types

	// NOTE: these values currently match standard Grams ASP files
	// (they are not currently consistent with the Spectrum.cs file)
	public enum DATAXTYPE
	{
		XT_ARB = 0,
		XT_WN = 1,
		XT_uM = 2,
		XT_nM = 3,
		XT_Seconds = 4,
		XT_Minutes = 5,
		XT_MassCharge = 9,
		XT_RAMSHFT = 13,
		XT_Points = 22,
		XT_Hours = 30,
		XT_Time = 40,
		XT_AMU = 50,
		XT_Custom = 51
	};

	// NOTE: these values currently match standard Grams ASP files
	// (they are not currently consistent with the Spectrum.cs file)
	public enum DATAYTYPE
	{
		YT_ARB = 0,
		YT_IGRAM = 1,
		YT_Abs = 2,
		YT_Percent = 11,
		YT_Intensity = 12,
		YT_RelAbundance = 13,
		YT_Concentration = 40,
		YT_Trans = 128,	// (reversed y-axis high/low values)
		YT_Refl = 129,// (reversed y-axis high/low values)
		YT_Custom = 51,
		YT_Abundance = 52,
	};

	// JSA 04-04-07 Added component calculation type
	public enum COMPONENTCALCTYPE
	{
		INVALID = -1,
		PEAK_HEIGHT_NO_BASE = 1,
		PEAK_AREA_NO_BASE = 2,
		PEAK_HEIGHT_SINGLE_PT_BASE = 3,
		PEAK_AREA_SINGLE_PT_BASE = 4,
		PEAK_HEIGHT_DUAL_PT_BASE = 5,
        PEAK_AREA_DUAL_PT_BASE = 6,
        RMS_CALC = 7,
        PEAK_POSITION = 8,
#if _THERMOQUANT_
// JSA 11/20/07 removed the notion of PLS1 and PLS2 being separate calculation types
		// JAL added back for backward compatibility with methods previously saved with these component calc types
		QUANT_MODEL_PLS1 = 10, // Older seperation into pls1 and pls2
		QUANT_MODEL_PLS2 = 11,
		QUANT_MODEL = 12,
#else
		RESERVED1 = 10,
		RESERVED2 = 11,
		RESERVED3 = 12,
#endif
		PEAK_RATIO = 13,
	};

	public enum COMPONENTVALUETYPE
	{
		ACTUAL = 0,		// default - show actual value calculated
		PCT_LOW = 1,	// show as Percent of way to low critical value
		PCT_HIGH = 2,	// show as Percent of way to high critical value
	};


    public enum COMPONENTRESULTSTATE 
    { 
        CRS_Good = 1, 
        CRS_Unknown = 0, 
        CRS_MarginalLow = -1, 
        CRS_MarginalHigh = -2, 
        CRS_CriticalLow = -3, 
        CRS_CriticalHigh = -4, 
        CRS_BadCalculation = -5, 
        CRS_OverrideValue = -6, 
        CRS_UsupportedCalculationType = -7,
        CRS_DataRangeError = -8,
		CRS_LimitsRangeError = -9,
    };

	#endregion
}
