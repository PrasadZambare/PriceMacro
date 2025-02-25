using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PriceMacro.Models
{
	public class PowerAppsOutputModels
	{
        public string TermInMonths { get; set; }
        public string FirstTermPTPQ { get; set; }
        public string AssetCategory { get; set; }
        public string DepositPercentage { get; set; }
        public string GRVAsPercentageOfVolume { get; set; }
        public string RentalRebate { get; set; }
        public string Rating { get; set; }
        public string Volume { get; set; }
        public string GSTRate { get; set; }
        public string GSTValue { get; set; }
        public string TotalColumn { get; set; }
        public string DebtRateQuarterly { get; set; }
        public string EligibleGST { get; set; }
        public string VolumeConsideredForRentalWorking { get; set; }
        public string InterimDays { get; set; }
        public string FirmTermRentalDate { get; set; }
        public string RentalFrequency { get; set; }

        // Deal Details
        public string DealTargetPNI { get; set; }
        public string Slab1 { get; set; }

        // Funder Details
        public string FunderPVCap { get; set; }
        public string FunderDiscountingRate { get; set; }
        public string FunderDiscountingRateType { get; set; }
        public string QuarterlyRate { get; set; }
        public string AnnualizedRate { get; set; }
        public string DateOfDiscounting { get; set; }

        // Client Details
        public string ClientName { get; set; }
        public string Tenure { get; set; }
        public string PTPMLabel { get; set; }
        public string PTPMValue1 { get; set; }
        public string PTPMValue2 { get; set; }
        public string DepositValue1 { get; set; }
        public string DepositValue2 { get; set; }

        // Financial Calculations
        public string GVRValue1 { get; set; }
        public string GVRValue2 { get; set; }
        public string GSTValue1 { get; set; }
        public string GSTValue2 { get; set; }
        //public decimal RentalRebate { get; set; }
        public string XIRR { get; set; }
        public string PNI { get; set; }
        public string Rating6 { get; set; }

        // Debt Details
        public string DebtRate { get; set; }
        public string DebtRateLabel { get; set; }
        public string DebtRateValue { get; set; }

        // Additional Financial Details
        public string PNI7 { get; set; }
        public string FeeOrInvestment { get; set; }
        public string FunderPVCapPercentage { get; set; }
        public string ActualFunderPV { get; set; }
        public string RNSCount { get; set; }
        public string TotalRental { get; set; }
        public string PNIAdjustedVolume { get; set; }
        public string NumberOfPayments { get; set; }
        public string XIRRError { get; set; }
        public string TotalRentalError { get; set; }

        // Identifiers
        public string ID { get; set; }
        public string MacrosList { get; set; }
        public string Created { get; set; }
        public string RentalPaymentType { get; set; }
        public string Modified { get; set; }
    }
}