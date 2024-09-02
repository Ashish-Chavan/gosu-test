package com.zna.zsl.exports

uses com.zna.common.util.ReleaseDateUtil_ZNA
uses com.zna.services.ultraflex.dto.RateChangeData
uses com.zna.zsl.exports.dto.RateChangeInputData
uses com.zna.zsl.exports.dto.RateChangeOutputData
uses com.zna.zsl.exports.dto.Year1RateChangeData
uses com.zna.zsl.exports.dto.Year2RateChangeData
uses com.zna.zsl.financials.ZSLPricingDisplayUtil
uses gw.api.util.ConfigAccess
uses gw.lob.zsl.coverageinfo.v1.ZSLLineEPLCoverageInfo_ZNA
uses org.apache.poi.ss.usermodel.BorderStyle
uses org.apache.poi.ss.usermodel.DataFormatter
uses org.apache.poi.ss.usermodel.IndexedColors
uses org.apache.poi.ss.util.CellReference
uses org.apache.poi.xssf.usermodel.XSSFSheet
uses org.apache.poi.xssf.usermodel.XSSFWorkbook

uses java.io.ByteArrayOutputStream
uses java.io.FileInputStream
uses java.math.BigDecimal
uses java.math.RoundingMode
uses java.text.SimpleDateFormat

/**
 * Created by 841708 on 29-04-2022.
 */
class Year2RateChange {
  var premiumCalculationTemplate = "ZSP_Year2RateChange_ZNA_Template_v1.0.xlsx"
  var templateFilePath = "/config/resources/exporttemplates/" + premiumCalculationTemplate
  var outputName = "ZSP_Year2RateChange"
  var mimeType_xlsx = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  var styleYellowHeader : org.apache.poi.xssf.usermodel.XSSFCellStyle = null
  var _renewalPeriod : PolicyPeriod
  var _expiringPeriod : PolicyPeriod
  var _restatedPeriod : PolicyPeriod


  //Using Apache POI XSSF to create Excel workbooks in .xlxs format
  public function Year2RateChangeWorkbook(period : PolicyPeriod, expiringPeriod : PolicyPeriod) : Year2RateChangeData {
    var year2RateChangeData = new Year2RateChangeData()
    var inputData = new RateChangeInputData()
    var outputData = new RateChangeOutputData()
    _renewalPeriod = period
    _expiringPeriod = expiringPeriod
    _restatedPeriod = period.Job.Periods.firstWhere(\elt -> elt.PolicyPeriodType_ZNA == PolicyPeriodType_ZNA.TC_RESTATED)
    year2RateChangeData.InputData = inputData
    year2RateChangeData.OutputData = outputData

    year2RateChangeData = GetExpiringInputData(expiringPeriod, year2RateChangeData)
    year2RateChangeData = GetRenewingInputData(period, year2RateChangeData)
    if (_restatedPeriod != null) {
      year2RateChangeData = GetRestatedInputData(_restatedPeriod, year2RateChangeData)
    }

    var byteArrayOutputStream = CreateYear2RateChangeExport(year2RateChangeData)
    year2RateChangeData.WorkbookBAOS = byteArrayOutputStream.toByteArray()

    return year2RateChangeData
  }

 /* private function GetExpiringInputData(period : PolicyPeriod, year2RateChangeData : Year2RateChangeData) : Year1RateChangeData {
    //Moving expiring data from the GetUltraFlexRateChangeData service output to the Year 1 Rate Change Input
    var rateChangeInput = year2RateChangeData.InputData

    rateChangeInput.Expiring_PolicyNumber = (period.PolicyNumber == null) ? "" : period.PolicyNumber
    rateChangeInput.Expiring_Insured = (period.PrimaryInsuredName == null) ? "" : period.PrimaryInsuredName
    rateChangeInput.Expiring_EffectiveDate = ((period.PeriodStart == null) ? "" : period.PeriodStart.asMMddyyyy_ZNA) as String
    rateChangeInput.Expiring_ExpirationDate = ((period.PeriodEnd == null) ? "" : period.PeriodEnd.asMMddyyyy_ZNA) as String
    rateChangeInput.Expiring_Iteration = (period.Iteration == null) ? "" : expiringData.Iteration
    rateChangeInput.Expiring_DomiciledState = (period.BaseState.Code == null) ? "" : period.BaseState.Code
    rateChangeInput.Expiring_Private_NonProfit = period.Policy.Account.OtherOrgTypeDescription
    rateChangeInput.Expiring_TypeOfNonProfit = (expiringData.TypeOfNonProfit == null) ? "" : expiringData.TypeOfNonProfit
    rateChangeInput.Expiring_FidelityClassCode = (expiringData.FidelityClassCode == null) ? "" : expiringData.FidelityClassCode
    rateChangeInput.Expiring_Commission = (expiringData.Commission < 0) ? 0 : expiringData.Commission
    rateChangeInput.Expiring_Primary_Excess = (expiringData.Primary_Excess == null) ? "" : expiringData.Primary_Excess
    rateChangeInput.Expiring_PolicyLimit = (expiringData.PolicyLimit < 0) ? 0 : expiringData.PolicyLimit
    rateChangeInput.Expiring_PolicyAggregateLimit = (expiringData.PolicyAggregateLimit < 0) ? 0 : expiringData.PolicyAggregateLimit
    rateChangeInput.Expiring_PolicyAttachmentPoint = (expiringData.PolicyAttachmentPoint < 0) ? 0 : expiringData.PolicyAttachmentPoint
    rateChangeInput.Expiring_PolicySIR = (expiringData.PolicySIR < 0) ? 0 : expiringData.PolicySIR
    rateChangeInput.Expiring_PolicyAggSIR = (expiringData.PolicyAggSIR < 0) ? 0 : expiringData.PolicyAggSIR
    rateChangeInput.Expiring_UniqueAndUnusual = (expiringData.UniqueAndUnusual == null) ? "" : expiringData.UniqueAndUnusual
    rateChangeInput.Expiring_TotalAssets = (expiringData.TotalAssets < 0) ? 0 : expiringData.TotalAssets
    rateChangeInput.Expiring_TotalPlanAssets = (expiringData.TotalPlanAssets  < 0) ? 0 : expiringData.TotalPlanAssets
    rateChangeInput.Expiring_SharedLimit = (expiringData.SharedLimit == null) ? "" : expiringData.SharedLimit
    rateChangeInput.Expiring_ActualCharged = (expiringData.ActualCharged < 0) ? 0 : expiringData.ActualCharged
    rateChangeInput.Expiring_CrimeTotalLocations = (expiringData.CrimeTotalLocations < 0) ? 0 : expiringData.CrimeTotalLocations
    rateChangeInput.Expiring_FullTimeEmployeesUS = (expiringData.FullTimeEmployeesUS < 0) ? 0 : expiringData.FullTimeEmployeesUS
    rateChangeInput.Expiring_PartTimeEmployees = (expiringData.PartTimeEmployees < 0) ? 0 : expiringData.PartTimeEmployees
    rateChangeInput.Expiring_IndependentContractors = (expiringData.IndependentContractors < 0) ? 0 : expiringData.IndependentContractors
    rateChangeInput.Expiring_ForeignEmployees = (expiringData.ForeignEmployees < 0) ? 0 : expiringData.ForeignEmployees
    rateChangeInput.Expiring_UnionEmployees = (expiringData.UnionEmployees < 0) ? 0 : expiringData.UnionEmployees
    rateChangeInput.Expiring_State1 = (expiringData.State1 == null) ? "" : expiringData.State1
    rateChangeInput.Expiring_State2 = (expiringData.State2 == null) ? "" : expiringData.State2
    rateChangeInput.Expiring_State3 = (expiringData.State3 == null) ? "" : expiringData.State3
    rateChangeInput.Expiring_State4 = (expiringData.State4 == null) ? "" : expiringData.State4
    rateChangeInput.Expiring_State5 = (expiringData.State5 == null) ? "" : expiringData.State5
    rateChangeInput.Expiring_FullTimeEmployeesUSState1 = (expiringData.FullTimeEmployeesUSState1 < 0) ? 0 : expiringData.FullTimeEmployeesUSState1
    rateChangeInput.Expiring_FullTimeEmployeesUSState2 = (expiringData.FullTimeEmployeesUSState2 < 0) ? 0 : expiringData.FullTimeEmployeesUSState2
    rateChangeInput.Expiring_FullTimeEmployeesUSState3 = (expiringData.FullTimeEmployeesUSState3 < 0) ? 0 : expiringData.FullTimeEmployeesUSState3
    rateChangeInput.Expiring_FullTimeEmployeesUSState4 = (expiringData.FullTimeEmployeesUSState4 < 0) ? 0 : expiringData.FullTimeEmployeesUSState4
    rateChangeInput.Expiring_FullTimeEmployeesUSState5 = (expiringData.FullTimeEmployeesUSState5 < 0) ? 0 : expiringData.FullTimeEmployeesUSState5
    rateChangeInput.Expiring_FullTimeEmployeesForeign = (expiringData.FullTimeEmployeesForeign < 0) ? 0 : expiringData.FullTimeEmployeesForeign
    rateChangeInput.Expiring_FullTimeEmployeesAllOther = (expiringData.FullTimeEmployeesAllOther < 0) ? 0 : expiringData.FullTimeEmployeesAllOther
    rateChangeInput.Expiring_RateableEmployeesState1 = (expiringData.RateableEmployeesState1 < 0) ? 0 : Math.round(expiringData.RateableEmployeesState1)
    rateChangeInput.Expiring_RateableEmployeesState2 = (expiringData.RateableEmployeesState2 < 0) ? 0 : Math.round(expiringData.RateableEmployeesState2)
    rateChangeInput.Expiring_RateableEmployeesState3 = (expiringData.RateableEmployeesState3 < 0) ? 0 : Math.round(expiringData.RateableEmployeesState3)
    rateChangeInput.Expiring_RateableEmployeesState4 = (expiringData.RateableEmployeesState4 < 0) ? 0 : Math.round(expiringData.RateableEmployeesState4)
    rateChangeInput.Expiring_RateableEmployeesState5 = (expiringData.RateableEmployeesState5 < 0) ? 0 : Math.round(expiringData.RateableEmployeesState5)
    rateChangeInput.Expiring_RateableEmployeesForeign = (expiringData.RateableEmployeesForeign < 0) ? 0 : Math.round(expiringData.RateableEmployeesForeign)
    rateChangeInput.Expiring_RateableEmployeesAllOther = (expiringData.RateableEmployeesAllOther < 0) ? 0 : Math.round(expiringData.RateableEmployeesAllOther)
    rateChangeInput.Expiring_EPL_RateableEmployees = (expiringData.EPL_RateableEmployees < 0) ? 0 : Math.round(expiringData.EPL_RateableEmployees)
    rateChangeInput.Expiring_DO_YN = (expiringData.DO_YN == null) ? "" : expiringData.DO_YN
    rateChangeInput.Expiring_DO_SIR = (expiringData.DO_SIR < 0) ? 0 : expiringData.DO_SIR
    rateChangeInput.Expiring_DO_SharedLimit = (expiringData.DO_SharedLimit == null) ? "" : expiringData.DO_SharedLimit
    rateChangeInput.Expiring_DO_Limit = (expiringData.DO_Limit < 0) ? 0 : expiringData.DO_Limit
    rateChangeInput.Expiring_DO_AttachmentPoint = (expiringData.DO_AttachmentPoint < 0) ? 0 : expiringData.DO_AttachmentPoint
    rateChangeInput.Expiring_DO_CombinedCoverageDiscount = (expiringData.DO_CombinedCoverageDiscount < 0 ) ? 0 : expiringData.DO_CombinedCoverageDiscount
    rateChangeInput.Expiring_DO_BasePremium = (expiringData.DO_BasePremium < 0) ? 0 : expiringData.DO_BasePremium
    rateChangeInput.Expiring_DO_AnnualChargedPremium = (expiringData.DO_AnnualChargedPremium < 0) ? 0 : expiringData.DO_AnnualChargedPremium
    rateChangeInput.Expiring_EPL_YN = (expiringData.EPL_YN == null) ? "" : expiringData.EPL_YN
    rateChangeInput.Expiring_EPL_SIR = (expiringData.EPL_SIR < 0) ? 0 : expiringData.EPL_SIR
    rateChangeInput.Expiring_EPL_SharedLimit = (expiringData.EPL_SharedLimit == null) ? "" : expiringData.EPL_SharedLimit
    rateChangeInput.Expiring_EPL_Limit = (expiringData.EPL_Limit < 0) ? 0 : expiringData.EPL_Limit
    rateChangeInput.Expiring_EPL_AttachmentPoint = (expiringData.EPL_AttachmentPoint < 0) ? 0 : expiringData.EPL_AttachmentPoint
    rateChangeInput.Expiring_EPL_CombinedCoverageDiscount = (expiringData.EPL_CombinedCoverageDiscount < 0 ) ? 0 : expiringData.EPL_CombinedCoverageDiscount
    rateChangeInput.Expiring_EPL_SeparateLimitSurcharge = (expiringData.EPL_SeparateLimitSurcharge < 0) ? 0 : expiringData.EPL_SeparateLimitSurcharge
    rateChangeInput.Expiring_EPL_BasePremium = (expiringData.EPL_BasePremium < 0) ? 0 : expiringData.EPL_BasePremium
    rateChangeInput.Expiring_EPL_AnnualChargedPremium = (expiringData.EPL_AnnualChargedPremium < 0) ? 0 : expiringData.EPL_AnnualChargedPremium
    rateChangeInput.Expiring_FID_YN = (expiringData.FID_YN == null) ? "" : expiringData.FID_YN
    rateChangeInput.Expiring_FID_SIR = (expiringData.FID_SIR < 0) ? 0 : expiringData.FID_SIR
    rateChangeInput.Expiring_FID_SharedLimit = (expiringData.FID_SharedLimit == null) ? "" : expiringData.FID_SharedLimit
    rateChangeInput.Expiring_FID_Limit = (expiringData.FID_Limit < 0) ? 0 : expiringData.FID_Limit
    rateChangeInput.Expiring_FID_AttachmentPoint = (expiringData.FID_AttachmentPoint < 0) ? 0 : expiringData.FID_AttachmentPoint

    rateChangeInput.Expiring_FID_CombinedCoverageDiscount = (expiringData.FID_CombinedCoverageDiscount < 0) ? 0 : expiringData.FID_CombinedCoverageDiscount
    rateChangeInput.Expiring_FID_SeparateLimitSurcharge = (expiringData.FID_SeparateLimitSurcharge < 0) ? 0 : expiringData.FID_SeparateLimitSurcharge
    rateChangeInput.Expiring_FID_BasePremium = (expiringData.FID_BasePremium < 0) ? 0 : expiringData.FID_BasePremium
    rateChangeInput.Expiring_FID_AnnualChargedPremium = (expiringData.FID_AnnualChargedPremium < 0) ? 0 : expiringData.FID_AnnualChargedPremium
    rateChangeInput.Expiring_CR_YN = (expiringData.CR_YN == null) ? "" : expiringData.CR_YN
    rateChangeInput.Expiring_CR_SharedLimit = (expiringData.CR_SharedLimit == null) ? "" : expiringData.CR_SharedLimit
    rateChangeInput.Expiring_CR_Limit = (expiringData.CR_Limit < 0) ? 0 : expiringData.CR_Limit
    rateChangeInput.Expiring_CR_AttachmentPoint = (expiringData.CR_AttachmentPoint < 0) ? 0 : expiringData.CR_AttachmentPoint
    rateChangeInput.Expiring_CR_CombinedCoverageDiscount = (expiringData.CR_CombinedCoverageDiscount < 0) ? 0 : expiringData.CR_CombinedCoverageDiscount
    rateChangeInput.Expiring_CR_SeparateLimitSurcharge = (expiringData.CR_SeparateLimitSurcharge < 0) ? 0 : expiringData.CR_SeparateLimitSurcharge
    rateChangeInput.Expiring_CR_BasePremium = (expiringData.CR_BasePremium < 0) ? 0 : expiringData.CR_BasePremium
    rateChangeInput.Expiring_CR_EmplTheftAnnualChargedPremium = (expiringData.CR_EmployeeTheftAnnualChargedPremium < 0) ? 0 : expiringData.CR_EmployeeTheftAnnualChargedPremium
    rateChangeInput.Expiring_CR_TotalAnnualChargedPremium = (expiringData.CR_TotalAnnualChargedPremium < 0) ? 0 : expiringData.CR_TotalAnnualChargedPremium
    rateChangeInput.Expiring_CR_EndorsementPremium = (expiringData.CR_EndorsementPremium < 0) ? 0 : expiringData.CR_EndorsementPremium
    rateChangeInput.Expiring_CR_LimitPerClaim = (expiringData.CR_LimitPerClaim < 0) ? 0 : expiringData.CR_LimitPerClaim
    rateChangeInput.Expiring_CR_Deductible = (expiringData.CR_Deductible < 0) ? 0 : expiringData.CR_Deductible
    rateChangeInput.Expiring_CR_RateableEmployees = (expiringData.CR_RateableEmployees < 0) ? 0 : Math.round(expiringData.CR_RateableEmployees)

    if (rateChangeInput.Expiring_Primary_Excess == "Excess") {
      //calcs to determine premium at coverage part level for XS
      rateChangeInput.Expiring_XS_YN = (expiringData.XS_YN == null) ? "" : expiringData.XS_YN
      rateChangeInput.Expiring_XS_AnnualChargedPremium = (expiringData.XS_AnnualChargedPremium < 0) ? 0 : expiringData.XS_AnnualChargedPremium
      calculateCoveragePremiumForExcess(expiringData, rateChangeInput)
      //calc to determine ratable employees
      calculateRatableEmployeesForExcess(expiringData, rateChangeInput)

    }

    return year1RateChangeData

  }*/

  private function calculateCoveragePremiumForExcess(expiringData : RateChangeData, rateChangeInput : RateChangeInputData) {

    var totalUnderlyingPremium : double = 0
    var ratio : BigDecimal

    totalUnderlyingPremium += (expiringData.DO_UnderlyingPolicyPremium < 0) ? 0 : expiringData.DO_UnderlyingPolicyPremium
    totalUnderlyingPremium += (expiringData.EPL_UnderlyingPolicyPremium < 0) ? 0 : expiringData.EPL_UnderlyingPolicyPremium
    totalUnderlyingPremium += (expiringData.FID_UnderlyingPolicyPremium < 0) ? 0 : expiringData.FID_UnderlyingPolicyPremium
    totalUnderlyingPremium += (expiringData.CR_UnderlyingPolicyPremium < 0) ? 0 : expiringData.CR_UnderlyingPolicyPremium

    if (totalUnderlyingPremium != 0) {

      if (expiringData.DO_UnderlyingPolicy_YN == "Yes") {
        rateChangeInput.Expiring_DO_AnnualChargedPremium = Math.round(expiringData.DO_UnderlyingPolicyPremium / totalUnderlyingPremium * rateChangeInput.Expiring_XS_AnnualChargedPremium)
      }
      if (expiringData.EPL_UnderlyingPolicy_YN == "Yes") {
        rateChangeInput.Expiring_EPL_AnnualChargedPremium = Math.round(expiringData.EPL_UnderlyingPolicyPremium / totalUnderlyingPremium * rateChangeInput.Expiring_XS_AnnualChargedPremium)
      }
      if (expiringData.FID_UnderlyingPolicy_YN == "Yes") {
        rateChangeInput.Expiring_FID_AnnualChargedPremium = Math.round(expiringData.FID_UnderlyingPolicyPremium / totalUnderlyingPremium * rateChangeInput.Expiring_XS_AnnualChargedPremium)
      }
      if (expiringData.CR_UnderlyingPolicy_YN == "Yes") {
        rateChangeInput.Expiring_CR_EmployeeTheftAnnualChargedPremium = Math.round(expiringData.CR_UnderlyingPolicyPremium / totalUnderlyingPremium * rateChangeInput.Expiring_XS_AnnualChargedPremium)
      }

    }

  }

  private function calculateRatableEmployeesExpiring(y2RateChangeData : Year2RateChangeData , expiringPeriod : PolicyPeriod) : Year2RateChangeData {

    var totalRatableEmployees : double = 0
    var totalFullTimeUS : double = 0
    var totalState1 : double = 0
    var totalState2 : double = 0
    var totalState3 : double = 0
    var totalState4 : double = 0
    var totalState5 : double = 0
    var totalCA : double = 0
    var totalCW : double = 0

    var coverageInfo = new ZSLLineEPLCoverageInfo_ZNA(expiringPeriod.ZSLLine.ZSL_EPL_Cov_ZNA, expiringPeriod.ZSLLine)
    coverageInfo.PopulateStates()

    totalFullTimeUS += expiringPeriod.ZSLLine.ZSLEPLIExposure_ZNA.Employees?.firstWhere(\emp -> emp.EmployeeType == ZSLEmployeeType_ZNA.TC_FULLTIME)?.CurrentYear == null ? 0 : expiringPeriod.ZSLLine.ZSLEPLIExposure_ZNA.Employees?.firstWhere(\emp -> emp.EmployeeType == ZSLEmployeeType_ZNA.TC_FULLTIME)?.CurrentYear
    totalState1 += coverageInfo.State1RatableEmployees == null ? 0 :  coverageInfo.State1RatableEmployees as double
    totalState2 += coverageInfo.State2RatableEmployees == null ? 0 :  coverageInfo.State2RatableEmployees as double
    totalState3 += coverageInfo.State3RatableEmployees == null ? 0 :  coverageInfo.State3RatableEmployees as double
    totalState4 += coverageInfo.State4RatableEmployees == null ? 0 :  coverageInfo.State4RatableEmployees as double
    totalState5 += coverageInfo.State5RatableEmployees == null ? 0 :  coverageInfo.State5RatableEmployees as double
    totalCA += coverageInfo.CaliforniaRatableEmployees == null ? 0 :  coverageInfo.CaliforniaRatableEmployees as double
    totalCW += coverageInfo.CountryWideRatableEmployees == null ? 0 :  coverageInfo.CountryWideRatableEmployees as double

    if (totalFullTimeUS != 0) {

      y2RateChangeData.InputData.Expiring_RateableEmployeesState1 = totalState1 != 0 ? totalState1 : 0
      y2RateChangeData.InputData.Expiring_RateableEmployeesState2 = totalState2 != 0 ? totalState2 : 0
      y2RateChangeData.InputData.Expiring_RateableEmployeesState3 = totalState3 != 0 ? totalState3 : 0
      y2RateChangeData.InputData.Expiring_RateableEmployeesState4 = totalState4 != 0 ? totalState4 : 0
      y2RateChangeData.InputData.Expiring_RateableEmployeesState5 = totalState5 != 0 ? totalState5 : 0
      y2RateChangeData.InputData.Expiring_RateableEmployeesCA = totalCA != 0 ? totalCA : 0
      y2RateChangeData.InputData.Expiring_RateableEmployeesAllOther = totalCW != 0 ? totalCW : 0

    }
    return y2RateChangeData
  }

  private function calculateRatableEmployeesRestated(y2RateChangeData : Year2RateChangeData , restatedPeriod : PolicyPeriod) : Year2RateChangeData {

    var totalRatableEmployees : double = 0
    var totalFullTimeUS : double = 0
    var totalState1 : double = 0
    var totalState2 : double = 0
    var totalState3 : double = 0
    var totalState4 : double = 0
    var totalState5 : double = 0
    var totalCA : double = 0
    var totalCW : double = 0

    var coverageInfo = new ZSLLineEPLCoverageInfo_ZNA(restatedPeriod.ZSLLine.ZSL_EPL_Cov_ZNA, restatedPeriod.ZSLLine)
    coverageInfo.PopulateStates()

    totalFullTimeUS += restatedPeriod.ZSLLine.ZSLEPLIExposure_ZNA.Employees?.firstWhere(\emp -> emp.EmployeeType == ZSLEmployeeType_ZNA.TC_FULLTIME)?.CurrentYear == null ? 0 : restatedPeriod.ZSLLine.ZSLEPLIExposure_ZNA.Employees?.firstWhere(\emp -> emp.EmployeeType == ZSLEmployeeType_ZNA.TC_FULLTIME)?.CurrentYear
    totalState1 += coverageInfo.State1RatableEmployees == null ? 0 :  coverageInfo.State1RatableEmployees as double
    totalState2 += coverageInfo.State2RatableEmployees == null ? 0 :  coverageInfo.State2RatableEmployees as double
    totalState3 += coverageInfo.State3RatableEmployees == null ? 0 :  coverageInfo.State3RatableEmployees as double
    totalState4 += coverageInfo.State4RatableEmployees == null ? 0 :  coverageInfo.State4RatableEmployees as double
    totalState5 += coverageInfo.State5RatableEmployees == null ? 0 :  coverageInfo.State5RatableEmployees as double
    totalCA += coverageInfo.CaliforniaRatableEmployees == null ? 0 :  coverageInfo.CaliforniaRatableEmployees as double
    totalCW += coverageInfo.CountryWideRatableEmployees == null ? 0 :  coverageInfo.CountryWideRatableEmployees as double

    if (totalFullTimeUS != 0) {

      y2RateChangeData.InputData.Restated_RateableEmployeesState1 = totalState1 != 0 ? totalState1 : 0
      y2RateChangeData.InputData.Restated_RateableEmployeesState2 = totalState2 != 0 ? totalState2 : 0
      y2RateChangeData.InputData.Restated_RateableEmployeesState3 = totalState3 != 0 ? totalState3 : 0
      y2RateChangeData.InputData.Restated_RateableEmployeesState4 = totalState4 != 0 ? totalState4 : 0
      y2RateChangeData.InputData.Restated_RateableEmployeesState5 = totalState5 != 0 ? totalState5 : 0
      y2RateChangeData.InputData.Restated_RateableEmployeesCA = totalCA != 0 ? totalCA : 0
      y2RateChangeData.InputData.Restated_RateableEmployeesAllOther = totalCW != 0 ? totalCW : 0

    }
    return y2RateChangeData
  }
private function filterInvalidFilenameCharacters(filename : String) : String {
    return filename.replaceAll("[:\\\\/*?|<>\" ']", "_")
    }
  private function calculateRatableEmployeesRenewing(y2RateChangeData : Year2RateChangeData , renewingPeriod : PolicyPeriod) : Year2RateChangeData {

    var totalRatableEmployees : double = 0
    var totalFullTimeUS : double = 0
    var totalState1 : double = 0
    var totalState2 : double = 0
    var totalState3 : double = 0
    var totalState4 : double = 0
    var totalState5 : double = 0
    var totalCA : double = 0
    var totalCW : double = 0

    var coverageInfo = new ZSLLineEPLCoverageInfo_ZNA(renewingPeriod.ZSLLine.ZSL_EPL_Cov_ZNA, renewingPeriod.ZSLLine)
    coverageInfo.PopulateStates()

    totalFullTimeUS += renewingPeriod.ZSLLine.ZSLEPLIExposure_ZNA.Employees?.firstWhere(\emp -> emp.EmployeeType == ZSLEmployeeType_ZNA.TC_FULLTIME)?.CurrentYear == null ? 0 : renewingPeriod.ZSLLine.ZSLEPLIExposure_ZNA.Employees?.firstWhere(\emp -> emp.EmployeeType == ZSLEmployeeType_ZNA.TC_FULLTIME)?.CurrentYear
    totalState1 += coverageInfo.State1RatableEmployees == null ? 0 :  coverageInfo.State1RatableEmployees as double
    totalState2 += coverageInfo.State2RatableEmployees == null ? 0 :  coverageInfo.State2RatableEmployees as double
    totalState3 += coverageInfo.State3RatableEmployees == null ? 0 :  coverageInfo.State3RatableEmployees as double
    totalState4 += coverageInfo.State4RatableEmployees == null ? 0 :  coverageInfo.State4RatableEmployees as double
    totalState5 += coverageInfo.State5RatableEmployees == null ? 0 :  coverageInfo.State5RatableEmployees as double
    totalCA += coverageInfo.CaliforniaRatableEmployees == null ? 0 :  coverageInfo.CaliforniaRatableEmployees as double
    totalCW += coverageInfo.CountryWideRatableEmployees == null ? 0 :  coverageInfo.CountryWideRatableEmployees as double

    if (totalFullTimeUS != 0) {

      y2RateChangeData.InputData.Renewing_RateableEmployeesState1 = totalState1 != 0 ? totalState1 : 0
      y2RateChangeData.InputData.Renewing_RateableEmployeesState2 = totalState2 != 0 ? totalState2 : 0
      y2RateChangeData.InputData.Renewing_RateableEmployeesState3 = totalState3 != 0 ? totalState3 : 0
      y2RateChangeData.InputData.Renewing_RateableEmployeesState4 = totalState4 != 0 ? totalState4 : 0
      y2RateChangeData.InputData.Renewing_RateableEmployeesState5 = totalState5 != 0 ? totalState5 : 0
      y2RateChangeData.InputData.Renewing_RateableEmployeesCA = totalCA != 0 ? totalCA : 0
      y2RateChangeData.InputData.Renewing_RateableEmployeesAllOther = totalCW != 0 ? totalCW : 0

    }
    return y2RateChangeData
  }

  private function getExcessPremiumRatio(period : PolicyPeriod, covPart : CoveragePart_ZNA) : double {
    //Determine ratio of coverate part premium for total underlying coverage premium - Primary Layer only
    var ratio : BigDecimal
    var totalULPremium =  period.UnderlyingCoverages_ZNA.where(\elt -> elt.ProgramLayer == ProgramLayer_ZNA.TC_PRIMARY).sum(\elt -> elt.Premium)
    var doULPremium = period.UnderlyingCoverages_ZNA.firstWhere(\elt -> elt.ProgramLayer == ProgramLayer_ZNA.TC_PRIMARY && elt.CoveragePart == covPart).Premium
    if (totalULPremium != null and totalULPremium != 0) {
      ratio = doULPremium / totalULPremium
    }

    var excessCovCosts = period.ZSLLine.ZSL_ExcessLiability_Cov_ZNA.ZSLLineCovCosts
    var excessPremium : BigDecimal
    excessCovCosts?.each(\elt -> {
      excessPremium = (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
    })
    return (excessPremium * ratio).doubleValue()
  }


  /*Used for fecthing Expiring Term data*/

  private function GetExpiringInputData(period:PolicyPeriod, year2RateChangeData:Year2RateChangeData) : Year2RateChangeData {
    //Get all of the necessary data from the renewing policy period
    year2RateChangeData = GetExpiringPolicyData(year2RateChangeData, period)

    year2RateChangeData = GetExpiringDOData(year2RateChangeData, period)

    year2RateChangeData = GetExpiringEPLIData(year2RateChangeData, period)

    year2RateChangeData = GetExpiringFIDData(year2RateChangeData, period)

    year2RateChangeData = GetExpiringCrimeData(year2RateChangeData, period)

    year2RateChangeData = GetExpiringQuestionSetData(year2RateChangeData, period)

    year2RateChangeData = GetExpiringMiscData(year2RateChangeData, period)
    if (period.ZSLLine.EPLICoveragePartExists ) {
      year2RateChangeData = calculateRatableEmployeesExpiring(year2RateChangeData, period)
    }
    return year2RateChangeData
  }

  private function GetExpiringPolicyData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Policy
    year2RateChangeData.InputData.Expiring_PolicyNumber = (period.PolicyNumber != null) ?  period.PolicyNumber : period.Job.JobNumber

    year2RateChangeData.InputData.Expiring_QuoteNumber = (period.Job.JobNumber != null) ?  period.Job.JobNumber : ""

    //Insured
    year2RateChangeData.InputData.Expiring_Insured = (period.PrimaryInsuredName == null) ? "" : period.PrimaryInsuredName

    //EffectiveDate
    year2RateChangeData.InputData.Expiring_EffectiveDate = ((period.PeriodStart == null) ? "" : period.PeriodStart.asMMddyyyy_ZNA) as String

    //ExpirationDate
    year2RateChangeData.InputData.Expiring_ExpirationDate = ((period.PeriodEnd == null) ? "" : period.PeriodEnd.asMMddyyyy_ZNA) as String

    //State
    year2RateChangeData.InputData.Expiring_DomiciledState = (period.BaseState.Code == null) ? "" : period.BaseState.Code

    //CompanyType
    year2RateChangeData.InputData.Expiring_Private_NonProfit =  period.Policy.Account.OtherOrgTypeDescription

    //Commission
    year2RateChangeData.InputData.Expiring_Commission = (period.CommissionPercent_ZNA == null) ? 0 : period.CommissionPercent_ZNA.doubleValue()

    //NAICS
    year2RateChangeData.InputData.Expiring_NAICSCode = (period.PrimaryNamedInsured.NAICS_ZNA.Code == null) ? "" : period.PrimaryNamedInsured.NAICS_ZNA.Code

    //NAICS Description
    year2RateChangeData.InputData.Expiring_NAICSDescription = (period.PrimaryNamedInsured.NAICS_ZNA.Classification == null) ? " " : period.PrimaryNamedInsured.NAICS_ZNA.Classification

    //Industry Type
    year2RateChangeData.InputData.Expiring_IndustryType = (period.ZSLLine.ZSLIndustryType.DisplayName == null) ? "" : period.ZSLLine.ZSLIndustryType.DisplayName.toString()
    year2RateChangeData.InputData.Expiring_IndustryTypeCode = (period.ZSLLine.ZSLIndustryType.Code == null) ? "" : period.ZSLLine.ZSLIndustryType.Code

    //Premium
    year2RateChangeData.InputData.Expiring_ActualCharged = (period.TotalPremiumRPT_ZNA == null) ? 0 : period.TotalPremiumRPT_ZNA.doubleValue()

    //Asset Size
    year2RateChangeData.InputData.Expiring_TotalAssets = (period.ZSLLine.CurrentYearFinancial.TotalAssets.Amount == null) ? 0 : period.ZSLLine.CurrentYearFinancial.TotalAssets.Amount.doubleValue()

    //Plan Asset
    year2RateChangeData.InputData.Expiring_TotalPlanAssets = (period.ZSLLine.ZSLFiduciaryCovPart_ZNA.TotalPlanAssets == null) ? 0 : period.ZSLLine.ZSLFiduciaryCovPart_ZNA.TotalPlanAssets.doubleValue()

    //Policy Type - Primary or Excess
    year2RateChangeData.InputData.Expiring_Primary_Excess = (period.ZSLLine.PolicyType_ZNA == null) ? "" : period.ZSLLine.PolicyType_ZNA.toString()

    year2RateChangeData.InputData.Expiring_UniqueAndUnusual = (period.ZSLLine.UniqueUnusualInd) ? "Yes" : "No"


    return year2RateChangeData
  }

  private function GetExpiringDOData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    if (period.ZSLLine.ZSL_DandO_Cov_ZNA != null) {

      year2RateChangeData.InputData.Expiring_DO_YN = "Yes"

      if (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString != null) {
        year2RateChangeData.InputData.Expiring_DO_Limit = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Expiring_DO_Limit = (period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0 ) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      year2RateChangeData.InputData.Expiring_DO_LimitIsShared = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitIsShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitIsShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Expiring_DO_AttachmentPoint = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Expiring_DO_SIR = (period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA.doubleValue()

      var doCovPrices = period.ZSLLine.ZSL_DandO_Cov_ZNA.CoveragePrices
      doCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_DO_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Expiring_DO_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_DO_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Expiring_DO_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_DO)
      }

    }

    return year2RateChangeData
  }

  private function GetExpiringEPLIData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    if ( period.ZSLLine.ZSL_EPL_Cov_ZNA != null) {

      year2RateChangeData.InputData.Expiring_EPL_YN = "Yes"

      if(period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Expiring_EPL_Limit = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Expiring_EPL_Limit = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString.toDouble()
      }

      year2RateChangeData.InputData.Expiring_EPL_LimitIsShared = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZNA_EPL_LimitIsShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_EPL_Cov_ZNA.ZNA_EPL_LimitIsShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Expiring_EPL_AttachmentPoint = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Expiring_EPL_SIR = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString.toDouble()

      var eplCovPrice = period.ZSLLine.ZSL_EPL_Cov_ZNA.CoveragePrices.firstWhere(\elt -> elt.ZSPRatingPricingDetail.ZSLLineCostType != ZSLLineCovCostType_ZNA.TC_THIRD_PARTY_LIABILITY_EXCLUDED)
      year2RateChangeData.InputData.Expiring_EPL_BasePremium = (eplCovPrice.ZSPRatingPricingDetail.BasePremium == null) ? 0 : eplCovPrice.ZSPRatingPricingDetail.BasePremium.doubleValue()
      year2RateChangeData.InputData.Expiring_EPL_SharedLimitCredit = (eplCovPrice.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : eplCovPrice.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()

      //var zslLineCovCost = period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSLLineCovCosts.firstWhere(\elt -> !(elt typeis ZSLLineCovSubtypeCost_ZNA))
      //year1RateChangeData.InputData.Renewing_EPL_AnnualChargedPremium = (zslLineCovCost.ActualTermAmount_amt == null) ? 0 : zslLineCovCost.ActualTermAmount_amt.doubleValue()

      var zslLineCovCost = period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSLLineCovCosts.firstWhere(\elt -> !(elt typeis ZSLLineCovSubtypeCost_ZNA))
      year2RateChangeData.InputData.Expiring_EPL_AnnualChargedPremium += (zslLineCovCost.ActualTermAmount_amt == null) ? 0 : zslLineCovCost.ActualTermAmount_amt.doubleValue()

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Expiring_EPL_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_EPLI)
      }


      //Get EPL Rateable Employees Data
      if (period.ZSLLine.ZSLEPLIExposure_ZNA.Employees != null) {
        var sum : BigDecimal = 0
        var employees = period.ZSLLine.ZSLEPLIExposure_ZNA.Employees
        employees.each(\employee -> {
          var name = employee.EmployeeType.DisplayName
          var currentYear = (employee.CurrentYear < 1) ? 0 : employee.CurrentYear
          var ratableEmployeeFactor = (employee.RatableEmployeeFactor < 0) ? 0 : employee.RatableEmployeeFactor
          sum += (currentYear * ratableEmployeeFactor)
          switch (name) {
            case ZSLEmployeeType_ZNA.TC_FULLTIME.DisplayName:
              year2RateChangeData.InputData.Expiring_FullTimeEmployeesUS = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_PARTTIME.DisplayName:
              year2RateChangeData.InputData.Expiring_PartTimeEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_IND_CONTRACTERS.DisplayName:
              year2RateChangeData.InputData.Expiring_IndependentContractors = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_FOREIGN.DisplayName:
              year2RateChangeData.InputData.Expiring_ForeignEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_UNION.DisplayName:
              year2RateChangeData.InputData.Expiring_UnionEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_VOLUNTEERS.DisplayName:
              year2RateChangeData.InputData.Exipiring_Volunteers = currentYear
              break
            default:
              break
          }
        })
        year2RateChangeData.InputData.Expiring_EPL_RateableEmployees = (sum?.setScale(0, BigDecimal.ROUND_HALF_UP) < 0) ? 0 : sum?.setScale(0, BigDecimal.ROUND_HALF_UP).doubleValue()
      }

      //Get EPL Employees Data
      year2RateChangeData.InputData.Expiring_FullTimeEmployeesForeign = (period.ZSLLine.ZSLEPLIExposure_ZNA.ForeignEmployees == null) ? 0 : period.ZSLLine.ZSLEPLIExposure_ZNA.ForeignEmployees.doubleValue()

      //Get EPL State Employee Counts
      if (period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos != null) {
        var states = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos?.where(\elt -> !(elt.Jurisdiction == "CA" || elt.Jurisdiction == "All Other" || elt.Jurisdiction == null))
        states?.eachWithIndex(\state, indx -> {
          switch (indx) {
            case 0:
              year2RateChangeData.InputData.Expiring_State1 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Expiring_FullTimeEmployeesUSState1 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 1:
              year2RateChangeData.InputData.Expiring_State2 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Expiring_FullTimeEmployeesUSState2 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 2:
              year2RateChangeData.InputData.Expiring_State3 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Expiring_FullTimeEmployeesUSState3 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 3:
              year2RateChangeData.InputData.Expiring_State4 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Expiring_FullTimeEmployeesUSState4 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 4:
              year2RateChangeData.InputData.Expiring_State5 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Expiring_FullTimeEmployeesUSState5 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            default:
              break
          }
        })
        var stateCA = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos.firstWhere(\elt -> (elt.Jurisdiction == "CA"))
        year2RateChangeData.InputData.Expiring_FullTimeEmployeesUSStateCA = (stateCA.CurrentYear == null) ? 0 : stateCA.CurrentYear.doubleValue()

        var stateAllOther = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos.firstWhere(\elt -> (elt.Jurisdiction == "All Other"))
        year2RateChangeData.InputData.Expiring_FullTimeEmployeesAllOther = (stateAllOther.CurrentYear == null) ? 0 : stateAllOther.CurrentYear.doubleValue()
      }

    }

    return year2RateChangeData
  }

  private function GetExpiringFIDData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Get Fiduciary Data
    if (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA != null) {

      year2RateChangeData.InputData.Expiring_FID_YN = "Yes"

      if(period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Expiring_FID_Limit = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Expiring_FID_Limit = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      year2RateChangeData.InputData.Expiring_FID_LimitIsShared = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitisShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitisShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Expiring_FID_AttachmentPoint = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Expiring_FID_SIR = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm.ValueAsString.toDouble()

      var fidCovPrices = period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.CoveragePrices
      fidCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_FID_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Expiring_FID_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_FID_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Expiring_FID_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_FIDUCIARY)
      }

    }

    return year2RateChangeData
  }

  private function GetExpiringCrimeData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Get Crime Employee Theft Data
    if (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA != null) {

      year2RateChangeData.InputData.Expiring_CR_YN = "Yes"

      if(period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Expiring_CR_Limit = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Expiring_CR_Limit = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      //Get Employee Theft coverage data
      var crCovPrices = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.CoveragePrices
      crCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Expiring_CR_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
        year2RateChangeData.InputData.Expiring_CR_LocationFactor = (elt.ZSPRatingPricingDetail.LocationFactor == null) ? 0 : elt.ZSPRatingPricingDetail.LocationFactor.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_EmployeeTheftAnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Expiring_CR_EmployeeTheftAnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_CRIME)
      }

      year2RateChangeData.InputData.Expiring_CR_Deductible = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString.toDouble()

      //Crime Ratable Employees
      year2RateChangeData.InputData.Expiring_CR_RateableEmployees = (period.ZSLLine.ZSLCrimeCovPart_ZNA.RatableEmployeesCount.toDouble() < 0) ? 0 : period.ZSLLine.ZSLCrimeCovPart_ZNA.RatableEmployeesCount.toDouble()

      year2RateChangeData.InputData.Expiring_FidelityClassCode = (period.ZSLLine.ZSLCrimeCovPart_ZNA.ClassCode == null) ? "" : period.ZSLLine.ZSLCrimeCovPart_ZNA.ClassCode

      year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += year2RateChangeData.InputData.Renewing_CR_EmployeeTheftAnnualChargedPremium
    }

    //Get Total Crime Coverage premium

    if (period.ZSLLine.ZSL_ClientProperty_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_OnPremises_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_OnPremises_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_InTransit_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_InTransit_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNAExists) {
      var zslLineCondCosts = period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNA.ZSLLineCondCosts
      zslLineCondCosts?.each(\elt -> {
        year2RateChangeData.InputData.Expiring_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }

    return year2RateChangeData
  }

  private function GetExpiringQuestionSetData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    var periodLines = period.Lines
    periodLines?.each(\elt -> {
      var lineAnswers = elt.LineAnswers
      lineAnswers?.each(\elt2 -> {
        var answer = elt2.IntegerAnswer
        var codeID = elt2.Question.CodeIdentifier
        switch (codeID) {
          case "GITotalLocationsNumber_ZNA":
            year2RateChangeData.InputData.Expiring_CR_NumberOfLocations = (answer == null) ? 0 : answer.doubleValue()
            break
          default:
            break
        }
      })
    })

    return year2RateChangeData
  }

  private function GetExpiringMiscData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    var DO_LimitIsShared = (year2RateChangeData.InputData.Expiring_DO_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_DO_LimitIsShared
    var FID_LimitIsShared = (year2RateChangeData.InputData.Expiring_FID_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_FID_LimitIsShared
    var EPL_LimitIsShared = (year2RateChangeData.InputData.Expiring_EPL_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_EPL_LimitIsShared
    var DO_LimitofLiability = (year2RateChangeData.InputData.Expiring_DO_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_DO_Limit
    var FID_LimitofLiability = (year2RateChangeData.InputData.Expiring_FID_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_FID_Limit
    var EPL_LimitofLiability = (year2RateChangeData.InputData.Expiring_EPL_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_EPL_Limit
    var SharedAggregateLimit : double = 0
    var NumberOfSharedCovParts = 0
    var numberOfCovParts = 0

    if (year2RateChangeData.InputData.Expiring_DO_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Expiring_FID_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Expiring_EPL_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Expiring_CR_YN == "Yes") numberOfCovParts +=1

    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared != "true" ){
      year2RateChangeData.InputData.Expiring_SharedLimit = "D&O and EPL"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared != "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Expiring_SharedLimit = "D&O and Fiduciary"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared != "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Expiring_SharedLimit = "EPL & Fiduciary"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Expiring_SharedLimit = "D&O, EPL, and Fiduciary"
      NumberOfSharedCovParts = 3
    }
    if (DO_LimitIsShared != "true" && EPL_LimitIsShared != "true"  && FID_LimitIsShared != "true" ){
      year2RateChangeData.InputData.Expiring_SharedLimit = "Separate Limits"
    }
    if (numberOfCovParts == 1){
      year2RateChangeData.InputData.Expiring_SharedLimit = "Stand-Alone Coverage"
    }

    //if the limits are shared set the SharedAggregateLimit to the largest of the limits
    if (DO_LimitIsShared == "true" && (DO_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = DO_LimitofLiability
    }
    if (FID_LimitIsShared == "true" && (FID_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = FID_LimitofLiability
    }
    if (EPL_LimitIsShared == "true" && (EPL_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = EPL_LimitofLiability
    }

    year2RateChangeData.InputData.Expiring_NumberOfSharedCovParts = NumberOfSharedCovParts
    year2RateChangeData.InputData.Expiring_SharedAggregateLimit = SharedAggregateLimit





    return year2RateChangeData
  }




  private function GetRenewingInputData(period:PolicyPeriod, year2RateChangeData:Year2RateChangeData) : Year2RateChangeData {
    //Get all of the necessary data from the renewing policy period
    year2RateChangeData = GetRenewingPolicyData(year2RateChangeData, period)

    year2RateChangeData = GetRenewingDOData(year2RateChangeData, period)

    year2RateChangeData = GetRenewingEPLIData(year2RateChangeData, period)

    year2RateChangeData = GetRenewingFIDData(year2RateChangeData, period)

    year2RateChangeData = GetRenewingCrimeData(year2RateChangeData, period)

    year2RateChangeData = GetRenewingQuestionSetData(year2RateChangeData, period)

    year2RateChangeData = GetRenewingMiscData(year2RateChangeData, period)
    if (period.ZSLLine.EPLICoveragePartExists ) {
      year2RateChangeData = calculateRatableEmployeesRenewing(year2RateChangeData, period)
    }
    return year2RateChangeData
  }

  private function GetRestatedInputData(period:PolicyPeriod, year2RateChangeData:Year2RateChangeData) : Year2RateChangeData {
    //Get all of the necessary data from the renewing policy period
    year2RateChangeData = GetRestatedPolicyData(year2RateChangeData, period)

    year2RateChangeData = GetRestatedDOData(year2RateChangeData, period)

    year2RateChangeData = GetRestatedEPLIData(year2RateChangeData, period)

    year2RateChangeData = GetRestatedFIDData(year2RateChangeData, period)

    year2RateChangeData = GetRestatedCrimeData(year2RateChangeData, period)

    year2RateChangeData = GetRestatedQuestionSetData(year2RateChangeData, period)

    year2RateChangeData = GetRestatedMiscData(year2RateChangeData, period)
    if (period.ZSLLine.EPLICoveragePartExists ) {
      year2RateChangeData = calculateRatableEmployeesRestated(year2RateChangeData, period)
    }
    return year2RateChangeData
  }

  private function GetRenewingPolicyData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Policy
    year2RateChangeData.InputData.Renewing_PolicyNumber = (period.PolicyNumber != null) ?  period.PolicyNumber : period.Job.JobNumber

    year2RateChangeData.InputData.Renewing_QuoteNumber = (period.Job.JobNumber != null) ?  period.Job.JobNumber : ""
    year2RateChangeData.InputData.Renewing_BranchName = (period.BranchName != null) ?  period.BranchName : ""



    //Insured
    year2RateChangeData.InputData.Renewing_Insured = (period.PrimaryInsuredName == null) ? "" : period.PrimaryInsuredName

    //EffectiveDate
    year2RateChangeData.InputData.Renewing_EffectiveDate = ((period.PeriodStart == null) ? "" : period.PeriodStart.asMMddyyyy_ZNA) as String

    //ExpirationDate
    year2RateChangeData.InputData.Renewing_ExpirationDate = ((period.PeriodEnd == null) ? "" : period.PeriodEnd.asMMddyyyy_ZNA) as String

    //State
    year2RateChangeData.InputData.Renewing_DomiciledState = (period.BaseState.Code == null) ? "" : period.BaseState.Code

    //CompanyType
    year2RateChangeData.InputData.Renewing_Private_NonProfit =  period.Policy.Account.OtherOrgTypeDescription

    //Commission
    year2RateChangeData.InputData.Renewing_Commission = (period.CommissionPercent_ZNA == null) ? 0 : period.CommissionPercent_ZNA.doubleValue()

    //NAICS
    year2RateChangeData.InputData.Renewing_NAICSCode = (period.PrimaryNamedInsured.NAICS_ZNA.Code == null) ? "" : period.PrimaryNamedInsured.NAICS_ZNA.Code

    //NAICS Description
    year2RateChangeData.InputData.Renewing_NAICSDescription = (period.PrimaryNamedInsured.NAICS_ZNA.Classification == null) ? " " : period.PrimaryNamedInsured.NAICS_ZNA.Classification


    //Industry Type
    year2RateChangeData.InputData.Renewing_IndustryType = (period.ZSLLine.ZSLIndustryType.DisplayName == null) ? "" : period.ZSLLine.ZSLIndustryType.DisplayName.toString()
    year2RateChangeData.InputData.Renewing_IndustryTypeCode = (period.ZSLLine.ZSLIndustryType.Code == null) ? "" : period.ZSLLine.ZSLIndustryType.Code

    //Premium
    year2RateChangeData.InputData.Renewing_ActualCharged = (period.TotalPremiumRPT_ZNA == null) ? 0 : period.TotalPremiumRPT_ZNA.doubleValue()

    //Asset Size
    year2RateChangeData.InputData.Renewing_TotalAssets = (period.ZSLLine.CurrentYearFinancial.TotalAssets.Amount == null) ? 0 : period.ZSLLine.CurrentYearFinancial.TotalAssets.Amount.doubleValue()

    //Plan Asset
    year2RateChangeData.InputData.Renewing_PlanAssets = (period.ZSLLine.ZSLFiduciaryCovPart_ZNA.TotalPlanAssets == null) ? 0 : period.ZSLLine.ZSLFiduciaryCovPart_ZNA.TotalPlanAssets.doubleValue()

    //Policy Type - Primary or Excess
    year2RateChangeData.InputData.Renewing_Primary_Excess = (period.ZSLLine.PolicyType_ZNA == null) ? "" : period.ZSLLine.PolicyType_ZNA.toString()

    year2RateChangeData.InputData.Renewing_UniqueAndUnusual = (period.ZSLLine.UniqueUnusualInd) ? "Yes" : "No"

    return year2RateChangeData
  }

  private function GetRenewingDOData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    if (period.ZSLLine.ZSL_DandO_Cov_ZNA != null) {

      year2RateChangeData.InputData.Renewing_DO_YN = "Yes"

      if (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString != null) {
        year2RateChangeData.InputData.Renewing_DO_Limit = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Renewing_DO_Limit = (period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0 ) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      year2RateChangeData.InputData.Renewing_DO_LimitIsShared = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitIsShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitIsShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Renewing_DO_AttachmentPoint = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Renewing_DO_SIR = (period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA.doubleValue()

      var doCovPrices = period.ZSLLine.ZSL_DandO_Cov_ZNA.CoveragePrices
      doCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_DO_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Renewing_DO_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_DO_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Renewing_DO_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_DO)
      }

    }

    return year2RateChangeData
  }

  private function GetRenewingEPLIData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    if ( period.ZSLLine.ZSL_EPL_Cov_ZNA != null) {

      year2RateChangeData.InputData.Renewing_EPL_YN = "Yes"

      if(period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Renewing_EPL_Limit = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Renewing_EPL_Limit = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString.toDouble()
      }

      year2RateChangeData.InputData.Renewing_EPL_LimitIsShared = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZNA_EPL_LimitIsShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_EPL_Cov_ZNA.ZNA_EPL_LimitIsShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Renewing_EPL_AttachmentPoint = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Renewing_EPL_SIR = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString.toDouble()

      var eplCovPrice = period.ZSLLine.ZSL_EPL_Cov_ZNA.CoveragePrices.firstWhere(\elt -> elt.ZSPRatingPricingDetail.ZSLLineCostType != ZSLLineCovCostType_ZNA.TC_THIRD_PARTY_LIABILITY_EXCLUDED)
      year2RateChangeData.InputData.Renewing_EPL_BasePremium = (eplCovPrice.ZSPRatingPricingDetail.BasePremium == null) ? 0 : eplCovPrice.ZSPRatingPricingDetail.BasePremium.doubleValue()
      year2RateChangeData.InputData.Renewing_EPL_SharedLimitCredit = (eplCovPrice.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : eplCovPrice.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()

      //var zslLineCovCost = period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSLLineCovCosts.firstWhere(\elt -> !(elt typeis ZSLLineCovSubtypeCost_ZNA))
      //year1RateChangeData.InputData.Renewing_EPL_AnnualChargedPremium = (zslLineCovCost.ActualTermAmount_amt == null) ? 0 : zslLineCovCost.ActualTermAmount_amt.doubleValue()

      var zslLineCovCost = period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSLLineCovCosts.firstWhere(\elt -> !(elt typeis ZSLLineCovSubtypeCost_ZNA))
      year2RateChangeData.InputData.Renewing_EPL_AnnualChargedPremium += (zslLineCovCost.ActualTermAmount_amt == null) ? 0 : zslLineCovCost.ActualTermAmount_amt.doubleValue()

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Renewing_EPL_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_EPLI)
      }


      //Get EPL Rateable Employees Data
      if (period.ZSLLine.ZSLEPLIExposure_ZNA.Employees != null) {
        var sum : BigDecimal = 0
        var employees = period.ZSLLine.ZSLEPLIExposure_ZNA.Employees
        employees.each(\employee -> {
          var name = employee.EmployeeType.DisplayName
          var currentYear = (employee.CurrentYear < 1) ? 0 : employee.CurrentYear
          var ratableEmployeeFactor = (employee.RatableEmployeeFactor < 0) ? 0 : employee.RatableEmployeeFactor
          sum += (currentYear * ratableEmployeeFactor)
          switch (name) {
            case ZSLEmployeeType_ZNA.TC_FULLTIME.DisplayName:
              year2RateChangeData.InputData.Renewing_FullTimeEmployeesUS = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_PARTTIME.DisplayName:
              year2RateChangeData.InputData.Renewing_PartTimeEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_IND_CONTRACTERS.DisplayName:
              year2RateChangeData.InputData.Renewing_IndependentContractors = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_FOREIGN.DisplayName:
              year2RateChangeData.InputData.Renewing_ForeignEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_UNION.DisplayName:
              year2RateChangeData.InputData.Renewing_UnionEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_VOLUNTEERS.DisplayName:
              year2RateChangeData.InputData.Renewing_Volunteers = currentYear
              break
            default:
              break
          }
        })
        year2RateChangeData.InputData.Renewing_EPL_RatableEmployees = (sum?.setScale(0, BigDecimal.ROUND_HALF_UP) < 0) ? 0 : sum?.setScale(0, BigDecimal.ROUND_HALF_UP).doubleValue()
      }

      //Get EPL Employees Data
      year2RateChangeData.InputData.Renewing_FullTimeEmployeesForeign = (period.ZSLLine.ZSLEPLIExposure_ZNA.ForeignEmployees == null) ? 0 : period.ZSLLine.ZSLEPLIExposure_ZNA.ForeignEmployees.doubleValue()

      //Get EPL State Employee Counts
      if (period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos != null) {
        var states = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos?.where(\elt -> (elt.Jurisdiction != "CA" && elt.Jurisdiction != "All Other" && elt.Jurisdiction != null))
        states?.eachWithIndex(\state, indx -> {
          switch (indx) {
            case 0:
              year2RateChangeData.InputData.Renewing_State1 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Renewing_FullTimeEmployeesUSState1 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 1:
              year2RateChangeData.InputData.Renewing_State2 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Renewing_FullTimeEmployeesUSState2 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 2:
              year2RateChangeData.InputData.Renewing_State3 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Renewing_FullTimeEmployeesUSState3 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 3:
              year2RateChangeData.InputData.Renewing_State4 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Renewing_FullTimeEmployeesUSState4 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 4:
              year2RateChangeData.InputData.Renewing_State5 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Renewing_FullTimeEmployeesUSState5 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            default:
              break
          }
        })
        var stateCA = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos.firstWhere(\elt -> (elt.Jurisdiction == "CA"))
        year2RateChangeData.InputData.Renewing_FullTimeEmployeesUSStateCA = (stateCA.CurrentYear == null) ? 0 : stateCA.CurrentYear.doubleValue()

        var stateAllOther = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos.firstWhere(\elt -> (elt.Jurisdiction == "All Other"))
        year2RateChangeData.InputData.Renewing_FullTimeEmployeesAllOther = (stateAllOther.CurrentYear == null) ? 0 : stateAllOther.CurrentYear.doubleValue()
      }

    }

    return year2RateChangeData
  }

  private function GetRenewingFIDData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Get Fiduciary Data
    if (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA != null) {

      year2RateChangeData.InputData.Renewing_FID_YN = "Yes"

      if(period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Renewing_FID_Limit = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Renewing_FID_Limit = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      year2RateChangeData.InputData.Renewing_FID_LimitIsShared = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitisShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitisShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Renewing_FID_AttachmentPoint = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Renewing_FID_SIR = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm.ValueAsString.toDouble()

      var fidCovPrices = period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.CoveragePrices
      fidCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_FID_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Renewing_FID_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_FID_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Renewing_FID_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_FIDUCIARY)
      }

    }

    return year2RateChangeData
  }

  private function GetRenewingCrimeData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Get Crime Employee Theft Data
    if (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA != null) {

      year2RateChangeData.InputData.Renewing_CR_YN = "Yes"

      if(period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Renewing_CR_Limit = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Renewing_CR_Limit = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      //Get Employee Theft coverage data
      var crCovPrices = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.CoveragePrices
      crCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Renewing_CR_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
        year2RateChangeData.InputData.Renewing_CR_LocationFactor = (elt.ZSPRatingPricingDetail.LocationFactor == null) ? 0 : elt.ZSPRatingPricingDetail.LocationFactor.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_EmployeeTheftAnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Renewing_CR_EmployeeTheftAnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_CRIME)
      }

      year2RateChangeData.InputData.Renewing_CR_Deductible = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString.toDouble()

      //Crime Ratable Employees
      year2RateChangeData.InputData.Renewing_CR_RatableEmployees = (period.ZSLLine.ZSLCrimeCovPart_ZNA.RatableEmployeesCount.toDouble() < 0) ? 0 : period.ZSLLine.ZSLCrimeCovPart_ZNA.RatableEmployeesCount.toDouble()

      year2RateChangeData.InputData.Renewing_FidelityClassCode = (period.ZSLLine.ZSLCrimeCovPart_ZNA.ClassCode == null) ? "" : period.ZSLLine.ZSLCrimeCovPart_ZNA.ClassCode

      year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += year2RateChangeData.InputData.Renewing_CR_EmployeeTheftAnnualChargedPremium
    }

    //Get Total Crime Coverage premium

    if (period.ZSLLine.ZSL_ClientProperty_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_OnPremises_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_OnPremises_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_InTransit_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_InTransit_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNAExists) {
      var zslLineCondCosts = period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNA.ZSLLineCondCosts
      zslLineCondCosts?.each(\elt -> {
        year2RateChangeData.InputData.Renewing_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }

    return year2RateChangeData
  }

  private function GetRenewingQuestionSetData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    var periodLines = period.Lines
    periodLines?.each(\elt -> {
      var lineAnswers = elt.LineAnswers
      lineAnswers?.each(\elt2 -> {
        var answer = elt2.IntegerAnswer
        var codeID = elt2.Question.CodeIdentifier
        switch (codeID) {
          case "GITotalLocationsNumber_ZNA":
            year2RateChangeData.InputData.Renewing_CR_NumberOfLocations = (answer == null) ? 0 : answer.doubleValue()
            break
          default:
            break
        }
      })
    })

    return year2RateChangeData
  }

  private function GetRenewingMiscData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    var DO_LimitIsShared = (year2RateChangeData.InputData.Renewing_DO_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_DO_LimitIsShared
    var FID_LimitIsShared = (year2RateChangeData.InputData.Renewing_FID_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_FID_LimitIsShared
    var EPL_LimitIsShared = (year2RateChangeData.InputData.Renewing_EPL_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_EPL_LimitIsShared
    var DO_LimitofLiability = (year2RateChangeData.InputData.Renewing_DO_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_DO_Limit
    var FID_LimitofLiability = (year2RateChangeData.InputData.Renewing_FID_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_FID_Limit
    var EPL_LimitofLiability = (year2RateChangeData.InputData.Renewing_EPL_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_EPL_Limit
    var SharedAggregateLimit : double = 0
    var NumberOfSharedCovParts = 0
    var numberOfCovParts = 0

    if (year2RateChangeData.InputData.Renewing_DO_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Renewing_FID_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Renewing_EPL_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Renewing_CR_YN == "Yes") numberOfCovParts +=1

    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared != "true" ){
      year2RateChangeData.InputData.Renewing_SharedLimit = "D&O and EPL"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared != "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Renewing_SharedLimit = "D&O and Fiduciary"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared != "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Renewing_SharedLimit = "EPL & Fiduciary"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Renewing_SharedLimit = "D&O, EPL, and Fiduciary"
      NumberOfSharedCovParts = 3
    }
    if (DO_LimitIsShared != "true" && EPL_LimitIsShared != "true"  && FID_LimitIsShared != "true" ){
      year2RateChangeData.InputData.Renewing_SharedLimit = "Separate Limits"
    }
    if (numberOfCovParts == 1){
      year2RateChangeData.InputData.Renewing_SharedLimit = "Stand-Alone Coverage"
    }

    //if the limits are shared set the SharedAggregateLimit to the largest of the limits
    if (DO_LimitIsShared == "true" && (DO_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = DO_LimitofLiability
    }
    if (FID_LimitIsShared == "true" && (FID_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = FID_LimitofLiability
    }
    if (EPL_LimitIsShared == "true" && (EPL_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = EPL_LimitofLiability
    }

    year2RateChangeData.InputData.Renewing_NumberOfSharedCovParts = NumberOfSharedCovParts
    year2RateChangeData.InputData.Renewing_SharedAggregateLimit = SharedAggregateLimit





    return year2RateChangeData
  }

  private function GetRestatedPolicyData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Policy
    year2RateChangeData.InputData.Restated_PolicyNumber = (period.PolicyNumber != null) ?  period.PolicyNumber : period.Job.JobNumber

    year2RateChangeData.InputData.Restated_QuoteNumber = (period.Job.JobNumber != null) ?  period.Job.JobNumber : ""

    //Insured
    year2RateChangeData.InputData.Restated_Insured = (period.PrimaryInsuredName == null) ? "" : period.PrimaryInsuredName

    //EffectiveDate
    year2RateChangeData.InputData.Restated_EffectiveDate = ((period.PeriodStart == null) ? "" : period.PeriodStart.asMMddyyyy_ZNA) as String

    //ExpirationDate
    year2RateChangeData.InputData.Restated_ExpirationDate = ((period.PeriodEnd == null) ? "" : period.PeriodEnd.asMMddyyyy_ZNA) as String

    //State
    year2RateChangeData.InputData.Restated_DomiciledState = (period.BaseState.Code == null) ? "" : period.BaseState.Code

    //CompanyType
    year2RateChangeData.InputData.Restated_Private_NonProfit =  period.Policy.Account.OtherOrgTypeDescription

    //Commission
    year2RateChangeData.InputData.Restated_Commission = (period.CommissionPercent_ZNA == null) ? 0 : period.CommissionPercent_ZNA.doubleValue()

    //NAICS
    year2RateChangeData.InputData.Restated_NAICSCode = (period.PrimaryNamedInsured.NAICS_ZNA.Code == null) ? "" : period.PrimaryNamedInsured.NAICS_ZNA.Code

    //NAICS Description
    year2RateChangeData.InputData.Restated_NAICSDescription = (period.PrimaryNamedInsured.NAICS_ZNA.Classification == null) ? " " : period.PrimaryNamedInsured.NAICS_ZNA.Classification

    //Industry Type
    year2RateChangeData.InputData.Restated_IndustryType = (period.ZSLLine.ZSLIndustryType.DisplayName == null) ? "" : period.ZSLLine.ZSLIndustryType.DisplayName.toString()
    year2RateChangeData.InputData.Restated_IndustryTypeCode = (period.ZSLLine.ZSLIndustryType.Code == null) ? "" : period.ZSLLine.ZSLIndustryType.Code

    //Premium
    year2RateChangeData.InputData.Restated_ActualCharged = (period.TotalPremiumRPT_ZNA == null) ? 0 : period.TotalPremiumRPT_ZNA.doubleValue()

    //Asset Size
    year2RateChangeData.InputData.Restated_TotalAssets = (period.ZSLLine.CurrentYearFinancial.TotalAssets.Amount == null) ? 0 : period.ZSLLine.CurrentYearFinancial.TotalAssets.Amount.doubleValue()

    //Plan Asset
    year2RateChangeData.InputData.Restated_TotalPlanAssets = (period.ZSLLine.ZSLFiduciaryCovPart_ZNA.TotalPlanAssets == null) ? 0 : period.ZSLLine.ZSLFiduciaryCovPart_ZNA.TotalPlanAssets.doubleValue()

    //Policy Type - Primary or Excess
    year2RateChangeData.InputData.Restated_Primary_Excess = (period.ZSLLine.PolicyType_ZNA == null) ? "" : period.ZSLLine.PolicyType_ZNA.toString()

    year2RateChangeData.InputData.Restated_UniqueAndUnusual = (period.ZSLLine.UniqueUnusualInd) ? "Yes" : "No"


    return year2RateChangeData
  }

  private function GetRestatedDOData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    if (period.ZSLLine.ZSL_DandO_Cov_ZNA != null) {

      year2RateChangeData.InputData.Restated_DO_YN = "Yes"

      if (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString != null) {
        year2RateChangeData.InputData.Restated_DO_Limit = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Restated_DO_Limit = (period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0 ) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      year2RateChangeData.InputData.Restated_DO_LimitIsShared = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitIsShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitIsShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Restated_DO_AttachmentPoint = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Restated_DO_SIR = (period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA.doubleValue()

      var doCovPrices = period.ZSLLine.ZSL_DandO_Cov_ZNA.CoveragePrices
      doCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Restated_DO_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Restated_DO_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_DO_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Restated_DO_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_DO)
      }

    }

    return year2RateChangeData
  }

  private function GetRestatedEPLIData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    if ( period.ZSLLine.ZSL_EPL_Cov_ZNA != null) {

      year2RateChangeData.InputData.Restated_EPL_YN = "Yes"

      if(period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Restated_EPL_Limit = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Restated_EPL_Limit = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString.toDouble()
      }

      year2RateChangeData.InputData.Restated_EPL_LimitIsShared = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZNA_EPL_LimitIsShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_EPL_Cov_ZNA.ZNA_EPL_LimitIsShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Restated_EPL_AttachmentPoint = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Restated_EPL_SIR = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString.toDouble()

      var eplCovPrice = period.ZSLLine.ZSL_EPL_Cov_ZNA.CoveragePrices.firstWhere(\elt -> elt.ZSPRatingPricingDetail.ZSLLineCostType != ZSLLineCovCostType_ZNA.TC_THIRD_PARTY_LIABILITY_EXCLUDED)
      year2RateChangeData.InputData.Restated_EPL_BasePremium = (eplCovPrice.ZSPRatingPricingDetail.BasePremium == null) ? 0 : eplCovPrice.ZSPRatingPricingDetail.BasePremium.doubleValue()
      year2RateChangeData.InputData.Restated_EPL_SharedLimitCredit = (eplCovPrice.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : eplCovPrice.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()

      //var zslLineCovCost = period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSLLineCovCosts.firstWhere(\elt -> !(elt typeis ZSLLineCovSubtypeCost_ZNA))
      //year1RateChangeData.InputData.Renewing_EPL_AnnualChargedPremium = (zslLineCovCost.ActualTermAmount_amt == null) ? 0 : zslLineCovCost.ActualTermAmount_amt.doubleValue()

      var zslLineCovCost = period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSLLineCovCosts.firstWhere(\elt -> !(elt typeis ZSLLineCovSubtypeCost_ZNA))
      year2RateChangeData.InputData.Restated_EPL_AnnualChargedPremium += (zslLineCovCost.ActualTermAmount_amt == null) ? 0 : zslLineCovCost.ActualTermAmount_amt.doubleValue()

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Restated_EPL_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_EPLI)
      }


      //Get EPL Rateable Employees Data
      if (period.ZSLLine.ZSLEPLIExposure_ZNA.Employees != null) {
        var sum : BigDecimal = 0
        var employees = period.ZSLLine.ZSLEPLIExposure_ZNA.Employees
        employees.each(\employee -> {
          var name = employee.EmployeeType.DisplayName
          var currentYear = (employee.CurrentYear < 1) ? 0 : employee.CurrentYear
          var ratableEmployeeFactor = (employee.RatableEmployeeFactor < 0) ? 0 : employee.RatableEmployeeFactor
          sum += (currentYear * ratableEmployeeFactor)
          switch (name) {
            case ZSLEmployeeType_ZNA.TC_FULLTIME.DisplayName:
              year2RateChangeData.InputData.Restated_FullTimeEmployeesUS = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_PARTTIME.DisplayName:
              year2RateChangeData.InputData.Restated_PartTimeEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_IND_CONTRACTERS.DisplayName:
              year2RateChangeData.InputData.Restated_IndependentContractors = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_FOREIGN.DisplayName:
              year2RateChangeData.InputData.Restated_ForeignEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_UNION.DisplayName:
              year2RateChangeData.InputData.Restated_UnionEmployees = currentYear
              break
            case ZSLEmployeeType_ZNA.TC_VOLUNTEERS.DisplayName:
              year2RateChangeData.InputData.Exipiring_Volunteers = currentYear
              break
            default:
              break
          }
        })
        year2RateChangeData.InputData.Restated_EPL_RateableEmployees = (sum?.setScale(0, BigDecimal.ROUND_HALF_UP) < 0) ? 0 : sum?.setScale(0, BigDecimal.ROUND_HALF_UP).doubleValue()
      }

      //Get EPL Employees Data
      year2RateChangeData.InputData.Restated_FullTimeEmployeesForeign = (period.ZSLLine.ZSLEPLIExposure_ZNA.ForeignEmployees == null) ? 0 : period.ZSLLine.ZSLEPLIExposure_ZNA.ForeignEmployees.doubleValue()

      //Get EPL State Employee Counts
      if (period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos != null) {
        var states = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos?.where(\elt -> !(elt.Jurisdiction == "CA" || elt.Jurisdiction == "All Other" || elt.Jurisdiction == null))
        states?.eachWithIndex(\state, indx -> {
          switch (indx) {
            case 0:
              year2RateChangeData.InputData.Restated_State1 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Restated_FullTimeEmployeesUSState1 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 1:
              year2RateChangeData.InputData.Restated_State2 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Restated_FullTimeEmployeesUSState2 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 2:
              year2RateChangeData.InputData.Restated_State3 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Restated_FullTimeEmployeesUSState3 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 3:
              year2RateChangeData.InputData.Restated_State4 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Restated_FullTimeEmployeesUSState4 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            case 4:
              year2RateChangeData.InputData.Restated_State5 = (state.Jurisdiction == null) ? " " : state.Jurisdiction
              year2RateChangeData.InputData.Restated_FullTimeEmployeesUSState5 = (state.CurrentYear == null) ? 0 : state.CurrentYear.doubleValue()
              break
            default:
              break
          }
        })
        var stateCA = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos.firstWhere(\elt -> (elt.Jurisdiction == "CA"))
        year2RateChangeData.InputData.Restated_FullTimeEmployeesUSStateCA = (stateCA.CurrentYear == null) ? 0 : stateCA.CurrentYear.doubleValue()

        var stateAllOther = period.ZSLLine.ZSLEPLIExposure_ZNA.JurisdictionInfos.firstWhere(\elt -> (elt.Jurisdiction == "All Other"))
        year2RateChangeData.InputData.Restated_FullTimeEmployeesAllOther = (stateAllOther.CurrentYear == null) ? 0 : stateAllOther.CurrentYear.doubleValue()
      }

    }

    return year2RateChangeData
  }

  private function GetRestatedFIDData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Get Fiduciary Data
    if (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA != null) {

      year2RateChangeData.InputData.Restated_FID_YN = "Yes"

      if(period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Restated_FID_Limit = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Restated_FID_Limit = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      year2RateChangeData.InputData.Restated_FID_LimitIsShared = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitisShared_ZNATerm.ValueAsString != null) ? period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitisShared_ZNATerm.ValueAsString : ""

      year2RateChangeData.InputData.Restated_FID_AttachmentPoint = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_AttachmentPoint_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_AttachmentPoint_ZNATerm.ValueAsString.toDouble()

      year2RateChangeData.InputData.Restated_FID_SIR = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm.ValueAsString.toDouble()

      var fidCovPrices = period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.CoveragePrices
      fidCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Restated_FID_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Restated_FID_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_FID_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Restated_FID_AnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_FIDUCIARY)
      }

    }

    return year2RateChangeData
  }

  private function GetRestatedCrimeData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {
    //Get Crime Employee Theft Data
    if (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA != null) {

      year2RateChangeData.InputData.Restated_CR_YN = "Yes"

      if(period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString != null){
        year2RateChangeData.InputData.Restated_CR_Limit = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_LimitofLiabilityOther_ZNATerm.ValueAsString.toDouble()
      }
      else {
        year2RateChangeData.InputData.Restated_CR_Limit = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA.doubleValue()
      }

      //Get Employee Theft coverage data
      var crCovPrices = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.CoveragePrices
      crCovPrices?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_BasePremium = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium.doubleValue()
        year2RateChangeData.InputData.Restated_CR_SharedLimitCredit = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit.doubleValue()
        year2RateChangeData.InputData.Restated_CR_LocationFactor = (elt.ZSPRatingPricingDetail.LocationFactor == null) ? 0 : elt.ZSPRatingPricingDetail.LocationFactor.doubleValue()
      })

      var zslLineCovCosts = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_EmployeeTheftAnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })

      //Get Excess premium
      if (period.PolicyType_ZNA == PolicyType_ZNA.TC_EXCESS) {
        year2RateChangeData.InputData.Restated_CR_EmployeeTheftAnnualChargedPremium = getExcessPremiumRatio(period, CoveragePart_ZNA.TC_CRIME)
      }

      year2RateChangeData.InputData.Restated_CR_Deductible = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString.toDouble()

      //Crime Ratable Employees
      year2RateChangeData.InputData.Restated_CR_RateableEmployees = (period.ZSLLine.ZSLCrimeCovPart_ZNA.RatableEmployeesCount.toDouble() < 0) ? 0 : period.ZSLLine.ZSLCrimeCovPart_ZNA.RatableEmployeesCount.toDouble()

      year2RateChangeData.InputData.Restated_FidelityClassCode = (period.ZSLLine.ZSLCrimeCovPart_ZNA.ClassCode == null) ? "" : period.ZSLLine.ZSLCrimeCovPart_ZNA.ClassCode

      year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += year2RateChangeData.InputData.Renewing_CR_EmployeeTheftAnnualChargedPremium
    }

    //Get Total Crime Coverage premium

    if (period.ZSLLine.ZSL_ClientProperty_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_OnPremises_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_OnPremises_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_InTransit_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_InTransit_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNAExists) {
      var zslLineCovCosts = period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }
    if (period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNAExists) {
      var zslLineCondCosts = period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNA.ZSLLineCondCosts
      zslLineCondCosts?.each(\elt -> {
        year2RateChangeData.InputData.Restated_CR_AnnualChargedPremium += (elt.ActualTermAmount_amt == null) ? 0 : elt.ActualTermAmount_amt.doubleValue()
      })
    }

    return year2RateChangeData
  }

  private function GetRestatedQuestionSetData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    var periodLines = period.Lines
    periodLines?.each(\elt -> {
      var lineAnswers = elt.LineAnswers
      lineAnswers?.each(\elt2 -> {
        var answer = elt2.IntegerAnswer
        var codeID = elt2.Question.CodeIdentifier
        switch (codeID) {
          case "GITotalLocationsNumber_ZNA":
            year2RateChangeData.InputData.Restated_CR_NumberOfLocations = (answer == null) ? 0 : answer.doubleValue()
            break
          default:
            break
        }
      })
    })

    return year2RateChangeData
  }

  private function GetRestatedMiscData(year2RateChangeData:Year2RateChangeData, period:PolicyPeriod) : Year2RateChangeData {

    var DO_LimitIsShared = (year2RateChangeData.InputData.Restated_DO_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_DO_LimitIsShared
    var FID_LimitIsShared = (year2RateChangeData.InputData.Restated_FID_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_FID_LimitIsShared
    var EPL_LimitIsShared = (year2RateChangeData.InputData.Restated_EPL_LimitIsShared == null) ? "false" : year2RateChangeData.InputData.Renewing_EPL_LimitIsShared
    var DO_LimitofLiability = (year2RateChangeData.InputData.Restated_DO_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_DO_Limit
    var FID_LimitofLiability = (year2RateChangeData.InputData.Restated_FID_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_FID_Limit
    var EPL_LimitofLiability = (year2RateChangeData.InputData.Restated_EPL_Limit < 0 ) ? 0 : year2RateChangeData.InputData.Renewing_EPL_Limit
    var SharedAggregateLimit : double = 0
    var NumberOfSharedCovParts = 0
    var numberOfCovParts = 0

    if (year2RateChangeData.InputData.Restated_DO_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Restated_FID_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Restated_EPL_YN == "Yes") numberOfCovParts +=1
    if (year2RateChangeData.InputData.Restated_CR_YN == "Yes") numberOfCovParts +=1

    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared != "true" ){
      year2RateChangeData.InputData.Restated_SharedLimit = "D&O and EPL"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared != "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Restated_SharedLimit = "D&O and Fiduciary"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared != "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Restated_SharedLimit = "EPL & Fiduciary"
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "true" ){
      year2RateChangeData.InputData.Restated_SharedLimit = "D&O, EPL, and Fiduciary"
      NumberOfSharedCovParts = 3
    }
    if (DO_LimitIsShared != "true" && EPL_LimitIsShared != "true"  && FID_LimitIsShared != "true" ){
      year2RateChangeData.InputData.Restated_SharedLimit = "Separate Limits"
    }
    if (numberOfCovParts == 1){
      year2RateChangeData.InputData.Restated_SharedLimit = "Stand-Alone Coverage"
    }

    //if the limits are shared set the SharedAggregateLimit to the largest of the limits
    if (DO_LimitIsShared == "true" && (DO_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = DO_LimitofLiability
    }
    if (FID_LimitIsShared == "true" && (FID_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = FID_LimitofLiability
    }
    if (EPL_LimitIsShared == "true" && (EPL_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = EPL_LimitofLiability
    }

    year2RateChangeData.InputData.Restated_NumberOfSharedCovParts = NumberOfSharedCovParts
    year2RateChangeData.InputData.Restated_SharedAggregateLimit = SharedAggregateLimit





    return year2RateChangeData
  }

  private function CreateYear2RateChangeExport(year2RateChangeData : Year2RateChangeData) : ByteArrayOutputStream {
    var workbook : XSSFWorkbook
    var worksheet : XSSFSheet

    //get workbook template file stream
    var fileInputStream = new FileInputStream(ConfigAccess.getConfigFile(templateFilePath))

    //read template filestream into workbook
    workbook = new XSSFWorkbook(fileInputStream)

    //populate the Data worksheet in the workbook
    workbook = populateWorkbookDataSheet(workbook, year2RateChangeData)

    //evaluate formulas in workbook (executes formulas and updates references)
    var evaluator = workbook.getCreationHelper().createFormulaEvaluator()
    evaluator.evaluateAll()

    //populate the output dto from input data and calculated rate change data from workbook
    GetOutput(workbook, year2RateChangeData)

    //create byte array output stream
    var baos = new ByteArrayOutputStream()
    workbook.write(baos)

    //close workbook
    workbook.close()

    return baos
  }

  private function populateWorkbookDataSheet(workbook : XSSFWorkbook, year2RateChangeData : Year2RateChangeData) : XSSFWorkbook {
    var worksheet : XSSFSheet
    var createHelper = workbook.getCreationHelper();

    styleYellowHeader = workbook.createCellStyle()
    styleYellowHeader.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex())
    styleYellowHeader.setFillPattern(SOLID_FOREGROUND)
    styleYellowHeader.setBorderTop(BorderStyle.THIN)
    styleYellowHeader.setBorderBottom(BorderStyle.THIN)
    styleYellowHeader.setBorderLeft(BorderStyle.THIN)
    styleYellowHeader.setBorderRight(BorderStyle.THIN)

    //create the Data sheet in the workbook
    worksheet = workbook.getSheet("SystemData")

    var rowIndx = 0
    worksheet = SetWorkbookExpiringData(worksheet, year2RateChangeData, rowIndx)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetWorkbookRenewingData(worksheet, year2RateChangeData, rowIndx)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetWorkbookRestatedData(worksheet, year2RateChangeData, rowIndx)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++


    //adjust column sizes
    worksheet.autoSizeColumn(0)
    worksheet.autoSizeColumn(1)


    return workbook
  }

  private function SetEnhancement(worksheet : XSSFSheet, period : PolicyPeriod, rowIndx : int, isRenewingTerm : boolean) : XSSFSheet {
    // add a blank row
    rowIndx++

    //create Header Row
    var row = worksheet.createRow(rowIndx)
    var cellHA = row.createCell(0)
    cellHA.setCellValue("Enhancement")
    cellHA.setCellStyle(styleYellowHeader)

    rowIndx++

    row = worksheet.createRow(rowIndx)
    cellHA = row.createCell(0)
    cellHA.setCellValue("Description")
    cellHA.setCellStyle(styleYellowHeader)

    var cellHB = row.createCell(1)
    cellHB.setCellValue("Value")
    cellHB.setCellStyle(styleYellowHeader)

    rowIndx++

    var startingRowIndex = rowIndx
    var label = ""
    var dValue : double
    var strValue = ""

    //Get D&O Data
    if (period.ZSLLine.DOCoveragePartExists ) {
      label = "D&O Additional Defense Limit"
      dValue = (period.ZSLLine.ZSL_AdditionalLimitofliabilityforDefenseCosts_Cov_ZNA.ZSL_AdditionalLimitofliabilityforDefenseCosts_Limit_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_AdditionalLimitofliabilityforDefenseCosts_Cov_ZNA.ZSL_AdditionalLimitofliabilityforDefenseCosts_Limit_ZNATerm.ValueAsString?.toDouble()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++

      var doCovPrices = period.ZSLLine.ZSL_DandO_Cov_ZNA.CoveragePrices
      doCovPrices?.each(\elt -> {
        label = "D&O Base Rate Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O ILF Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.ILFAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.ILFAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O ILF Factor"
        dValue = (elt.ILFFactor == null) ? 1 : elt.ILFFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O Industry Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O Shared Limit Credit"
        dValue = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O State Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.StateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.StateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O Commission Adjustment Factor"
        dValue = (elt.CommissionAdjustmentFactor == null) ? 1 : elt.CommissionAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O Additional Defense Limit Factor"
        dValue = (elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O Additional Executive Limit of Liability Factor"
        dValue = (elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "D&O Minimum Premium"
        dValue = (elt.MinimumPremium == null) ? 1 : elt.MinimumPremium?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
      })

      /*
      label = "D&O Additional Defense Limit Selected"
      dValue = (period.ZSLLine.? == null) ? 0 : period.ZSLLine.?.doubleValue()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++

      label = "D&O Additional Executive Limit of Liability"
      dValue = (period.ZSLLine.? == null) ? 0 : period.ZSLLine.?.doubleValue()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++

      label = "D&O Additional Executive Limit of Liability Selected"
      dValue = (period.ZSLLine.? == null) ? 0 : period.ZSLLine.?.doubleValue()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++
       */


    }
    if (period.ZSLLine.EPLICoveragePartExists ) {

      //Get coverage price for the base EPL coverage.  And not for other EPL Coverage types like "Removal of Third party liability"
      var eplCovPrices = period.ZSLLine.ZSL_EPL_Cov_ZNA.CoveragePrices.where(\elt -> elt.ZSPRatingPricingDetail.ZSLLineCostType.DisplayName == null)
      eplCovPrices?.each(\elt -> {
        label = "EPL State Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.StateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.StateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Base Rate Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL ILF Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.ILFAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.ILFAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL ILF Factor"
        dValue = (elt.ILFFactor == null) ? 1 : elt.ILFFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Industry Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Shared Limit Credit"
        dValue = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Commission Adjustment Factor"
        dValue = (elt.CommissionAdjustmentFactor == null) ? 1 : elt.CommissionAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Additional Defense Limit Factor"
        dValue = (elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Additional Executive Limit of Liability Factor"
        dValue = (elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Third Party Limit Retention Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.ThirdPartyLmtRetentionAdj == null) ? 1 : elt.ZSPRatingPricingDetail.ThirdPartyLmtRetentionAdj?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 1 Factor"
        dValue = (elt.ZSPRatingPricingDetail.State1Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State1Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 2 Factor"
        dValue = (elt.ZSPRatingPricingDetail.State2Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State2Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 3 Factor"
        dValue = (elt.ZSPRatingPricingDetail.State3Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State3Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 4 Factor"
        dValue = (elt.ZSPRatingPricingDetail.State4Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State4Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 5 Factor"
        dValue = (elt.ZSPRatingPricingDetail.State5Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State5Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL California Factor"
        dValue = (elt.ZSPRatingPricingDetail.CaliforniaFactor == null) ? 1 : elt.ZSPRatingPricingDetail.CaliforniaFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Countrywide Factor"
        dValue = (elt.ZSPRatingPricingDetail.CountrywideFactor == null) ? 1 : elt.ZSPRatingPricingDetail.CountrywideFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Minimum Premium"
        dValue = (elt.MinimumPremium == null) ? 1 : elt.MinimumPremium?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        })

      //Get coverage price for other EPL Coverage types like "Removal of Third party liability"
      var eplCovPricesCostType = period.ZSLLine.ZSL_EPL_Cov_ZNA.CoveragePrices.where(\elt -> elt.ZSPRatingPricingDetail.ZSLLineCostType.DisplayName != null)
      eplCovPricesCostType?.each(\elt -> {
        var costTypeName = elt.ZSPRatingPricingDetail.ZSLLineCostType.DisplayName
        label = "EPL State Adjustment Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.StateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.StateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Base Rate Adjustment Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL ILF Adjustment Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.ILFAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.ILFAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL ILF Factor" + " " + costTypeName
        dValue = (elt.ILFFactor == null) ? 1 : elt.ILFFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Industry Adjustment Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Shared Limit Credit" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Commission Adjustment Factor" + " " + costTypeName
        dValue = (elt.CommissionAdjustmentFactor == null) ? 1 : elt.CommissionAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Additional Defense Limit Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Additional Executive Limit of Liability Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Third Party Limit Retention Adjustment Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.ThirdPartyLmtRetentionAdj == null) ? 1 : elt.ZSPRatingPricingDetail.ThirdPartyLmtRetentionAdj?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 1 Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.State1Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State1Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 2 Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.State2Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State2Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 3 Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.State3Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State3Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 4 Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.State4Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State4Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL State 5 Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.State5Factor == null) ? 1 : elt.ZSPRatingPricingDetail.State5Factor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL California Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.CaliforniaFactor == null) ? 1 : elt.ZSPRatingPricingDetail.CaliforniaFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Countrywide Factor" + " " + costTypeName
        dValue = (elt.ZSPRatingPricingDetail.CountrywideFactor == null) ? 1 : elt.ZSPRatingPricingDetail.CountrywideFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "EPL Minimum Premium" + " " + costTypeName
        dValue = (elt.MinimumPremium == null) ? 1 : elt.MinimumPremium?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
      })

      label = "EPL Additional Defense Limit"
      dValue = (period.ZSLLine.ZSL_AdditionalLimitofliabilityforDefenseCosts_Cov_ZNA.ZSL_AdditionalLimitofliabilityforDefenseCosts_Limit_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_AdditionalLimitofliabilityforDefenseCosts_Cov_ZNA.ZSL_AdditionalLimitofliabilityforDefenseCosts_Limit_ZNATerm.ValueAsString?.toDouble()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++
      label = "EPL Third Party Discrimilation Claims Sublimit"
      dValue = (period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_ThirdPtyDistnClmSubLt_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_ThirdPtyDistnClmSubLt_ZNATerm.ValueAsString?.toDouble()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++
      label = "EPL Third Party Limit of Liability - Other"
      dValue = (period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_LimitLiabilityOther_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_LimitLiabilityOther_ZNATerm.ValueAsString?.toDouble()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++
      label = "EPL Retention for each Third Party Discrimination Claim"
      dValue = (period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_RetentionPtyDisnClm_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_RetentionPtyDisnClm_ZNATerm.ValueAsString?.toDouble()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++
      /*label = "EPL Pending or Prior Date for Third Party Discrimination"
      dValue = (period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_PdgPrirThirdPtyDisctnclm_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_EPLThirdParty_Cov_ZNA.ZSL_EPLThirdParty_PdgPrirThirdPtyDisctnclm_ZNATerm.ValueAsString
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++*/
    }

    if (period.ZSLLine.FiduciaryCoveragePartExists) {

      //Get FID Data
      var fidCovPrices = period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.CoveragePrices
      fidCovPrices?.each(\elt -> {
        label = "FID Base Rate Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID ILF Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.ILFAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.ILFAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID ILF Factor"
        dValue = (elt.ILFFactor == null) ? 1 : elt.ILFFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID Industry Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID Shared Limit Credit"
        dValue = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID State Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.StateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.StateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID Commission Adjustment Factor"
        dValue = (elt.CommissionAdjustmentFactor == null) ? 1 : elt.CommissionAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID Additional Defense Limit Factor"
        dValue = (elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalDefenseLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID Additional Executive Limit of Liability Factor"
        dValue = (elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor == null) ? 1 : elt.ZSPRatingPricingDetail.AdditionalExecutiveLimitFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "FID Minimum Premium"
        dValue = (elt.MinimumPremium == null) ? 1 : elt.MinimumPremium?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
      })
      label = "FID Additional Defense Limit"
      dValue = (period.ZSLLine.ZSL_AdditionalLimitofliabilityforDefenseCosts_Cov_ZNA.ZSL_AdditionalLimitofliabilityforDefenseCosts_Limit_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_AdditionalLimitofliabilityforDefenseCosts_Cov_ZNA.ZSL_AdditionalLimitofliabilityforDefenseCosts_Limit_ZNATerm.ValueAsString?.toDouble()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++
    }

    if (period.ZSLLine.CrimeCoveragePartExists ) {

      var crCovPrices = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.CoveragePrices
      crCovPrices?.each(\elt -> {
        label = "CR Base Rate Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR ILF Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.ILFAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.ILFAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR ILF Factor"
        dValue = (elt.ILFFactor == null) ? 1 : elt.ILFFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Industry Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR State Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.StateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.StateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Shared Limit Credit"
        dValue = (elt.ZSPRatingPricingDetail.SharedLimitCredit == null) ? 0 : elt.ZSPRatingPricingDetail.SharedLimitCredit?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Commission Adjustment Factor"
        dValue = (elt.CommissionAdjustmentFactor == null) ? 1 : elt.CommissionAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Minimum Premium"
        dValue = (elt.MinimumPremium == null) ? 1 : elt.MinimumPremium?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
      })


      worksheet = SetCrimeData(worksheet,period,rowIndx)
      rowIndx = worksheet.LastRowNum
      rowIndx++

      var employees = gw.lob.zsl.ui.ZSLCrime_ZNAScreenHelper.getEmployees(period)
      employees?.each(\emp -> {
        label = "CR " + emp.EmployeeType
        dValue = (emp.CurrentYear == null) ? 0 : emp.CurrentYear?.intValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
      })
    }

    var strPrefix : String = ""
    var intRowCounter = startingRowIndex
    strPrefix = isRenewingTerm ? "Renewing_":"Expiring_"

    if (period.PolicyPeriodType_ZNA == PolicyPeriodType_ZNA.TC_RESTATED){
      strPrefix = "Restated_"
    }

    while(intRowCounter<rowIndx){
      if (worksheet.getRow(intRowCounter).getCell(0).getStringCellValue() != ""){
        worksheet.getRow(intRowCounter).getCell(0).setCellValue(strPrefix + worksheet.getRow(intRowCounter).getCell(0).getStringCellValue())
      }
      intRowCounter++
    }

    worksheet.autoSizeColumn(0)
    worksheet.setColumnWidth(1, 30*256)
    worksheet.setColumnWidth(2, 20*256)
    worksheet.setColumnWidth(3, 20*256)
    worksheet.setColumnWidth(4, 20*256)
    worksheet.setColumnWidth(5, 20*256)
    worksheet.setColumnWidth(6, 25*256)
    return worksheet
  }


  private function SetCrimeData(worksheet : XSSFSheet, period : PolicyPeriod, rowIndx : int) : XSSFSheet {
    var label = ""
    var dValue : double
    var strValue = ""

    if (period.ZSLLine.CrimeCoveragePartExists ) {

      label = "CR Coverage Selected"
      strValue = "Yes"
      createRowWith2Cells(worksheet, rowIndx, label, strValue)
      rowIndx++

      //Get Crime Coverage Data
      if (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA != null) {
        label = "CR EmployeeTheft Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR EmployeeTheft Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        strValue = (period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString == null) ? "" : period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSL_EmployeeTheft_Deductible_ZNATerm.ValueAsString
        label = "CR EmployeeTheft Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, strValue?.toDouble())
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR EmployeeTheft Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var employeeTheftCovPrices = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.CoveragePrices
        employeeTheftCovPrices?.each(\elt -> {
          label = "CR EmployeeTheft Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_ClientProperty_Cov_ZNAExists){
        label = "CR ClientProperty Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR ClientProperty Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.ZSL_ClientProperty_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.ZSL_ClientProperty_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR ClientProperty Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR ClientProperty Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var clientPropertyCovPrices = period.ZSLLine.ZSL_ClientProperty_Cov_ZNA.CoveragePrices
        clientPropertyCovPrices?.each(\elt -> {
          label = "CR ClientProperty Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNAExists){
        label = "CR ForgeryorAlterationChecksForgery Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR ForgeryorAlterationChecksForgery Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.ZSL_ForgeryorAlterationChecksForgery_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.ZSL_ForgeryorAlterationChecksForgery_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR ForgeryorAlterationChecksForgery Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR ForgeryorAlterationChecksForgery Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var forgeryorAlterationChecksForgeryCovPrices = period.ZSLLine.ZSL_ForgeryorAlterationChecksForgery_Cov_ZNA.CoveragePrices
        forgeryorAlterationChecksForgeryCovPrices?.each(\elt -> {
          label = "CR ForgeryorAlterationChecksForgery Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNAExists){
        label = "CR ForgeryorAltrationCreditDebitorChrgeCrdFrgry Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR ForgeryorAltrationCreditDebitorChrgeCrdFrgry Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR ForgeryorAltrationCreditDebitorChrgeCrdFrgry Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR ForgeryorAltrationCreditDebitorChrgeCrdFrgry Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var forgeryorAltrationCreditDebitorChrgeCrdFrgryCovPrices = period.ZSLLine.ZSL_ForgeryorAltrationCreditDebitorChrgeCrdFrgry_Cov_ZNA.CoveragePrices
        forgeryorAltrationCreditDebitorChrgeCrdFrgryCovPrices?.each(\elt -> {
          label = "CR ForgeryorAltrationCreditDebitorChrgeCrdFrgry Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_OnPremises_Cov_ZNAExists){
        label = "CR OnPremises Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_OnPremises_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_OnPremises_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR OnPremises Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_OnPremises_Cov_ZNA.ZSL_OnPremises_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_OnPremises_Cov_ZNA.ZSL_OnPremises_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR OnPremises Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_OnPremises_Cov_ZNA.ZSLLineCovCosts.sum(\elt -> elt.ActualTermAmount_amt)
        dValue = termAmt?.doubleValue()
        label = "CR OnPremises Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var onPremisesCovPrices = period.ZSLLine.ZSL_OnPremises_Cov_ZNA.CoveragePrices
        onPremisesCovPrices?.eachWithIndex(\elt, indx -> {
          label = "CR OnPremises Base Premium " + indx as String
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          createRowWith2Cells(worksheet, rowIndx, label, dValue)
          rowIndx++
        })
      }

      if (period.ZSLLine.ZSL_InTransit_Cov_ZNAExists){
        label = "CR InTransit Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_InTransit_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_InTransit_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR InTransit Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_InTransit_Cov_ZNA.ZSL_InTransit_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_InTransit_Cov_ZNA.ZSL_InTransit_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR InTransit Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_InTransit_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR InTransit Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var inTransitCovPrices = period.ZSLLine.ZSL_InTransit_Cov_ZNA.CoveragePrices
        inTransitCovPrices?.each(\elt -> {
          label = "CR InTransit Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNAExists){
        label = "CR ComputerFraudandFundsTransferFraud Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR ComputerFraudandFundsTransferFraud Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.ZSL_ComputerFraudandFundsTransferFraud_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.ZSL_ComputerFraudandFundsTransferFraud_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR ComputerFraudandFundsTransferFraud Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR ComputerFraudandFundsTransferFraud Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var computerFraudandFundsTransferFraudCovPrices = period.ZSLLine.ZSL_ComputerFraudandFundsTransferFraud_Cov_ZNA.CoveragePrices
        computerFraudandFundsTransferFraudCovPrices?.each(\elt -> {
          label = "CR ComputerFraudandFundsTransferFraud Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNAExists){
        label = "CR FraudulentImpersonation Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSL_FraudulentImpersonation_LimitofLiability_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSL_FraudulentImpersonation_LimitofLiability_ZNATerm.ValueAsString?.toDouble()
        label = "CR FraudulentImpersonation Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSL_FraudulentImpersonation_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSL_FraudulentImpersonation_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR FraudulentImpersonation Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        strValue = (period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSL_FraudulentImpersonation_Verification_ZNATerm.ValueAsString_ZNA == null) ? "" : period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSL_FraudulentImpersonation_Verification_ZNATerm.ValueAsString_ZNA
        label = "CR FraudulentImpersonation Verification"
        createRowWith2Cells(worksheet, rowIndx, label, strValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR FraudulentImpersonation Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var fraudulentImpersonationCovPrices = period.ZSLLine.ZSL_FraudulentImpersonation_Cov_ZNA.CoveragePrices
        fraudulentImpersonationCovPrices?.each(\elt -> {
          label = "CR FraudulentImpersonation Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNAExists){
        label = "CR MoneyOrdersandCounterfeitMoney Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR MoneyOrdersandCounterfeitMoney Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.ZSL_MoneyOrdersandCounterfeitMoney_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.ZSL_MoneyOrdersandCounterfeitMoney_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR MoneyOrdersandCounterfeitMoney Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR MoneyOrdersandCounterfeitMoney Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var moneyOrdersandCounterfeitMoneyCovPrices = period.ZSLLine.ZSL_MoneyOrdersandCounterfeitMoney_Cov_ZNA.CoveragePrices
        moneyOrdersandCounterfeitMoneyCovPrices?.each(\elt -> {
          label = "CR MoneyOrdersandCounterfeitMoney Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNAExists){
        label = "CR ElectronicDataorComputerProgramsRestorationCosts Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.LiabilityLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
        label = "CR ElectronicDataorComputerProgramsRestorationCosts Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        dValue = (period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Deductible_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Deductible_ZNATerm.ValueAsString?.toDouble()
        label = "CR ElectronicDataorComputerProgramsRestorationCosts Deductible"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR ElectronicDataorComputerProgramsRestorationCosts Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var electronicDataorComputerProgramsRestorationCostsCovPrices = period.ZSLLine.ZSL_ElectronicDataorComputerProgramsRestorationCosts_Cov_ZNA.CoveragePrices
        electronicDataorComputerProgramsRestorationCostsCovPrices?.each(\elt -> {
          label = "CR ElectronicDataorComputerProgramsRestorationCosts Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNAExists){
        label = "CR InvestigativeExpenses Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        dValue = (period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNA.ZSL_InvestigativeExpenses_LimitofLiability_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNA.ZSL_InvestigativeExpenses_LimitofLiability_ZNATerm.ValueAsString?.toDouble()
        label = "CR InvestigativeExpenses Limit"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNA.ZSLLineCovCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "CR InvestigativeExpenses Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var investigativeExpensesCovPrices = period.ZSLLine.ZSL_InvestigativeExpenses_Cov_ZNA.CoveragePrices
        investigativeExpensesCovPrices?.each(\elt -> {
          label = "CR InvestigativeExpenses Base Premium"
          dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 0 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      if (period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNAExists){
        label = "On Premises - Robbery of a Watchperson or Burglary of Property Coverage Added - U-PCNPP-3283CB Selected"
        createRowWith2Cells(worksheet, rowIndx, label, "Yes")
        rowIndx++

        var termAmt = period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNA.ZSLLineCondCosts?.firstWhere(\elt -> elt.ActualTermAmount_amt > 0)
        dValue = termAmt.ActualTermAmount_amt?.doubleValue()
        label = "On Premises - Robbery of a Watchperson or Burglary of Property Coverage Added - U-PCNPP-3283CB  Actual Premium"
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

        var roberyOfWatchPersonCovPrices = period.ZSLLine.ZSL_OnPremises_Cond_UPCNPP3283CB_ZNA.ConditionPrices
        roberyOfWatchPersonCovPrices?.each(\elt -> {
          label = "On Premises - Robbery of a Watchperson or Burglary of Property Coverage Added - U-PCNPP-3283CB  Base Premium"
          dValue = (elt.BasePremium == null) ? 0 : elt.BasePremium?.doubleValue()
          if (dValue > 0) {
            createRowWith2Cells(worksheet, rowIndx, label, dValue)
            rowIndx++
          }
        })
      }

      //Crime Ratable Employees
      var employees = period.ZSLLine.ZSLCrimeCovPart_ZNA.Employees
      var sum : double = 0
      var class1CurrentYear = period.ZSLLine.ZSLCrimeCovPart_ZNA.Employees?.firstWhere(\elt -> elt.EmployeeType == ZSLEmployeeType_ZNA.TC_CLASS_1_EMPLOYEES).CurrentYear
      employees.each(\elt -> {
        var currentYear = elt.CurrentYear
        var ratableEmployeeFactor = elt.RatableEmployeeFactor
        if (currentYear == null){
          currentYear = 0
        }
        sum += (currentYear * ratableEmployeeFactor) as double
      })
      if (sum < 5) {
        sum = 5
      }
      label = "CR Ratable Employees"
      createRowWith2Cells(worksheet, rowIndx, label, sum)
      rowIndx++


      var zslLineCovCosts = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.ZSLLineCovCosts
      zslLineCovCosts?.each(\elt -> {
        var actualBaseRate = elt.ActualBaseRate
        label = "CR BaseRate"
        dValue = (actualBaseRate == null) ? 1 : actualBaseRate?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++

      })

      label = "CR RatingMod"
      dValue = (period.ZSLLine.ZSP_CrimeModifier.RateModifier == null) ? 0 : period.ZSLLine.ZSP_CrimeModifier.RateModifier?.doubleValue()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++

      var crCovPrices = period.ZSLLine.ZSL_EmployeeTheft_Cov_ZNA.CoveragePrices
      crCovPrices?.each(\elt -> {
        label = "CRT Minimum Premium"
        dValue = (elt.MinimumPremium == null) ? 0 : elt.MinimumPremium?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Base Premium"
        dValue = (elt.ZSPRatingPricingDetail.BasePremium == null) ? 1 : elt.ZSPRatingPricingDetail.BasePremium?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Base Rate Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.BaseRateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Deductible Multiplier"
        dValue = (elt.DeductibleFactor == null) ? 1 : elt.DeductibleFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR ILF Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.ILFAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.ILFAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Industry Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.IndustryAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Location Factor"
        dValue = (elt.ZSPRatingPricingDetail.LocationFactor == null) ? 1 : elt.ZSPRatingPricingDetail.LocationFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Loss Cost Units"
        dValue = (elt.ZSPRatingPricingDetail.LossCostUnitsFactor == null) ? 1 : elt.ZSPRatingPricingDetail.LossCostUnitsFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR State Adjustment Factor"
        dValue = (elt.ZSPRatingPricingDetail.StateAdjustmentFactor == null) ? 1 : elt.ZSPRatingPricingDetail.StateAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR State LCM"
        dValue = (elt.ZSPRatingPricingDetail.StateLCMFactor == null) ? 1 : elt.ZSPRatingPricingDetail.StateLCMFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
        label = "CR Commission Adjustment Factor"
        dValue = (elt.CommissionAdjustmentFactor == null) ? 1 : elt.CommissionAdjustmentFactor?.doubleValue()
        createRowWith2Cells(worksheet, rowIndx, label, dValue)
        rowIndx++
      })

      /*
      label = "CR Fraudulent Impersonation Verification Option"
      dValue = (period.ZSLLine.? == null) ? 0 : period.ZSLLine.?.doubleValue()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++

      label = "CR LimitIsShared"
      dValue = (period.ZSLLine.? == null) ? 0 : period.ZSLLine.?.doubleValue()
      createRowWith2Cells(worksheet, rowIndx, label, dValue)
      rowIndx++

      */


    }

    return worksheet
  }


  private function SetRateFactors(worksheet : XSSFSheet, period : PolicyPeriod, rowIndx : int, isRenewingTerm : boolean) : XSSFSheet {
    if (period.ZSLLine.ZSLLineModifiers.length < 1) {
      return worksheet
    }

    //create Header Row
    var row = worksheet.createRow(rowIndx)
    var cellHA = row.createCell(0)
    cellHA.setCellValue("Modifiers")
    cellHA.setCellStyle(styleYellowHeader)

    rowIndx++

    row = worksheet.createRow(rowIndx)
    cellHA = row.createCell(0)
    cellHA.setCellValue("Category")
    cellHA.setCellStyle(styleYellowHeader)

    var cellHB = row.createCell(1)
    cellHB.setCellValue("Classification")
    cellHB.setCellStyle(styleYellowHeader)

    var cellHC = row.createCell(2)
    cellHC.setCellValue("Min")
    cellHC.setCellStyle(styleYellowHeader)

    var cellHD = row.createCell(3)
    cellHD.setCellValue("Max")
    cellHD.setCellStyle(styleYellowHeader)

    var cellHE = row.createCell(4)
    cellHE.setCellValue("Recommended")
    cellHE.setCellStyle(styleYellowHeader)

    var cellHF = row.createCell(5)
    cellHF.setCellValue("Selected")
    cellHF.setCellStyle(styleYellowHeader)

    var cellHG = row.createCell(6)
    cellHG.setCellValue("RecommendedCategory")
    cellHG.setCellStyle(styleYellowHeader)

    var cellHH = row.createCell(7)
    cellHH.setCellValue("Recommended Factor of Recommended Classification")
    cellHH.setCellStyle(styleYellowHeader)

    rowIndx++


    //get policy period factor data
    period.ZSLLine.ZSLLineModifiers?.each(\elt -> {

      //use the code identifier to determine and create a line of business prefix
      //lobPrefix is used to create unique factor names
      var codeIdentifier = elt.Pattern.CodeIdentifier
      var lobPrefix = ""
      var Crime = codeIdentifier.indexOf("_Crime")
      var DandO = codeIdentifier.indexOf("_DO")
      var EPL = codeIdentifier.indexOf("_EPL")
      var Fiduciary = codeIdentifier.indexOf("_Fiduciary")
      var SP = codeIdentifier.indexOf("_SP_")

      if (Crime > 0) {
        lobPrefix = "CR"
      }
      if (DandO > 0 || codeIdentifier == "zsl_SelectPlus_ScheduleRatePlan_zna") {
        lobPrefix = "D&O"
      }
      if (EPL > 0) {
        lobPrefix = "EPL"
      }
      if (Fiduciary > 0) {
        lobPrefix = "FID"
      }
      if (SP > 0) {
        lobPrefix = "S&P"
      }

      if (period.PolicyPeriodType_ZNA == PolicyPeriodType_ZNA.TC_RESTATED){
        lobPrefix = "Restated_" + lobPrefix
      } else if (isRenewingTerm){
        lobPrefix = "Renewing_" + lobPrefix
      } else {
        lobPrefix = "Expiring_" + lobPrefix
      }

      var ZSLLineRF_ZNA = elt.ZSLLineRateFactors.sortBy(\eltsort -> elt.Pattern.Priority)

      var map = elt.RateFactorDefaultsForSelectedClassificationMap

      ZSLLineRF_ZNA?.each(\elt2 -> {
        row = worksheet.createRow(rowIndx)
        var cellA = row.createCell(0)
        var cellB = row.createCell(1)
        var cellC = row.createCell(2)
        var cellD = row.createCell(3)
        var cellE = row.createCell(4)
        var cellF = row.createCell(5)
        var cellG = row.createCell(6)
        var cellH = row.createCell(7)

        var refData = map.get(elt2)

        if (elt2.Pattern.Name != null) {
          var factorName = lobPrefix + " " + elt2.Pattern.Name
          cellA.setCellValue(factorName)
        } else {
          cellA.setCellValue(" ")
        }
        if (elt2.SelectedCategory != null) {
          cellB.setCellValue(elt2.SelectedCategory.DisplayName)
        }
        if (refData.InnerMin != null) {
          cellC.setCellValue(refData.InnerMin?.doubleValue())
        } else {
          cellC.setCellValue(" ")
        }
        if (refData.InnerMax != null) {
          cellD.setCellValue(refData.InnerMax?.doubleValue())
        } else {
          cellD.setCellValue(" ")
        }
        if (refData.RecommendedFactor != null) {
          cellE.setCellValue(refData.RecommendedFactor?.doubleValue())
        } else {
          cellE.setCellValue(" ")
        }
        //if the selected factor is 0 - change it to 1 so the spreadsheet calculations work
        if (elt2.AssessmentWithinLimits != null) {
          if (elt2.AssessmentWithinLimits?.doubleValue() == 0) {
            cellF.setCellValue(1)
          } else if (codeIdentifier.contains("_ScheduleRatePlan_") and lobPrefix != "S&P" and elt2.AssessmentWithinLimits < 1) {
            cellF.setCellValue(elt2.AssessmentWithinLimits?.doubleValue() + 1)
          } else {
            cellF.setCellValue(elt2.AssessmentWithinLimits?.doubleValue())
          }
        } else {
          cellF.setCellValue(" ")
        }
        if (elt2.RecommendedCategory != null and elt2.RateFactorType != RateFactorType.TC_ZSL_INDUSTRY_MODIFIER_ZNA) {
          cellG.setCellValue(elt2.RecommendedCategory.DisplayName)
        } else if(elt2.RateFactorType == RateFactorType.TC_ZSL_INDUSTRY_MODIFIER_ZNA and elt2.Branch.Policy.Account.OtherOrgTypeDescription == "Not for Profit" and elt2.ZSLLineModifier.Pattern.CodeIdentifier == "zsl_SelectPlus_DOMod_zna") {
          cellG.setCellValue(elt2.Branch.ZSLLine.ZSLIndustryType.DisplayName)
        } else {
          cellG.setCellValue(" ")
        }
        if (elt2.RecommendedFactor == 0 or elt2.RecommendedFactor == null) {
          cellH.setCellValue(1)
        } else if (elt2.RecommendedFactor != 0 and elt2.RecommendedFactor != null) {
          cellH.setCellValue(elt2.RecommendedFactor?.doubleValue())
        } else {
          cellH.setCellValue(" ")
        }
        rowIndx++
      })

      if (elt.DisplayName == "Schedule Rating Plan") {
        row = worksheet.createRow(rowIndx)
        var cellSchModTotalLabel = row.createCell(0)
        var cellSchModTotal = row.createCell(1)
        var schedValue = elt.RateWithinLimits?.doubleValue() + 1
        schedValue = (schedValue == 0) ? 1 : schedValue
        cellSchModTotalLabel.setCellValue(lobPrefix + " "  + elt.DisplayName + " Total")
        cellSchModTotal.setCellValue(schedValue)
        rowIndx++
      }

    })

    return worksheet
  }


  private function SetPremiumCosts(worksheet : XSSFSheet, period : PolicyPeriod, rowIndx : int, isRenewingTerm : boolean) : XSSFSheet {
    //create Header Rows
    rowIndx++  //add blank row

    var row = worksheet.createRow(rowIndx)
    var cellHA = row.createCell(0)
    cellHA.setCellValue("Premium Cost Amounts")
    cellHA.setCellStyle(styleYellowHeader)

    rowIndx++

    row = worksheet.createRow(rowIndx)
    cellHA = row.createCell(0)
    cellHA.setCellValue("Coverage")
    cellHA.setCellStyle(styleYellowHeader)

    var cellHB = row.createCell(1)
    cellHB.setCellValue("StandardBaseRate")
    cellHB.setCellStyle(styleYellowHeader)

    var cellHC = row.createCell(2)
    cellHC.setCellValue("StandardAmount")
    cellHC.setCellStyle(styleYellowHeader)

    var cellHD = row.createCell(3)
    cellHD.setCellValue("ActualAdjRate")
    cellHD.setCellStyle(styleYellowHeader)

    var cellHE = row.createCell(4)
    cellHE.setCellValue("ActualBaseRate")
    cellHE.setCellStyle(styleYellowHeader)

    var cellHF = row.createCell(5)
    cellHF.setCellValue("ActualTermAmount")
    cellHF.setCellStyle(styleYellowHeader)

    var cellHG = row.createCell(6)
    cellHG.setCellValue("ActualAmount")
    cellHG.setCellStyle(styleYellowHeader)

    var cellHH = row.createCell(7)
    cellHH.setCellValue("Foreign Amount")
    cellHH.setCellStyle(styleYellowHeader)

    rowIndx++

    //ALM#5849 - ZSP - Exclude foreign allocation and foreign admin fees from the coverage premiums in the rate book.
    var expiringPeriodCalcPremPriorForeignAdjAvlbl = (isRenewingTerm and (period.PeriodStart.afterOrEqual(ReleaseDateUtil_ZNA.ZSP_RateChange_CalculatedPremiumPriorToForeignAdjustment_EFF_DATE_ZNA)) or
        (!(isRenewingTerm) and (period.PeriodEnd.afterOrEqual(ReleaseDateUtil_ZNA.ZSP_RateChange_CalculatedPremiumPriorToForeignAdjustment_EFF_DATE_ZNA))))
    var premiumCosts = period.ZSLLine.ChargedPremium_ZNA.PremiumCosts
    premiumCosts?.each(\elt -> {

      row = worksheet.createRow(rowIndx)
      var cellA = row.createCell(0)
      var cellB = row.createCell(1)
      var cellC = row.createCell(2)
      var cellD = row.createCell(3)
      var cellE = row.createCell(4)
      var cellF = row.createCell(5)
      var cellG = row.createCell(6)
      var cellH = row.createCell(7)

      var displayName = (elt.CostDisplayName == null) ? "" : elt.CostDisplayName
      if (displayName != ""){
        if (period.PolicyPeriodType_ZNA == PolicyPeriodType_ZNA.TC_RESTATED){
          displayName = "Restated_" + displayName
        } else if  (isRenewingTerm){
          displayName = "Renewing_" + displayName
        }else{
          displayName = "Expiring_" + displayName
        }
      }
      var standardBaseRate = (elt.StandardBaseRate == null) ? 1 : elt.StandardBaseRate?.doubleValue()
      var standardAmount = (elt.StandardAmount == null) ? 0 : elt.StandardAmount?.doubleValue()
      var actualAdjRate = (elt.ActualAdjRate == null) ? 1 : elt.ActualAdjRate?.doubleValue()
      var actualBaseRate = (elt.ActualBaseRate == null) ? 1 : elt.ActualBaseRate?.doubleValue()
      var actualTermAmount = expiringPeriodCalcPremPriorForeignAdjAvlbl ?
          (elt.CalculatedPremiumPriorToForeignAdjustmentTermAmount == null) ? 0 : elt.CalculatedPremiumPriorToForeignAdjustmentTermAmount.doubleValue() :
          (elt.ActualTermAmount == null) ? 0 : elt.ActualTermAmount?.doubleValue()
      var actualAmount = expiringPeriodCalcPremPriorForeignAdjAvlbl ?
          (elt.CalculatedPremiumPriorToForeignAdjustmentAmount == null) ? 0 : elt.CalculatedPremiumPriorToForeignAdjustmentAmount.doubleValue() :
          (elt.ActualAmount == null) ? 0 : elt.ActualAmount?.doubleValue()
      var foreignAmount = elt.ForeignAmount?.doubleValue()

      cellA.setCellValue(displayName)
      cellB.setCellValue(standardBaseRate)
      cellC.setCellValue(standardAmount)
      cellD.setCellValue(actualAdjRate)
      cellE.setCellValue(actualBaseRate)
      cellF.setCellValue(actualTermAmount)
      cellG.setCellValue(actualAmount)
      cellH.setCellValue(foreignAmount)

      rowIndx++
    })

    return worksheet
  }


  private function SetCoverageData(worksheet : XSSFSheet, period : PolicyPeriod, rowIndx : int, isRenewingTerm : boolean) : XSSFSheet {

    rowIndx++  //add blank row

    var row = worksheet.createRow(rowIndx)
    var cellHA = row.createCell(0)
    cellHA.setCellValue("Other Policy Data")
    cellHA.setCellStyle(styleYellowHeader)

    rowIndx++

    row = worksheet.createRow(rowIndx)
    cellHA = row.createCell(0)
    cellHA.setCellValue("Description")
    cellHA.setCellStyle(styleYellowHeader)

    var cellHB = row.createCell(1)
    cellHB.setCellValue("Value")
    cellHB.setCellStyle(styleYellowHeader)

    rowIndx++

    var startingRowIndex = rowIndx
    //Variables for Coverage Parts Sharing Limits
    var DO_LimitIsShared = "false"
    var FID_LimitIsShared = "false"
    var EPL_LimitIsShared = "false"
    var DO_LimitofLiability : double = 0
    var FID_LimitofLiability : double = 0
    var EPL_LimitofLiability : double = 0
    var DO_SIR : double = 0
    var EPL_SIR : double = 0
    var FID_SIR : double = 0
    var SharedAggregateLimit : double = 0
    var NumberOfSharedCovParts = 0

    //D&O
    if (period.ZSLLine.ZSL_DandO_Cov_ZNA != null) {
      if (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString != null) {
        DO_LimitofLiability = (period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString?.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.ZSL_DandO_LimitofLiabilityOther_ZNATerm.ValueAsString?.toDouble()
      } else {
        DO_LimitofLiability = (period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
      }
      DO_LimitIsShared = (period.ZSLLine.ZSL_DandO_Cov_ZNA.HasZSL_DandO_LimitIsShared_ZNATerm) ? "true" : "false"
      DO_SIR = (period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA == null) ? 0 : period.ZSLLine.ZSL_DandO_Cov_ZNA.RetentionLimit_ZNA?.doubleValue()

      createRowWith2Cells(worksheet, rowIndx, "D&O LimitIsShared", DO_LimitIsShared)
      rowIndx++
      createRowWith2Cells(worksheet, rowIndx, "D&O LimitofLiability", DO_LimitofLiability)
      rowIndx++
      createRowWith2Cells(worksheet, rowIndx, "D&O SelfInsuredRetention", DO_SIR)
      rowIndx++

    }

    //EPL
    if ( period.ZSLLine.ZSL_EPL_Cov_ZNA != null) {
      if (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString != null) {
        EPL_LimitofLiability = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString?.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitLiabOther_ZNATerm.ValueAsString?.toDouble()
      } else {
        EPL_LimitofLiability = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString?.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_LimitOfLiability_ZNATerm.ValueAsString?.toDouble()
      }
      EPL_LimitIsShared = (period.ZSLLine.ZSL_EPL_Cov_ZNA.HasZNA_EPL_LimitIsShared_ZNATerm) ? "true" : "false"
      EPL_SIR = (period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString == null) ? 0 : period.ZSLLine.ZSL_EPL_Cov_ZNA.ZSL_EPL_RetnForEachEmployPractClaim_ZNATerm.ValueAsString?.toDouble()

      createRowWith2Cells(worksheet, rowIndx, "EPL LimitIsShared", EPL_LimitIsShared)
      rowIndx++
      createRowWith2Cells(worksheet, rowIndx, "EPL LimitofLiability", EPL_LimitofLiability)
      rowIndx++
      createRowWith2Cells(worksheet, rowIndx, "EPL SelfInsuredRetention", EPL_SIR)
      rowIndx++

    }

    //Fiduciary
    if (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA != null) {
      if (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString != null) {
        FID_LimitofLiability = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString?.toDouble() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitofLiabilityOther_ZNATerm.ValueAsString?.toDouble()
      } else {
        FID_LimitofLiability = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue() < 0) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.LiabilityLimit_ZNA?.doubleValue()
      }
      FID_LimitIsShared = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_LimitisShared_ZNATerm.Value) ? "true" : "false"
      FID_SIR = (period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm == null) ? 0 : period.ZSLLine.ZSL_FiduciaryCoverage_Cov_ZNA.ZSL_FiduciaryCoverage_SelfInsuredRetention_ZNATerm.ValueAsString?.toDouble()

      createRowWith2Cells(worksheet, rowIndx, "FID LimitIsShared", FID_LimitIsShared)
      rowIndx++
      createRowWith2Cells(worksheet, rowIndx, "FID LimitofLiability", FID_LimitofLiability)
      rowIndx++
      createRowWith2Cells(worksheet, rowIndx, "FID SelfInsuredRetention", FID_SIR)
      rowIndx++

    }

    //Create Number of Shared Coverage parts
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "false" ){
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "false"  && FID_LimitIsShared == "true" ){
      NumberOfSharedCovParts = 2
    }
    if (DO_LimitIsShared == "true" && EPL_LimitIsShared == "true"  && FID_LimitIsShared == "true" ){
      NumberOfSharedCovParts = 3
    }
    createRowWith2Cells(worksheet, rowIndx, "NumberOfSharedCovParts", NumberOfSharedCovParts)
    rowIndx++


    //if any limits are shared set the SharedAggregateLimit to the largest of the limits
    if (DO_LimitIsShared == "true" && (DO_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = DO_LimitofLiability
    }
    if (FID_LimitIsShared == "true" && (FID_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = FID_LimitofLiability
    }
    if (EPL_LimitIsShared == "true" && (EPL_LimitofLiability > SharedAggregateLimit)) {
      SharedAggregateLimit = EPL_LimitofLiability
    }
    createRowWith2Cells(worksheet, rowIndx, "SharedAggregateLimit", SharedAggregateLimit)
    rowIndx++


    //get actual and model premiums
    var line = period.ZSLLine
    var zslPricingDisplayUtil = new ZSLPricingDisplayUtil(line)

    if (zslPricingDisplayUtil.coverages().Count > 0) {
      var coverages = zslPricingDisplayUtil.coverages()
      coverages?.each(\cov -> {
        var covId = cov.Pattern.CodeIdentifier
        /***
         * navdesin 05/15/2020
         * Created separated method to get actual premium for rating export to avoid null issues in subsequent transactions.
         ***/
        var actualPremium = (zslPricingDisplayUtil.actualPremiumRatingExport(period, cov) == null) ? 0 : zslPricingDisplayUtil.actualPremiumRatingExport(period, cov)
        var modelPrice = (zslPricingDisplayUtil.termModelPrice(cov) == null) ? 0 : zslPricingDisplayUtil.termModelPrice(cov)

        switch (covId) {
          case "ZSL_DandO_Cov_ZNA":
            createRowWith2Cells(worksheet, rowIndx, "D&O ActualPremium", actualPremium?.doubleValue())
            rowIndx++
            createRowWith2Cells(worksheet, rowIndx, "D&O ModelPrice", modelPrice?.doubleValue())
            rowIndx++
            break
          case "ZSL_EPL_Cov_ZNA":
            createRowWith2Cells(worksheet, rowIndx, "EPL ActualPremium", actualPremium?.doubleValue())
            rowIndx++
            createRowWith2Cells(worksheet, rowIndx, "EPL ModelPrice", modelPrice?.doubleValue())
            rowIndx++
            break
          case "ZSL_FiduciaryCoverage_Cov_ZNA":
            createRowWith2Cells(worksheet, rowIndx, "FID ActualPremium", actualPremium?.doubleValue())
            rowIndx++
            createRowWith2Cells(worksheet, rowIndx, "FID ModelPrice", modelPrice?.doubleValue())
            rowIndx++
            break
          case "ZSL_EmployeeTheft_Cov_ZNA":
            createRowWith2Cells(worksheet, rowIndx, "CR ActualPremium", actualPremium?.doubleValue())
            rowIndx++
            createRowWith2Cells(worksheet, rowIndx, "CR ModelPrice", modelPrice?.doubleValue())
            rowIndx++
            break
          default:
            break
        }
      })
    }

    var strPrefix : String = ""
    var intRowCounter = startingRowIndex
    strPrefix = isRenewingTerm ? "Renewing_":"Expiring_"

    if (period.PolicyPeriodType_ZNA == PolicyPeriodType_ZNA.TC_RESTATED){
      strPrefix = "Restated_"
    }

    while(intRowCounter<rowIndx){
      if (worksheet.getRow(intRowCounter).getCell(0).getStringCellValue() != ""){
        worksheet.getRow(intRowCounter).getCell(0).setCellValue(strPrefix + worksheet.getRow(intRowCounter).getCell(0).getStringCellValue())
      }
      intRowCounter++
    }
    return worksheet
  }


  private function SetWorkbookExpiringData(worksheet:XSSFSheet, year2RateChangeData:Year2RateChangeData, rowIndx:int) : XSSFSheet {

    var valueString = ""
    var valueDouble : double = 0
    var rateChangeInput = year2RateChangeData.InputData

    var row = worksheet.createRow(rowIndx)
    var cellHA = row.createCell(0)
    cellHA.setCellValue("Expiring Data")
    cellHA.setCellStyle(styleYellowHeader)

    rowIndx++


    valueString = (rateChangeInput.Expiring_PolicyNumber == null) ? "" : rateChangeInput.Expiring_PolicyNumber
    createRowWith2Cells(worksheet, rowIndx, "Expiring_PolicyNumber", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_QuoteNumber == null) ? "" : rateChangeInput.Expiring_QuoteNumber
    createRowWith2Cells(worksheet, rowIndx, "Expiring_QuoteNumber", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_Iteration == null) ? "" : rateChangeInput.Expiring_Iteration
    createRowWith2Cells(worksheet, rowIndx, "Expiring_Iteration", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_Insured == null) ? "" : rateChangeInput.Expiring_Insured
    createRowWith2Cells(worksheet, rowIndx, "Expiring_Insured", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_EffectiveDate == null) ? "" : rateChangeInput.Expiring_EffectiveDate
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EffectiveDate", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_ExpirationDate == null) ? "" : rateChangeInput.Expiring_ExpirationDate
    createRowWith2Cells(worksheet, rowIndx, "Expiring_ExpirationDate", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_DomiciledState == null) ? "" : rateChangeInput.Expiring_DomiciledState
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DomiciledState", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_Private_NonProfit == null) ? "" : rateChangeInput.Expiring_Private_NonProfit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_Private_NonProfit", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_TypeOfNonProfit == null) ? "" : rateChangeInput.Expiring_TypeOfNonProfit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_TypeOfNonProfit", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_FidelityClassCode == null) ? "" : rateChangeInput.Expiring_FidelityClassCode
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FidelityClassCode", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_Commission < 0) ? 0 : rateChangeInput.Expiring_Commission
    createRowWith2Cells(worksheet, rowIndx, "Expiring_Commission", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_NAICSCode == null) ? "" : rateChangeInput.Expiring_NAICSCode
    createRowWith2Cells(worksheet, rowIndx, "Expiring_NAICSCode", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_NAICSDescription == null) ? "" : rateChangeInput.Expiring_NAICSDescription
    createRowWith2Cells(worksheet, rowIndx, "Expiring_NAICSDescription", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_Primary_Excess == null) ? "" : rateChangeInput.Expiring_Primary_Excess
    createRowWith2Cells(worksheet, rowIndx, "Expiring_Primary_Excess", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_PolicyLimit < 0) ? 0 : rateChangeInput.Expiring_PolicyLimit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_PolicyLimit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_PolicyAggregateLimit < 0) ? 0 : rateChangeInput.Expiring_PolicyAggregateLimit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_PolicyAggregateLimit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_PolicyAttachmentPoint < 0) ? 0 : rateChangeInput.Expiring_PolicyAttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Expiring_PolicyAttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_PolicySIR < 0) ? 0 : rateChangeInput.Expiring_PolicySIR
    createRowWith2Cells(worksheet, rowIndx, "Expiring_PolicySIR", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_PolicyAggSIR < 0) ? 0 : rateChangeInput.Expiring_PolicyAggSIR
    createRowWith2Cells(worksheet, rowIndx, "Expiring_PolicyAggSIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_UniqueAndUnusual == null) ? "No" : rateChangeInput.Expiring_UniqueAndUnusual
    createRowWith2Cells(worksheet, rowIndx, "Expiring_UniqueAndUnusual", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_TotalAssets < 0) ? 0 : rateChangeInput.Expiring_TotalAssets
    createRowWith2Cells(worksheet, rowIndx, "Expiring_TotalAssets", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_TotalPlanAssets  < 0) ? 0 : rateChangeInput.Expiring_TotalPlanAssets
    createRowWith2Cells(worksheet, rowIndx, "Expiring_TotalPlanAssets", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_SharedLimit == null) ? "" : rateChangeInput.Expiring_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_ActualCharged < 0) ? 0 : rateChangeInput.Expiring_ActualCharged
    createRowWith2Cells(worksheet, rowIndx, "Expiring_ActualCharged", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CrimeTotalLocations < 0) ? 0 : rateChangeInput.Expiring_CrimeTotalLocations
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CrimeTotalLocations", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesUS < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesUS
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesUS", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_PartTimeEmployees < 0) ? 0 : rateChangeInput.Expiring_PartTimeEmployees
    createRowWith2Cells(worksheet, rowIndx, "Expiring_PartTimeEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_IndependentContractors < 0) ? 0 : rateChangeInput.Expiring_IndependentContractors
    createRowWith2Cells(worksheet, rowIndx, "Expiring_IndependentContractors", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_ForeignEmployees < 0) ? 0 : rateChangeInput.Expiring_ForeignEmployees
    createRowWith2Cells(worksheet, rowIndx, "Expiring_ForeignEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_UnionEmployees < 0) ? 0 : rateChangeInput.Expiring_UnionEmployees
    createRowWith2Cells(worksheet, rowIndx, "Expiring_UnionEmployees", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_State1 == null) ? "" : rateChangeInput.Expiring_State1
    createRowWith2Cells(worksheet, rowIndx, "Expiring_State1", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_State2 == null) ? "" : rateChangeInput.Expiring_State2
    createRowWith2Cells(worksheet, rowIndx, "Expiring_State2", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_State3 == null) ? "" : rateChangeInput.Expiring_State3
    createRowWith2Cells(worksheet, rowIndx, "Expiring_State3", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_State4 == null) ? "" : rateChangeInput.Expiring_State4
    createRowWith2Cells(worksheet, rowIndx, "Expiring_State4", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_State5 == null) ? "" : rateChangeInput.Expiring_State5
    createRowWith2Cells(worksheet, rowIndx, "Expiring_State5", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesUSState1 < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesUSState1
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesUSState1", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesUSState2 < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesUSState2
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesUSState2", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesUSState3 < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesUSState3
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesUSState3", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesUSState4 < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesUSState4
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesUSState4", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesUSState5 < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesUSState5
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesUSState5", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesForeign < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesForeign
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesForeign", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesUSStateCA < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesUSStateCA
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesCA", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FullTimeEmployeesAllOther < 0) ? 0 : rateChangeInput.Expiring_FullTimeEmployeesAllOther
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FullTimeEmployeesAllOther", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesState1 < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesState1
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesState1", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesState2 < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesState2
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesState2", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesState3 < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesState3
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesState3", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesState4 < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesState4
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesState4", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesState5 < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesState5
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesState5", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesForeign < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesForeign
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesForeign", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesCA < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesCA
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesCA", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_RateableEmployeesAllOther < 0) ? 0 : rateChangeInput.Expiring_RateableEmployeesAllOther
    createRowWith2Cells(worksheet, rowIndx, "Expiring_RateableEmployeesAllOther", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_RateableEmployees < 0) ? 0 : rateChangeInput.Expiring_EPL_RateableEmployees
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_RateableEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_RateableEmployees < 0) ? 0 : rateChangeInput.Expiring_CR_RateableEmployees
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_RateableEmployees", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_DO_YN == null) ? "" : rateChangeInput.Expiring_DO_YN
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_DO_SIR < 0) ? 0 : rateChangeInput.Expiring_DO_SIR
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_SIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_DO_SharedLimit == null) ? "" : rateChangeInput.Expiring_DO_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_DO_Limit < 0) ? 0 : rateChangeInput.Expiring_DO_Limit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_DO_AttachmentPoint < 0) ? 0 : rateChangeInput.Expiring_DO_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_AttachmentPoint", valueDouble)
    rowIndx++
/*
    valueDouble = (rateChangeInput.Expiring_DO_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Expiring_DO_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_SharedLimitDiscount", valueDouble)
    rowIndx++
*/
    valueDouble = (rateChangeInput.Expiring_DO_BasePremium < 0) ? 0 : rateChangeInput.Expiring_DO_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_DO_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Expiring_DO_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_DO_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_EPL_YN == null) ? "" : rateChangeInput.Expiring_EPL_YN
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_SIR < 0) ? 0 : rateChangeInput.Expiring_EPL_SIR
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_SIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_EPL_SharedLimit == null) ? "" : rateChangeInput.Expiring_EPL_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_Limit < 0) ? 0 : rateChangeInput.Expiring_EPL_Limit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_AttachmentPoint < 0) ? 0 : rateChangeInput.Expiring_EPL_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Expiring_EPL_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_SharedLimitDiscount", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_SeparateLimitSurcharge < 0) ? 0 : rateChangeInput.Expiring_EPL_SeparateLimitSurcharge
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_SeparateLimitSurcharge", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_BasePremium < 0) ? 0 : rateChangeInput.Expiring_EPL_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_EPL_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Expiring_EPL_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_EPL_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_FID_YN == null) ? "" : rateChangeInput.Expiring_FID_YN
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FID_SIR < 0) ? 0 : rateChangeInput.Expiring_FID_SIR
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_SIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_FID_SharedLimit == null) ? "" : rateChangeInput.Expiring_FID_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FID_Limit < 0) ? 0 : rateChangeInput.Expiring_FID_Limit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FID_AttachmentPoint < 0) ? 0 : rateChangeInput.Expiring_FID_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FID_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Expiring_FID_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_SharedLimitDiscount", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FID_SeparateLimitSurcharge < 0) ? 0 : rateChangeInput.Expiring_FID_SeparateLimitSurcharge
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_SeparateLimitSurcharge", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FID_BasePremium < 0) ? 0 : rateChangeInput.Expiring_FID_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_FID_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Expiring_FID_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_FID_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Expiring_CR_YN == null) ? "" : rateChangeInput.Expiring_CR_YN
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_YN", valueString)
    rowIndx++

    valueString = (rateChangeInput.Expiring_CR_SharedLimit == null) ? "" : rateChangeInput.Expiring_CR_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_Limit < 0) ? 0 : rateChangeInput.Expiring_CR_Limit
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_AttachmentPoint < 0) ? 0 : rateChangeInput.Expiring_CR_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Expiring_CR_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_SharedLimitDiscount", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_SeparateLimitSurcharge < 0) ? 0 : rateChangeInput.Expiring_CR_SeparateLimitSurcharge
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_SeparateLimitSurcharge", valueDouble)
    rowIndx++
/*
    valueDouble = (rateChangeInput.Expiring_CR_BasePremium < 0) ? 0 : rateChangeInput.Expiring_CR_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_BasePremium", valueDouble)
    rowIndx++
*/
    valueDouble = (rateChangeInput.Expiring_CR_EmployeeTheftAnnualChargedPremium < 0) ? 0 : rateChangeInput.Expiring_CR_EmployeeTheftAnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_EmployeeTheftAnnualChargedPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_TotalAnnualChargedPremium < 0) ? 0 : rateChangeInput.Expiring_CR_TotalAnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_EndorsementPremium < 0) ? 0 : rateChangeInput.Expiring_CR_EndorsementPremium
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_EndorsementPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_LimitPerClaim < 0) ? 0 : rateChangeInput.Expiring_CR_LimitPerClaim
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_LimitPerClaim", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_Deductible < 0) ? 0 : rateChangeInput.Expiring_CR_Deductible
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_Deductible", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_NumberOfLocations < 0) ? 0 : rateChangeInput.Expiring_CR_NumberOfLocations
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_NumberOfLocations", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Expiring_CR_LocationFactor < 0) ? 0 : rateChangeInput.Expiring_CR_LocationFactor
    createRowWith2Cells(worksheet, rowIndx, "Expiring_CR_LocationFactor", valueDouble)
    rowIndx++

    worksheet = SetRateFactors(worksheet, _expiringPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetPremiumCosts(worksheet, _expiringPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetCoverageData(worksheet, _expiringPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetEnhancement(worksheet, _expiringPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    return worksheet
  }

  private function SetWorkbookRestatedData(worksheet:XSSFSheet, year2RateChangeData:Year2RateChangeData, rowIndx:int) : XSSFSheet {

    var valueString = ""
    var valueDouble : double = 0
    var rateChangeInput = year2RateChangeData.InputData

    var row = worksheet.createRow(rowIndx)
    var cellHA = row.createCell(0)
    cellHA.setCellValue("Restated Data")
    cellHA.setCellStyle(styleYellowHeader)

    rowIndx++


    valueString = (rateChangeInput.Restated_PolicyNumber == null) ? "" : rateChangeInput.Restated_PolicyNumber
    createRowWith2Cells(worksheet, rowIndx, "Restated_PolicyNumber", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_QuoteNumber == null) ? "" : rateChangeInput.Restated_QuoteNumber
    createRowWith2Cells(worksheet, rowIndx, "Restated_QuoteNumber", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_Iteration == null) ? "" : rateChangeInput.Restated_Iteration
    createRowWith2Cells(worksheet, rowIndx, "Restated_Iteration", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_Insured == null) ? "" : rateChangeInput.Restated_Insured
    createRowWith2Cells(worksheet, rowIndx, "Restated_Insured", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_EffectiveDate == null) ? "" : rateChangeInput.Restated_EffectiveDate
    createRowWith2Cells(worksheet, rowIndx, "Restated_EffectiveDate", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_ExpirationDate == null) ? "" : rateChangeInput.Restated_ExpirationDate
    createRowWith2Cells(worksheet, rowIndx, "Restated_ExpirationDate", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_DomiciledState == null) ? "" : rateChangeInput.Restated_DomiciledState
    createRowWith2Cells(worksheet, rowIndx, "Restated_DomiciledState", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_Private_NonProfit == null) ? "" : rateChangeInput.Restated_Private_NonProfit
    createRowWith2Cells(worksheet, rowIndx, "Restated_Private_NonProfit", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_TypeOfNonProfit == null) ? "" : rateChangeInput.Restated_TypeOfNonProfit
    createRowWith2Cells(worksheet, rowIndx, "Restated_TypeOfNonProfit", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_FidelityClassCode == null) ? "" : rateChangeInput.Restated_FidelityClassCode
    createRowWith2Cells(worksheet, rowIndx, "Restated_FidelityClassCode", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_Commission < 0) ? 0 : rateChangeInput.Restated_Commission
    createRowWith2Cells(worksheet, rowIndx, "Restated_Commission", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_NAICSCode == null) ? "" : rateChangeInput.Restated_NAICSCode
    createRowWith2Cells(worksheet, rowIndx, "Restated_NAICSCode", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_NAICSDescription == null) ? "" : rateChangeInput.Restated_NAICSDescription
    createRowWith2Cells(worksheet, rowIndx, "Restated_NAICSDescription", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_Primary_Excess == null) ? "" : rateChangeInput.Restated_Primary_Excess
    createRowWith2Cells(worksheet, rowIndx, "Restated_Primary_Excess", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_PolicyLimit < 0) ? 0 : rateChangeInput.Restated_PolicyLimit
    createRowWith2Cells(worksheet, rowIndx, "Restated_PolicyLimit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_PolicyAggregateLimit < 0) ? 0 : rateChangeInput.Restated_PolicyAggregateLimit
    createRowWith2Cells(worksheet, rowIndx, "Restated_PolicyAggregateLimit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_PolicyAttachmentPoint < 0) ? 0 : rateChangeInput.Restated_PolicyAttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Restated_PolicyAttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_PolicySIR < 0) ? 0 : rateChangeInput.Restated_PolicySIR
    createRowWith2Cells(worksheet, rowIndx, "Restated_PolicySIR", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_PolicyAggSIR < 0) ? 0 : rateChangeInput.Restated_PolicyAggSIR
    createRowWith2Cells(worksheet, rowIndx, "Restated_PolicyAggSIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_UniqueAndUnusual == null) ? "No" : rateChangeInput.Restated_UniqueAndUnusual
    createRowWith2Cells(worksheet, rowIndx, "Restated_UniqueAndUnusual", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_TotalAssets < 0) ? 0 : rateChangeInput.Restated_TotalAssets
    createRowWith2Cells(worksheet, rowIndx, "Restated_TotalAssets", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_TotalPlanAssets  < 0) ? 0 : rateChangeInput.Restated_TotalPlanAssets
    createRowWith2Cells(worksheet, rowIndx, "Restated_TotalPlanAssets", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_SharedLimit == null) ? "" : rateChangeInput.Restated_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Restated_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_ActualCharged < 0) ? 0 : rateChangeInput.Restated_ActualCharged
    createRowWith2Cells(worksheet, rowIndx, "Restated_ActualCharged", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CrimeTotalLocations < 0) ? 0 : rateChangeInput.Restated_CrimeTotalLocations
    createRowWith2Cells(worksheet, rowIndx, "Restated_CrimeTotalLocations", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesUS < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesUS
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesUS", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_PartTimeEmployees < 0) ? 0 : rateChangeInput.Restated_PartTimeEmployees
    createRowWith2Cells(worksheet, rowIndx, "Restated_PartTimeEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_IndependentContractors < 0) ? 0 : rateChangeInput.Restated_IndependentContractors
    createRowWith2Cells(worksheet, rowIndx, "Restated_IndependentContractors", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_ForeignEmployees < 0) ? 0 : rateChangeInput.Restated_ForeignEmployees
    createRowWith2Cells(worksheet, rowIndx, "Restated_ForeignEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_UnionEmployees < 0) ? 0 : rateChangeInput.Restated_UnionEmployees
    createRowWith2Cells(worksheet, rowIndx, "Restated_UnionEmployees", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_State1 == null) ? "" : rateChangeInput.Restated_State1
    createRowWith2Cells(worksheet, rowIndx, "Restated_State1", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_State2 == null) ? "" : rateChangeInput.Restated_State2
    createRowWith2Cells(worksheet, rowIndx, "Restated_State2", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_State3 == null) ? "" : rateChangeInput.Restated_State3
    createRowWith2Cells(worksheet, rowIndx, "Restated_State3", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_State4 == null) ? "" : rateChangeInput.Restated_State4
    createRowWith2Cells(worksheet, rowIndx, "Restated_State4", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_State5 == null) ? "" : rateChangeInput.Restated_State5
    createRowWith2Cells(worksheet, rowIndx, "Restated_State5", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesUSState1 < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesUSState1
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesUSState1", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesUSState2 < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesUSState2
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesUSState2", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesUSState3 < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesUSState3
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesUSState3", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesUSState4 < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesUSState4
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesUSState4", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesUSState5 < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesUSState5
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesUSState5", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesForeign < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesForeign
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesForeign", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesUSStateCA < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesUSStateCA
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesCA", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FullTimeEmployeesAllOther < 0) ? 0 : rateChangeInput.Restated_FullTimeEmployeesAllOther
    createRowWith2Cells(worksheet, rowIndx, "Restated_FullTimeEmployeesAllOther", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesState1 < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesState1
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesState1", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesState2 < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesState2
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesState2", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesState3 < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesState3
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesState3", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesState4 < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesState4
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesState4", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesState5 < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesState5
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesState5", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesForeign < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesForeign
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesForeign", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesCA < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesCA
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesCA", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_RateableEmployeesAllOther < 0) ? 0 : rateChangeInput.Restated_RateableEmployeesAllOther
    createRowWith2Cells(worksheet, rowIndx, "Restated_RateableEmployeesAllOther", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_RateableEmployees < 0) ? 0 : rateChangeInput.Restated_EPL_RateableEmployees
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_RateableEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_RateableEmployees < 0) ? 0 : rateChangeInput.Restated_CR_RateableEmployees
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_RateableEmployees", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_DO_YN == null) ? "" : rateChangeInput.Restated_DO_YN
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_DO_SIR < 0) ? 0 : rateChangeInput.Restated_DO_SIR
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_SIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_DO_SharedLimit == null) ? "" : rateChangeInput.Restated_DO_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_DO_Limit < 0) ? 0 : rateChangeInput.Restated_DO_Limit
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_DO_AttachmentPoint < 0) ? 0 : rateChangeInput.Restated_DO_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_AttachmentPoint", valueDouble)
    rowIndx++
/*
    valueDouble = (rateChangeInput.Restated_DO_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Restated_DO_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_SharedLimitDiscount", valueDouble)
    rowIndx++
*/
    valueDouble = (rateChangeInput.Restated_DO_BasePremium < 0) ? 0 : rateChangeInput.Restated_DO_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_DO_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Restated_DO_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_DO_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_EPL_YN == null) ? "" : rateChangeInput.Restated_EPL_YN
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_SIR < 0) ? 0 : rateChangeInput.Restated_EPL_SIR
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_SIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_EPL_SharedLimit == null) ? "" : rateChangeInput.Restated_EPL_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_Limit < 0) ? 0 : rateChangeInput.Restated_EPL_Limit
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_AttachmentPoint < 0) ? 0 : rateChangeInput.Restated_EPL_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Restated_EPL_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_SharedLimitDiscount", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_SeparateLimitSurcharge < 0) ? 0 : rateChangeInput.Restated_EPL_SeparateLimitSurcharge
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_SeparateLimitSurcharge", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_BasePremium < 0) ? 0 : rateChangeInput.Restated_EPL_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_EPL_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Restated_EPL_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_EPL_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_FID_YN == null) ? "" : rateChangeInput.Restated_FID_YN
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FID_SIR < 0) ? 0 : rateChangeInput.Restated_FID_SIR
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_SIR", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_FID_SharedLimit == null) ? "" : rateChangeInput.Restated_FID_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FID_Limit < 0) ? 0 : rateChangeInput.Restated_FID_Limit
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FID_AttachmentPoint < 0) ? 0 : rateChangeInput.Restated_FID_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FID_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Restated_FID_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_SharedLimitDiscount", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FID_SeparateLimitSurcharge < 0) ? 0 : rateChangeInput.Restated_FID_SeparateLimitSurcharge
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_SeparateLimitSurcharge", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FID_BasePremium < 0) ? 0 : rateChangeInput.Restated_FID_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_FID_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Restated_FID_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_FID_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Restated_CR_YN == null) ? "" : rateChangeInput.Restated_CR_YN
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_YN", valueString)
    rowIndx++

    valueString = (rateChangeInput.Restated_CR_SharedLimit == null) ? "" : rateChangeInput.Restated_CR_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_Limit < 0) ? 0 : rateChangeInput.Restated_CR_Limit
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_AttachmentPoint < 0) ? 0 : rateChangeInput.Restated_CR_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_CombinedCoverageDiscount < 0) ? 0 : rateChangeInput.Restated_CR_CombinedCoverageDiscount
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_SharedLimitDiscount", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_SeparateLimitSurcharge < 0) ? 0 : rateChangeInput.Restated_CR_SeparateLimitSurcharge
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_SeparateLimitSurcharge", valueDouble)
    rowIndx++
/*
    valueDouble = (rateChangeInput.Restated_CR_BasePremium < 0) ? 0 : rateChangeInput.Restated_CR_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_BasePremium", valueDouble)
    rowIndx++
*/
    valueDouble = (rateChangeInput.Restated_CR_EmployeeTheftAnnualChargedPremium < 0) ? 0 : rateChangeInput.Restated_CR_EmployeeTheftAnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_EmployeeTheftAnnualChargedPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_TotalAnnualChargedPremium < 0) ? 0 : rateChangeInput.Restated_CR_TotalAnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_EndorsementPremium < 0) ? 0 : rateChangeInput.Restated_CR_EndorsementPremium
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_EndorsementPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_LimitPerClaim < 0) ? 0 : rateChangeInput.Restated_CR_LimitPerClaim
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_LimitPerClaim", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_Deductible < 0) ? 0 : rateChangeInput.Restated_CR_Deductible
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_Deductible", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_NumberOfLocations < 0) ? 0 : rateChangeInput.Restated_CR_NumberOfLocations
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_NumberOfLocations", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Restated_CR_LocationFactor < 0) ? 0 : rateChangeInput.Restated_CR_LocationFactor
    createRowWith2Cells(worksheet, rowIndx, "Restated_CR_LocationFactor", valueDouble)
    rowIndx++

    worksheet = SetRateFactors(worksheet, _restatedPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetPremiumCosts(worksheet, _restatedPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetCoverageData(worksheet, _restatedPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetEnhancement(worksheet, _restatedPeriod, rowIndx, false)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    return worksheet
  }

  private function SetWorkbookRenewingData(worksheet:XSSFSheet, year2RateChangeData:Year2RateChangeData, rowIndx:int) : XSSFSheet {
    // add two blank rows
    rowIndx++
    rowIndx++
    rowIndx++

    var valueString = ""
    var valueDouble : double
    var rateChangeInput = year2RateChangeData.InputData

    var row = worksheet.createRow(rowIndx)
    var cellHA = row.createCell(0)
    cellHA.setCellValue("Renewing Data")
    cellHA.setCellStyle(styleYellowHeader)

    rowIndx++


    valueString = (rateChangeInput.Renewing_PolicyNumber == null) ? "" : rateChangeInput.Renewing_PolicyNumber
    createRowWith2Cells(worksheet, rowIndx, "Renewing_PolicyNumber", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_QuoteNumber == null) ? "" : rateChangeInput.Renewing_QuoteNumber
    createRowWith2Cells(worksheet, rowIndx, "Renewing_QuoteNumber", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_BranchName == null) ? "" : rateChangeInput.Renewing_BranchName
    createRowWith2Cells(worksheet, rowIndx, "Renewing_BranchName", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_Insured == null) ? "" : rateChangeInput.Renewing_Insured
    createRowWith2Cells(worksheet, rowIndx, "Renewing_Insured", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_EffectiveDate == null) ? "" : rateChangeInput.Renewing_EffectiveDate
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EffectiveDate", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_ExpirationDate == null) ? "" : rateChangeInput.Renewing_ExpirationDate
    createRowWith2Cells(worksheet, rowIndx, "Renewing_ExpirationDate", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_DomiciledState == null) ? "" : rateChangeInput.Renewing_DomiciledState
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DomiciledState", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_Private_NonProfit == null) ? "" : rateChangeInput.Renewing_Private_NonProfit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_Private_NonProfit", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_FidelityClassCode == null) ? "" : rateChangeInput.Renewing_FidelityClassCode
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FidelityClassCode", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_Commission < 0) ? 0 : rateChangeInput.Renewing_Commission
    createRowWith2Cells(worksheet, rowIndx, "Renewing_Commission", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_Primary_Excess == null) ? "" : rateChangeInput.Renewing_Primary_Excess
    createRowWith2Cells(worksheet, rowIndx, "Renewing_Primary_Excess", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_NAICSCode == null) ? "" : rateChangeInput.Renewing_NAICSCode
    createRowWith2Cells(worksheet, rowIndx, "Renewing_NAICSCode", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_NAICSDescription == null) ? "" : rateChangeInput.Renewing_NAICSDescription
    createRowWith2Cells(worksheet, rowIndx, "Renewing_NAICSDescription", valueString)
    rowIndx++


    valueString = (rateChangeInput.Renewing_IndustryType == null ) ? "" : rateChangeInput.Renewing_IndustryType
    createRowWith2Cells(worksheet, rowIndx, "Renewing_IndustryType", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_IndustryTypeCode == null) ? "" : rateChangeInput.Renewing_IndustryTypeCode
    createRowWith2Cells(worksheet, rowIndx, "Renewing_IndustryTypeCode", valueString)
    rowIndx++

    //valueDouble = (rateChangeInput.Renewing_ActualCharged < 0) ? 0 : rateChangeInput.Renewing_ActualCharged
    //createRowWith2Cells(worksheet, rowIndx, "Renewing_ActualCharged", valueDouble)
    //rowIndx++

    valueString = (rateChangeInput.Renewing_UniqueAndUnusual == null) ? "No" : rateChangeInput.Renewing_UniqueAndUnusual
    createRowWith2Cells(worksheet, rowIndx, "Renewing_UniqueAndUnusual", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_TotalAssets < 0) ? 0 : rateChangeInput.Renewing_TotalAssets
    createRowWith2Cells(worksheet, rowIndx, "Renewing_TotalAssets", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_PlanAssets  < 0) ? 0 : rateChangeInput.Renewing_PlanAssets
    createRowWith2Cells(worksheet, rowIndx, "Renewing_TotalPlanAssets", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_SharedLimit < "") ? "" : rateChangeInput.Renewing_SharedLimit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_SharedLimit", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_NumberOfSharedCovParts < 0) ? 0 : rateChangeInput.Renewing_NumberOfSharedCovParts
    createRowWith2Cells(worksheet, rowIndx, "Renewing_NumberOfSharedCovParts", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_SharedAggregateLimit < 0) ? 0 : rateChangeInput.Renewing_SharedAggregateLimit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_SharedAggregateLimit", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_DO_YN == null) ? "" : rateChangeInput.Renewing_DO_YN
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_DO_Limit < 0) ? 0 : rateChangeInput.Renewing_DO_Limit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_Limit", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_DO_LimitIsShared == null) ? "" : rateChangeInput.Renewing_DO_LimitIsShared
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_LimitIsShared", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_DO_AttachmentPoint < 0) ? 0 : rateChangeInput.Renewing_DO_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_DO_SIR < 0) ? 0 : rateChangeInput.Renewing_DO_SIR
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_SIR", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_DO_BasePremium < 0) ? 0 : rateChangeInput.Renewing_DO_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_DO_SharedLimitCredit < 0) ? 0 : rateChangeInput.Renewing_DO_SharedLimitCredit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_SharedLimitCredit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_DO_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Renewing_DO_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_DO_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_EPL_YN == null) ? "" : rateChangeInput.Renewing_EPL_YN
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_EPL_Limit < 0) ? 0 : rateChangeInput.Renewing_EPL_Limit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_Limit", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_EPL_LimitIsShared == null) ? "" : rateChangeInput.Renewing_EPL_LimitIsShared
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_LimitIsShared", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_EPL_AttachmentPoint < 0) ? 0 : rateChangeInput.Renewing_EPL_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_EPL_SIR < 0) ? 0 : rateChangeInput.Renewing_EPL_SIR
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_SIR", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_EPL_BasePremium < 0) ? 0 : rateChangeInput.Renewing_EPL_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_EPL_SharedLimitCredit < 0) ? 0 : rateChangeInput.Renewing_EPL_SharedLimitCredit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_SharedLimitCredit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_EPL_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Renewing_EPL_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesUS < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesUS
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUS", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_PartTimeEmployees < 0) ? 0 : rateChangeInput.Renewing_PartTimeEmployees
    createRowWith2Cells(worksheet, rowIndx, "Renewing_PartTimeEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_IndependentContractors < 0) ? 0 : rateChangeInput.Renewing_IndependentContractors
    createRowWith2Cells(worksheet, rowIndx, "Renewing_IndependentContractors", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_ForeignEmployees < 0) ? 0 : rateChangeInput.Renewing_ForeignEmployees
    createRowWith2Cells(worksheet, rowIndx, "Renewing_ForeignEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_UnionEmployees < 0) ? 0 : rateChangeInput.Renewing_UnionEmployees
    createRowWith2Cells(worksheet, rowIndx, "Renewing_UnionEmployees", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_Volunteers < 0) ? 0 : rateChangeInput.Renewing_Volunteers
    createRowWith2Cells(worksheet, rowIndx, "Renewing_Volunteers", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_State1 == null) ? "" : rateChangeInput.Renewing_State1
    createRowWith2Cells(worksheet, rowIndx, "Renewing_State1", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_State2 == null) ? "" : rateChangeInput.Renewing_State2
    createRowWith2Cells(worksheet, rowIndx, "Renewing_State2", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_State3 == null) ? "" : rateChangeInput.Renewing_State3
    createRowWith2Cells(worksheet, rowIndx, "Renewing_State3", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_State4 == null) ? "" : rateChangeInput.Renewing_State4
    createRowWith2Cells(worksheet, rowIndx, "Renewing_State4", valueString)
    rowIndx++

    valueString = (rateChangeInput.Renewing_State5 == null) ? "" : rateChangeInput.Renewing_State5
    createRowWith2Cells(worksheet, rowIndx, "Renewing_State5", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesForeign < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesForeign
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesForeign", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesUSState1 < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesUSState1
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUSState1", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesUSState2 < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesUSState2
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUSState2", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesUSState3 < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesUSState3
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUSState3", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesUSState4 < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesUSState4
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUSState4", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesUSState5 < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesUSState5
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUSState5", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesUSStateCA < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesUSStateCA
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUSStateCA", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FullTimeEmployeesAllOther < 0) ? 0 : rateChangeInput.Renewing_FullTimeEmployeesAllOther
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FullTimeEmployeesUSAllOther", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesState1 < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesState1
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesState1", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesState2 < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesState2
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesState2", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesState3 < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesState3
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesState3", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesState4 < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesState4
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesState4", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesState5 < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesState5
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesState5", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesForeign < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesForeign
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesForeign", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesCA < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesCA
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesCA", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_RateableEmployeesAllOther < 0) ? 0 : rateChangeInput.Renewing_RateableEmployeesAllOther
    createRowWith2Cells(worksheet, rowIndx, "Renewing_RateableEmployeesAllOther", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_EPL_RatableEmployees < 0) ? 0 : rateChangeInput.Renewing_EPL_RatableEmployees
    createRowWith2Cells(worksheet, rowIndx, "Renewing_EPL_RatableEmployees", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_FID_YN == null) ? "" : rateChangeInput.Renewing_FID_YN
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FID_Limit < 0) ? 0 : rateChangeInput.Renewing_FID_Limit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_Limit", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_FID_LimitIsShared == null) ? "" : rateChangeInput.Renewing_FID_LimitIsShared
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_LimitIsShared", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FID_AttachmentPoint < 0) ? 0 : rateChangeInput.Renewing_FID_AttachmentPoint
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_AttachmentPoint", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FID_SIR < 0) ? 0 : rateChangeInput.Renewing_FID_SIR
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_SIR", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FID_BasePremium < 0) ? 0 : rateChangeInput.Renewing_FID_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FID_SharedLimitCredit < 0) ? 0 : rateChangeInput.Renewing_FID_SharedLimitCredit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_SharedLimitCredit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_FID_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Renewing_FID_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_FID_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_CR_YN == null) ? "" : rateChangeInput.Renewing_CR_YN
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_YN", valueString)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_Limit < 0) ? 0 : rateChangeInput.Renewing_CR_Limit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_Limit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_BasePremium < 0) ? 0 : rateChangeInput.Renewing_CR_BasePremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_BasePremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_SharedLimitCredit < 0) ? 0 : rateChangeInput.Renewing_CR_SharedLimitCredit
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_SharedLimitCredit", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_EmployeeTheftAnnualChargedPremium < 0) ? 0 : rateChangeInput.Renewing_CR_EmployeeTheftAnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_EmployeeTheftAnnualChargedPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_AnnualChargedPremium < 0) ? 0 : rateChangeInput.Renewing_CR_AnnualChargedPremium
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_AnnualChargedPremium", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_Deductible < 0) ? 0 : rateChangeInput.Renewing_CR_Deductible
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_SIR", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_NumberOfLocations < 0) ? 0 : rateChangeInput.Renewing_CR_NumberOfLocations
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_NumberOfLocations", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_LocationFactor < 0) ? 0 : rateChangeInput.Renewing_CR_LocationFactor
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_LocationFactor", valueDouble)
    rowIndx++

    valueDouble = (rateChangeInput.Renewing_CR_RatableEmployees < 0) ? 0 : rateChangeInput.Renewing_CR_RatableEmployees
    createRowWith2Cells(worksheet, rowIndx, "Renewing_CR_RatableEmployees", valueDouble)
    rowIndx++

    valueString = (rateChangeInput.Renewing_UniqueAndUnusual == null) ? "No" : rateChangeInput.Renewing_UniqueAndUnusual
    createRowWith2Cells(worksheet, rowIndx, "Renewing_UniqueAndUnusual", valueString)
    rowIndx++

    worksheet = SetRateFactors(worksheet, _renewalPeriod, rowIndx, true)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetPremiumCosts(worksheet, _renewalPeriod, rowIndx, true)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetCoverageData(worksheet, _renewalPeriod, rowIndx, true)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    worksheet = SetEnhancement(worksheet, _renewalPeriod, rowIndx,true)
    rowIndx = worksheet.LastRowNum as int
    rowIndx++

    return worksheet
  }

  private function GetOutput(workbook:XSSFWorkbook, year2RateChangeData:Year2RateChangeData) : XSSFWorkbook{

    //Populate Output DTO direcly from the Input values
    year2RateChangeData = MoveInputDirectlyToOutput(year2RateChangeData)

    //Populate Output DTO from Named Ranges in workbook
    var namedRanges = workbook.AllNames
    var namedRangesSheet : XSSFSheet

    namedRanges.each(\namedRange -> {
      var nrName : String
      var nrReference : String
      var nrValue : String

      try {
        nrName = namedRange.NameName
        nrReference = namedRange.RefersToFormula
      } catch (e:Exception) { }

      try {

        if(nrReference.length() > 0) {
          switch (nrName) {
            case "Restated_Expiring_Annual_D_O_Charged_Premium": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Restated_Expiring_Annual_D_O_Charged_Premium = nrValue} break
            case "Restated_Expiring_Annual_EPL_Charged_Premium": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) { year2RateChangeData.OutputData.Restated_Expiring_Annual_EPL_Charged_Premium = nrValue} break
            case "Restated_Expiring_Annual_Em_ee_Theft_Charged_Premium": nrValue = GetRefValue(workbook, nrReference) if (nrValue != ""&& isNumeric(nrValue)) {year2RateChangeData.OutputData.Restated_Expiring_Annual_Em_ee_Theft_Charged_Premium = nrValue} break
            case "Restated_Expiring_Annual_FLC_Charged_Premium": nrValue = GetRefValue(workbook, nrReference) if (nrValue != ""&& isNumeric(nrValue)) {year2RateChangeData.OutputData.Restated_Expiring_Annual_FID_Charged_Premium = nrValue} break
            case "Expiring_Actual_Charged": nrValue = GetRefValue(workbook, nrReference) if (nrValue != ""&& isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_Actual_Charged = nrValue} break
            case "Expiring_Total_Base_Premium": nrValue = GetRefValue(workbook, nrReference) if (nrValue != ""&& isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_Total_Base_Premium = nrValue} break
            case "Expiring_Deviation": nrValue = GetRefValue(workbook, nrReference) if (nrValue != ""&& isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_Deviation = nrValue} break

            case "Renewing_Actual_Charged": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewing_Actual_Charged = nrValue} break
            case "Renewing_Total_Base_Premium": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewing_Total_Base_Premium = nrValue} break
            case "Renewing_Deviation": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewing_Deviation = nrValue} break

            case "Crime_Rate_Change": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Crime_Rate_Change = nrValue} break
            case "D_O_Rate_Change": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.D_O_Rate_Change = nrValue} break
            case "EPL_Rate_Change": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.EPL_Rate_Change = nrValue} break
            case "FLC_Rate_Change": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.FID_Rate_Change = nrValue} break
            case "Total_Rate_Change": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Total_Rate_Change = nrValue} break

            case "Renewal_ChargedPremium_D_O": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_D_O = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_EPL": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_EPL = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Fiduciary": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Fiduciary = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Employee_Theft_Per_Loss_Coverage": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Employee_Theft_Per_Loss_Coverage = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Clients_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Clients_Property = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Forgery_or_Alteration_Checks_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Forgery_or_Alteration_Checks_Forgery = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Inside_Premises_Theft_of_Money_and_Securities": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Inside_Premises_Theft_of_Money_and_Securities = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Outside_The_Premises_In_Transit": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Outside_The_Premises_In_Transit = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Computer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Computer_Fraud = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Funds_Transfer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Funds_Transfer_Fraud = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Money_Orders_and_Counterfeit_Money": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Money_Orders_and_Counterfeit_Money = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Electronic_Data_or_Computer_Programs_Restoration_Costs": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Electronic_Data_or_Computer_Programs_Restoration_Costs = nrValue.toDouble()} break
            case "Renewal_ChargedPremium_Investigative_Expenses": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Investigative_Expenses = nrValue.toDouble()} break

            case "Renewal_BasePremium_D_O": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_D_O = nrValue.toDouble()} break
            case "Renewal_BasePremium_EPL": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_EPL = nrValue.toDouble()} break
            case "Renewal_BasePremium_Fiduciary": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Fiduciary = nrValue.toDouble()} break
            case "Renewal_BasePremium_Employee_Theft_Per_Loss_Coverage": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Employee_Theft_Per_Loss_Coverage = nrValue.toDouble()} break
            case "Renewal_BasePremium_Clients_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Clients_Property = nrValue.toDouble()} break
            case "Renewal_BasePremium_Forgery_or_Alteration_Checks_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Forgery_or_Alteration_Checks_Forgery = nrValue.toDouble()} break
            case "Renewal_BasePremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery = nrValue.toDouble()} break
            case "Renewal_BasePremium_Inside_Premises_Theft_of_Money_and_Securities": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Inside_Premises_Theft_of_Money_and_Securities = nrValue.toDouble()} break
            case "Renewal_BasePremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property = nrValue.toDouble()} break
            case "Renewal_BasePremium_Outside_The_Premises_In_Transit": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Outside_The_Premises_In_Transit = nrValue.toDouble()} break
            case "Renewal_BasePremium_Computer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Computer_Fraud = nrValue.toDouble()} break
            case "Renewal_BasePremium_Funds_Transfer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Funds_Transfer_Fraud = nrValue.toDouble()} break
            case "Renewal_BasePremium_Money_Orders_and_Counterfeit_Money": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Money_Orders_and_Counterfeit_Money = nrValue.toDouble()} break
            case "Renewal_BasePremium_Electronic_Data_or_Computer_Programs_Restoration_Costs": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Electronic_Data_or_Computer_Programs_Restoration_Costs = nrValue.toDouble()} break
            case "Renewal_BasePremium_Investigative_Expenses": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Investigative_Expenses = nrValue.toDouble()} break

            case "Expiring_ChargedPremium_D_O": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_D_O = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_EPL": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_EPL = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Fiduciary": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Fiduciary = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Employee_Theft_Per_Loss_Coverage": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Employee_Theft_Per_Loss_Coverage = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Clients_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Clients_Property = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Forgery_or_Alteration_Checks_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Forgery_or_Alteration_Checks_Forgery = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Inside_Premises_Theft_of_Money_and_Securities": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Inside_Premises_Theft_of_Money_and_Securities = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Outside_The_Premises_In_Transit": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Outside_The_Premises_In_Transit = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Computer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Computer_Fraud = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Funds_Transfer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Funds_Transfer_Fraud = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Money_Orders_and_Counterfeit_Money": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Money_Orders_and_Counterfeit_Money = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Electronic_Data_or_Computer_Programs_Restoration_Costs": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Electronic_Data_or_Computer_Programs_Restoration_Costs = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Investigative_Expenses": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Investigative_Expenses = nrValue.toDouble()} break

            case "Expiring_BasePremium_D_O": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_D_O = nrValue.toDouble()} break
            case "Expiring_BasePremium_EPL": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_EPL = nrValue.toDouble()} break
            case "Expiring_BasePremium_Fiduciary": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Fiduciary = nrValue.toDouble()} break
            case "Expiring_BasePremium_Employee_Theft_Per_Loss_Coverage": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Employee_Theft_Per_Loss_Coverage = nrValue.toDouble()} break
            case "Expiring_BasePremium_Clients_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Clients_Property = nrValue.toDouble()} break
            case "Expiring_BasePremium_Forgery_or_Alteration_Checks_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Forgery_or_Alteration_Checks_Forgery = nrValue.toDouble()} break
            case "Expiring_BasePremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Forgery_or_Alteration_Credit_Debit_or_Charge_Card_Forgery = nrValue.toDouble()} break
            case "Expiring_BasePremium_Inside_Premises_Theft_of_Money_and_Securities": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Inside_Premises_Theft_of_Money_and_Securities = nrValue.toDouble()} break
            case "Expiring_BasePremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Inside_Premises_Robbery_or_Safe_Burglary_of_Other_Property = nrValue.toDouble()} break
            case "Expiring_BasePremium_Outside_The_Premises_In_Transit": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Outside_The_Premises_In_Transit = nrValue.toDouble()} break
            case "Expiring_BasePremium_Computer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Computer_Fraud = nrValue.toDouble()} break
            case "Expiring_BasePremium_Funds_Transfer_Fraud": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Funds_Transfer_Fraud = nrValue.toDouble()} break
            case "Expiring_BasePremium_Money_Orders_and_Counterfeit_Money": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Money_Orders_and_Counterfeit_Money = nrValue.toDouble()} break
            case "Expiring_BasePremium_Electronic_Data_or_Computer_Programs_Restoration_Costs": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Electronic_Data_or_Computer_Programs_Restoration_Costs = nrValue.toDouble()} break
            case "Expiring_BasePremium_Investigative_Expenses": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Investigative_Expenses = nrValue.toDouble()} break

            case "Renewal_ChargedPremium_Total": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_ChargedPremium_Total = nrValue.toDouble()} break
            case "Renewal_BasePremium_Total": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Renewal_BasePremium_Total = nrValue.toDouble()} break
            case "Expiring_ChargedPremium_Total": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_ChargedPremium_Total = nrValue.toDouble()} break
            case "Expiring_BasePremium_Total": nrValue = GetRefValue(workbook, nrReference) if (nrValue != "" && isNumeric(nrValue)) {year2RateChangeData.OutputData.Expiring_BasePremium_Total = nrValue.toDouble()} break
            default:
              break
          }
        }

      } catch (e:Exception) {}

    })

    return workbook

  }

  private function MoveInputDirectlyToOutput (year2RateChangeData:Year2RateChangeData) : Year2RateChangeData {
    //Some of the input values can be moved directly to the output dto
    year2RateChangeData.OutputData.Expiring_Commission = (year2RateChangeData.InputData.Expiring_Commission < 0) ? 0 : year2RateChangeData.InputData.Expiring_Commission
    year2RateChangeData.OutputData.Renewing_Commission = (year2RateChangeData.InputData.Renewing_Commission < 0) ? 0 : year2RateChangeData.InputData.Renewing_Commission

    year2RateChangeData.OutputData.Expiring_TotalAssets_D_O = (year2RateChangeData.InputData.Expiring_TotalAssets < 0) ? 0 : year2RateChangeData.InputData.Expiring_TotalAssets
    year2RateChangeData.OutputData.ExpiringLiabilityLimit_D_O = (year2RateChangeData.InputData.Expiring_DO_Limit < 0) ? 0 : year2RateChangeData.InputData.Expiring_DO_Limit
    year2RateChangeData.OutputData.Expiring_SIR_D_O = (year2RateChangeData.InputData.Expiring_DO_SIR < 0) ? 0 : year2RateChangeData.InputData.Expiring_DO_SIR
    year2RateChangeData.OutputData.ExpiringRatableEmp_EPL = (year2RateChangeData.InputData.Expiring_EPL_RateableEmployees < 0) ? 0 : year2RateChangeData.InputData.Expiring_EPL_RateableEmployees
    year2RateChangeData.OutputData.ExpiringLiabilityLimit_EPL = (year2RateChangeData.InputData.Expiring_EPL_Limit < 0) ? 0 : year2RateChangeData.InputData.Expiring_EPL_Limit
    year2RateChangeData.OutputData.Expiring_SIR_EPL = (year2RateChangeData.InputData.Expiring_EPL_SIR < 0) ? 0 : year2RateChangeData.InputData.Expiring_EPL_SIR
    year2RateChangeData.OutputData.Expiring_PlanAssets_FID = (year2RateChangeData.InputData.Expiring_TotalPlanAssets < 0) ? 0 : year2RateChangeData.InputData.Expiring_TotalPlanAssets
    year2RateChangeData.OutputData.ExpiringLiabilityLimit_FID = (year2RateChangeData.InputData.Expiring_FID_Limit < 0) ? 0 : year2RateChangeData.InputData.Expiring_FID_Limit
    year2RateChangeData.OutputData.Expiring_SIR_FID = (year2RateChangeData.InputData.Expiring_FID_SIR < 0) ? 0 : year2RateChangeData.InputData.Expiring_FID_SIR
    year2RateChangeData.OutputData.ExpiringRatableEmp_CR = (year2RateChangeData.InputData.Expiring_CR_RateableEmployees < 0) ? 0 : year2RateChangeData.InputData.Expiring_CR_RateableEmployees

    year2RateChangeData.OutputData.Renewing_TotalAssets_D_O = (year2RateChangeData.InputData.Renewing_TotalAssets < 0) ? 0 : year2RateChangeData.InputData.Renewing_TotalAssets
    year2RateChangeData.OutputData.Renewing_LiabilityLimit_D_O = (year2RateChangeData.InputData.Renewing_DO_Limit < 0) ? 0 : year2RateChangeData.InputData.Renewing_DO_Limit
    year2RateChangeData.OutputData.Renewing_SIR_D_O = (year2RateChangeData.InputData.Renewing_DO_SIR < 0) ? 0 : year2RateChangeData.InputData.Renewing_DO_SIR
    year2RateChangeData.OutputData.RenewingRatableEmp_EPL = (year2RateChangeData.InputData.Renewing_EPL_RatableEmployees < 0) ? 0 : year2RateChangeData.InputData.Renewing_EPL_RatableEmployees
    year2RateChangeData.OutputData.RenewingLiabilityLimit_EPL = (year2RateChangeData.InputData.Renewing_EPL_Limit < 0) ? 0 : year2RateChangeData.InputData.Renewing_EPL_Limit
    year2RateChangeData.OutputData.Renewing_SIR_EPL = (year2RateChangeData.InputData.Renewing_EPL_SIR < 0) ? 0 : year2RateChangeData.InputData.Renewing_EPL_SIR
    year2RateChangeData.OutputData.Renewing_PlanAssets_FID = (year2RateChangeData.InputData.Renewing_PlanAssets < 0) ? 0 : year2RateChangeData.InputData.Renewing_PlanAssets
    year2RateChangeData.OutputData.RenewingLiabilityLimit_FID = (year2RateChangeData.InputData.Renewing_FID_Limit < 0) ? 0 : year2RateChangeData.InputData.Renewing_FID_Limit
    year2RateChangeData.OutputData.Renewing_SIR_FID = (year2RateChangeData.InputData.Renewing_FID_SIR < 0) ? 0 : year2RateChangeData.InputData.Renewing_FID_SIR
    year2RateChangeData.OutputData.RenewingRatableEmp_CR = (year2RateChangeData.InputData.Renewing_CR_RatableEmployees < 0) ? 0 : year2RateChangeData.InputData.Renewing_CR_RatableEmployees

    return year2RateChangeData
  }


  private function createRowWith2Cells(worksheet:XSSFSheet,  rowIndx: int, cellValueA:String, cellValueB:String) {
    var row = worksheet.createRow(rowIndx)
    row.createCell(0).setCellValue(cellValueA)
    row.createCell(1).setCellValue(cellValueB)
  }

  private function createRowWith2Cells(worksheet:XSSFSheet,  rowIndx: int, cellValueA:String, cellValueB:Double) {
    var row = worksheet.createRow(rowIndx)
    row.createCell(0).setCellValue(cellValueA)
    row.createCell(1).setCellValue(cellValueB)
  }

  private function createRowWith3Cells(worksheet:XSSFSheet,  rowIndx: int, cellValueA:String, cellValueB:Double, cellValueC:String) {
    var row = worksheet.createRow(rowIndx)
    row.createCell(0).setCellValue(cellValueA)
    row.createCell(1).setCellValue(cellValueB)
    row.createCell(2).setCellValue(cellValueC)
  }

  private function GetRefValue(workbook:XSSFWorkbook, ref:String) : String {
    //Cell reference example : 'RateChange_Year 1'!$F$90
    //Check for valid reference - If the cell reference starts with # then it is not a valid reference
    if (ref == null || ref == "") {
      return ""
    }
    if (ref.substring(0,1) == "#") {
      return ""
    }

    var evaluator = workbook.getCreationHelper().createFormulaEvaluator()
    var formatter = new DataFormatter()
    var cellValue : String
    var worksheetName = ""
    var cellRef = ""
    var rowNumber = ""
    var cellNumber = ""
    var rowNo : int

    try {
      var refArray = new String[0]
      refArray = ref.split("!")
      worksheetName = refArray[0].replace("'", "")
      cellRef = refArray[1]
      cellRef = cellRef.substring(1,cellRef.length()).replace("$",",")  //replace "$" because split function doesn't work with it

      var cellArray = new String[0]
      cellArray = cellRef.split(",")
      cellNumber = cellArray[0]
      rowNumber = cellArray[1]
      rowNo = rowNumber.toInt()
      rowNo-- //subtract 1 because the "getRow()" function adds 1

      var sheet = workbook.getSheet(worksheetName)
      var row = sheet.getRow(rowNo)
      var cell = row.getCell(CellReference.convertColStringToIndex(cellNumber), RETURN_NULL_AND_BLANK)

      cellValue = formatter.formatCellValue(cell,evaluator)

      //format numeric values - remove commas, convert percentage
      if (cellValue.length() > 0){
        cellValue = cellValue.replace(",","")

        if(cellValue.indexOf("%") > 0 ){
          cellValue = cellValue.replace("%","")
          var percent = cellValue.toDouble()/100
          cellValue = percent as String
        }

      }

    }
    catch (e:Exception) {  }

    if (cellValue == null) { cellValue = ""}
    if ((cellValue == "- 0") || (cellValue == "-0")) { cellValue = "0"}

    return cellValue
  }


  private function constructFileName(period : PolicyPeriod, workbookName : String) : String {
    var dateFormat = new SimpleDateFormat("yyyyMMdd-HHmm")
    var timeStamp = dateFormat.format(Date.CurrentDate)
    var fileSuffix = "_" + timeStamp
    var jobTypeName = period.Job.Subtype.DisplayName
    var jobID = period.Job.JobNumber
    var policyId = period.PolicyNumberAssigned ? period.PolicyNumber : null

    var fileName : String
    if (policyId != null) {
      fileName = workbookName + "_" + policyId + "_" + jobTypeName + "_" + jobID + fileSuffix
    } else {
      fileName = workbookName + "_" + jobTypeName + "_" + jobID + fileSuffix
    }

    return filterInvalidFilenameCharacters(fileName + ".xlsx")
  }

  

  private function isNumeric(str : String) : Boolean {
    if (str == null) {
      return false
    }
    try {
      var dbl = Double.parseDouble(str)
    }
    catch (e : Exception){
      return false
    }
    return true
  }

}