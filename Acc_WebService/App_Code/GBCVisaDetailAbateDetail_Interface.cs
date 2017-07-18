using System.Collections.Generic;

/// <summary>
/// GBCJSONRecord 的摘要描述
/// </summary>
public interface GBCVisaDetailAbateDetail_Interface
{
    int EstimateMoney(string accNo, string accYear, string vouNo, string vouNoDtl);

    int EstimateMoneyAbate(string accNo, string accYear, string vouNo, string vouNoDtl);

    List<VouDetailVO> FindEstimateVouNo(Vw_GBCVisaDetail vw_GBCVisaDetail);

    int FindLog(Vw_GBCVisaDetail vw_GBCVisaDetail);

    int FindPrePay(Vw_GBCVisaDetail vw_GBCVisaDetail);

    List<VouDetailVO> FindPrePayVouNo(Vw_GBCVisaDetail vw_GBCVisaDetail);

    void Insert(Vw_GBCVisaDetail vw_GBCVisaDetail);

    int PrePayMoney(string accNo, string accYear, string vouNo, string vouNoDtl);

    int PrePayMoneyAbate(string accNo, string accYear, string vouNo, string vouNoDtl);
}