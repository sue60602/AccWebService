using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Web.Configuration;

/// <summary>
/// GBCVisaDetailAbateDetailDAO 的摘要描述
/// </summary>
public class GBCVisaDetailAbateDetailDAO : GBCVisaDetailAbateDetail_Interface
{
    private const string ESTIMATE_MONEY_ABATE_STMT =
        "select ISNULL(sum(Amount),0) as 估列沖銷數 from VouDetail where FundNo=@基金代碼 and AccYear=@PK_會計年度 and DC='借' and SubNo='2125' and RelatedVouNo = @F_傳票號碼 + '-' + @明細號";

    private const string ESTIMATE_MONEY_STMT =
        "select ISNULL(sum(Amount),0) as 已估列數 from VouDetail where FundNo=@基金代碼 and AccYear=@PK_會計年度 and DC='貸' and SubNo='2125' and AccMainNo=ISNULL((select AccMainno from VouMain where FundNo=@基金代碼 and AccYear=@PK_會計年度 and VouNo=@F_傳票號碼),'0')";

    private const string FILL_VOU_1_STMT =
        "update GBCVisaDetailAbateDetail set F_傳票年度=@傳票年度, F_傳票號1=@傳票號1, F_傳票明細號1=@傳票明細號1, F_製票日期1=@製票日期1 where 基金代碼=@基金代碼 and  PK_會計年度=@PK_會計年度 and  PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and PK_次別=@PK_次別 and PK_明細號=@PK_明細號";

    private const string FILL_VOU_2_STMT =
        //"update GBCVisaDetailAbateDetail set F_傳票年度=@傳票年度, F_傳票號2=@傳票號2, F_傳票明細號2=@傳票明細號2, F_製票日期2=@製票日期2 where 基金代碼=@基金代碼 and  PK_會計年度=@PK_會計年度 and  PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and PK_次別=@PK_次別 and PK_明細號=@PK_明細號";
        "update GBCVisaDetailAbateDetail set F_傳票年度=@傳票年度, F_傳票號2=@傳票號1, F_傳票明細號2=@傳票明細號1, F_製票日期2=@製票日期1 where 基金代碼=@基金代碼 and  PK_會計年度=@PK_會計年度 and  PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and PK_次別=@PK_次別 and PK_明細號=@PK_明細號";

    private const string FIND_LOG_STMT =
        "select count(*) from GBCVisaDetailAbateDetail where 基金代碼=@基金代碼 and  PK_會計年度=@PK_會計年度 and  PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and PK_次別=@PK_次別 and PK_明細號=@PK_明細號";

    private const string FIND_PREPAY_STMT =
        //"select PK_會計年度, PK_動支編號, PK_種類, PK_次別, PK_明細號, F_核定金額, F_傳票年度, F_傳票號1, F_傳票明細號1, F_製票日期1, F_傳票號2, F_傳票明細號2, F_製票日期2,F_受款人,F_受款人編號 from GBCVisaDetailAbateDetail where PK_會計年度=@PK_會計年度 and PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and F_受款人編號=@F_受款人編號";
        "select count(*) from GBCVisaDetailAbateDetail where 基金代碼=@基金代碼 and PK_會計年度=@PK_會計年度 and PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and F_受款人編號=@F_受款人編號";

    private const string FIND_VOUNO =
        "select F_傳票號1,F_傳票明細號1 FROM [GBCVisaDetailAbateDetail] where 基金代碼=@基金代碼 and PK_會計年度=@PK_會計年度 and PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and F_受款人編號=@F_受款人編號";

    private const string INSERT_STMT =
        //"INSERT INTO GBCVisaDetailAbateDetail ([PK_會計年度],[PK_動支編號],[PK_種類],[PK_次別],[PK_明細號],[F_核定金額],F_傳票種類,F_傳票號1,F_傳票明細號1,F_製票日期1,F_傳票號2,F_傳票明細號2,F_製票日期2) values(@PK_會計年度,@PK_動支編號,@PK_種類,@PK_次別,@PK_明細號,@F_核定金額,@F_傳票種類,@F_傳票號1,@F_傳票明細號1,@F_製票日期1,@F_傳票號2,@F_傳票明細號2,@F_製票日期2)";
        "INSERT INTO GBCVisaDetailAbateDetail ([基金代碼],[PK_會計年度],[PK_動支編號],[PK_種類],[PK_次別],[PK_明細號],[F_核定金額],[F_受款人],[F_受款人編號]) values(@基金代碼,@PK_會計年度,@PK_動支編號,@PK_種類,@PK_次別,@PK_明細號,@F_核定金額,@F_受款人,@F_受款人編號)";

    private const string IS_VOUNO_STMT =
        "select ISNULL(F_傳票號1,'') AS F_傳票號1 FROM [GBCVisaDetailAbateDetail] where 基金代碼=@基金代碼 and PK_會計年度=@PK_會計年度 and PK_動支編號=@PK_動支編號 and PK_種類=@PK_種類 and PK_次別=@PK_次別 and PK_明細號=@PK_明細號";

    private const string PREPAY_MONEY_ABATE_STMT =
        "select ISNULL(sum(Amount),0) as 預付轉正數 from VouDetail where FundNo=@基金代碼 and AccYear=@PK_會計年度 and DC='貸' and SubNo='1154' and RelatedVouNo = @F_傳票號碼 + '-' + @明細號";

    private const string PREPAY_MONEY_STMT =
        "select ISNULL(sum(Amount),0) as 已預付數 from VouDetail where FundNo=@基金代碼 and AccYear=@PK_會計年度 and DC='借' and SubNo='1154' and AccMainNo=ISNULL((select AccMainno from VouMain where FundNo=@基金代碼 and AccYear=@PK_會計年度 and VouNo=@F_傳票號碼),'0')";

    private const string UPDATE_STMT =
        //"INSERT INTO GBCVisaDetailAbateDetail ([PK_會計年度],[PK_動支編號],[PK_種類],[PK_次別],[PK_明細號],[F_核定金額],F_傳票種類,F_傳票號1,F_傳票明細號1,F_製票日期1,F_傳票號2,F_傳票明細號2,F_製票日期2) values(@PK_會計年度,@PK_動支編號,@PK_種類,@PK_次別,@PK_明細號,@F_核定金額,@F_傳票種類,@F_傳票號1,@F_傳票明細號1,@F_製票日期1,@F_傳票號2,@F_傳票明細號2,@F_製票日期2)";
        "UPDATE GBCVisaDetailAbateDetail set [基金代碼] = @基金代碼,[PK_會計年度] = @PK_會計年度,[PK_動支編號] = @PK_動支編號,[PK_種類] = @PK_種類,[PK_次別] = @PK_次別,[PK_明細號] = @PK_明細號,[F_核定金額] = @F_核定金額,[F_受款人] = @F_受款人,[F_受款人編號] = @F_受款人編號 where 基金代碼 = @基金代碼 and PK_會計年度 = @PK_會計年度 and PK_動支編號 = @PK_動支編號 and PK_種類 = @PK_種類 and PK_次別 = @PK_次別 and PK_明細號 = @PK_明細號";

    /// <summary>
    /// 計算已估列數
    /// </summary>
    /// <param name="accYear"></param>
    /// <param name="vouNo"></param>
    /// <returns></returns>
    public int EstimateMoney(string accNo, string accYear, string vouNo)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(ESTIMATE_MONEY_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", accNo);
        com.Parameters.AddWithValue("@PK_會計年度", accYear);
        com.Parameters.AddWithValue("@F_傳票號碼", vouNo);
        con.Open();
        int estimateMoney = int.Parse(com.ExecuteScalar().ToString());
        con.Close();

        return estimateMoney;
    }

    /// <summary>
    /// 計算估列已沖銷數
    /// </summary>
    /// <param name="accYear"></param>
    /// <param name="vouNo"></param>
    /// <param name="vouNoDtl"></param>
    /// <returns></returns>
    public int EstimateMoneyAbate(string accNo, string accYear, string vouNo, string vouNoDtl)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(ESTIMATE_MONEY_ABATE_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", accNo);
        com.Parameters.AddWithValue("@PK_會計年度", accYear);
        com.Parameters.AddWithValue("@F_傳票號碼", vouNo);
        com.Parameters.AddWithValue("@明細號", vouNoDtl);
        con.Open();
        int estimateMoneyAbate = int.Parse(com.ExecuteScalar().ToString());
        con.Close();

        return estimateMoneyAbate;
    }

    /// <summary>
    /// 尋找估列傳票號
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <returns></returns>
    public List<VouDetailVO> FindEstimateVouNo(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        List<VouDetailVO> vouNoList = new List<VouDetailVO>();
        VouDetailVO vouDetailVO = null;
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(FIND_VOUNO, con);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PK_種類", "估列");
        com.Parameters.AddWithValue("@F_受款人編號", vw_GBCVisaDetail.F_受款人編號);
        con.Open();

        SqlDataReader dr = com.ExecuteReader();
        while (dr.Read())
        {
            vouDetailVO = new VouDetailVO();
            vouDetailVO.傳票號 = dr["F_傳票號1"].ToString();
            vouDetailVO.傳票明細號 = dr["F_傳票明細號1"].ToString();
            vouNoList.Add(vouDetailVO);
        }

        con.Close();
        return vouNoList;
    }

    /// <summary>
    /// 尋找有無寫入紀錄
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <returns></returns>
    public int FindLog(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(FIND_LOG_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PK_種類", vw_GBCVisaDetail.PK_種類);
        com.Parameters.AddWithValue("@PK_次別", vw_GBCVisaDetail.PK_次別);
        com.Parameters.AddWithValue("@PK_明細號", vw_GBCVisaDetail.PK_明細號);
        con.Open();

        int prePayCnt = int.Parse(com.ExecuteScalar().ToString());

        con.Close();

        return prePayCnt;
    }

    /// <summary>
    /// 尋找有無預付紀錄
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <returns></returns>
    public int FindPrePay(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(FIND_PREPAY_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PK_種類", "預付");
        com.Parameters.AddWithValue("@F_受款人編號", vw_GBCVisaDetail.F_受款人編號);
        con.Open();

        int prePayCnt = int.Parse(com.ExecuteScalar().ToString());

        con.Close();

        return prePayCnt;
    }

    /// <summary>
    /// 尋找預付傳票號
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <returns></returns>
    public List<VouDetailVO> FindPrePayVouNo(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        List<VouDetailVO> vouNoList = new List<VouDetailVO>();
        VouDetailVO vouDetailVO = null;
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(FIND_VOUNO, con);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PK_種類", "預付");
        com.Parameters.AddWithValue("@F_受款人編號", vw_GBCVisaDetail.F_受款人編號);
        con.Open();

        SqlDataReader dr = com.ExecuteReader();
        while (dr.Read())
        {
            vouDetailVO = new VouDetailVO();
            vouDetailVO.傳票號 = dr["F_傳票號1"].ToString();
            vouDetailVO.傳票明細號 = dr["F_傳票明細號1"].ToString();
            vouNoList.Add(vouDetailVO);
        }

        con.Close();
        return vouNoList;
    }

    /// <summary>
    /// 尋找傳票號
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <returns></returns>
    public string FindVouNo(string 基金代碼, string 會計年度, string 動支編號, string 種類, string 次別, string 明細號)
    {
        string vouNo = "";
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(IS_VOUNO_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", 基金代碼);
        com.Parameters.AddWithValue("@PK_會計年度", 會計年度);
        com.Parameters.AddWithValue("@PK_動支編號", 動支編號);
        com.Parameters.AddWithValue("@PK_種類", 種類);
        com.Parameters.AddWithValue("@PK_次別", 次別);
        com.Parameters.AddWithValue("@PK_明細號", 明細號);
        con.Open();

        SqlDataReader dr = com.ExecuteReader();
        while (dr.Read())
        {
            vouNo = dr["F_傳票號1"].ToString();
        }

        return vouNo;
    }

    /// <summary>
    /// 寫入沖銷對應紀錄表
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    public void Insert(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        SqlConnection conn = null;
        SqlCommand com = null;

        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        conn.Open();
        com = new SqlCommand(INSERT_STMT, conn);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PK_種類", vw_GBCVisaDetail.PK_種類);
        com.Parameters.AddWithValue("@PK_次別", vw_GBCVisaDetail.PK_次別);
        com.Parameters.AddWithValue("@PK_明細號", vw_GBCVisaDetail.PK_明細號);
        com.Parameters.AddWithValue("@F_核定金額", vw_GBCVisaDetail.F_核定金額);
        com.Parameters.AddWithValue("@F_受款人編號", vw_GBCVisaDetail.F_受款人編號);
        com.Parameters.AddWithValue("@F_受款人", vw_GBCVisaDetail.F_受款人);

        com.CommandType = CommandType.Text;
        com.ExecuteNonQuery();

        conn.Close();
    }

    /// <summary>
    /// 計算已預付數
    /// </summary>
    /// <param name="accYear"></param>
    /// <param name="vouNo"></param>
    /// <param name="vouNoDtl"></param>
    /// <returns></returns>
    public int PrePayMoney(string accNo, string accYear, string vouNo)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(PREPAY_MONEY_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", accNo);
        com.Parameters.AddWithValue("@PK_會計年度", accYear);
        com.Parameters.AddWithValue("@F_傳票號碼", vouNo);
        con.Open();
        int prePayMoney = int.Parse(com.ExecuteScalar().ToString());
        con.Close();

        return prePayMoney;
    }

    /// <summary>
    /// 計算預付已轉正數
    /// </summary>
    /// <param name="accYear"></param>
    /// <param name="vouNo"></param>
    /// <param name="vouNoDtl"></param>
    /// <returns></returns>
    public int PrePayMoneyAbate(string accNo, string accYear, string vouNo, string vouNoDtl)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(PREPAY_MONEY_ABATE_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", accNo);
        com.Parameters.AddWithValue("@PK_會計年度", accYear);
        com.Parameters.AddWithValue("@F_傳票號碼", vouNo);
        com.Parameters.AddWithValue("@明細號", vouNoDtl);
        con.Open();
        int abatePrePayMoney = int.Parse(com.ExecuteScalar().ToString());
        con.Close();

        return abatePrePayMoney;
    }

    /// <summary>
    /// 更新沖銷對應紀錄表
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    public void Update(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        SqlConnection conn = null;
        SqlCommand com = null;

        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        conn.Open();
        com = new SqlCommand(UPDATE_STMT, conn);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PK_種類", vw_GBCVisaDetail.PK_種類);
        com.Parameters.AddWithValue("@PK_次別", vw_GBCVisaDetail.PK_次別);
        com.Parameters.AddWithValue("@PK_明細號", vw_GBCVisaDetail.PK_明細號);
        com.Parameters.AddWithValue("@F_核定金額", vw_GBCVisaDetail.F_核定金額);
        com.Parameters.AddWithValue("@F_受款人編號", vw_GBCVisaDetail.F_受款人編號);
        com.Parameters.AddWithValue("@F_受款人", vw_GBCVisaDetail.F_受款人);

        com.CommandType = CommandType.Text;
        com.ExecuteNonQuery();

        conn.Close();
    }

    /// <summary>
    /// 回填傳票號1
    /// </summary>
    /// <param name="gbcVisaDetailAbateDetailVO"></param>
    public void UpdateVouNo1(GBCVisaDetailAbateDetailVO gbcVisaDetailAbateDetailVO)
    {
        SqlConnection conn = null;
        SqlCommand com = null;

        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        conn.Open();
        com = new SqlCommand(FILL_VOU_1_STMT, conn);

        com.Parameters.AddWithValue("@基金代碼", gbcVisaDetailAbateDetailVO.get基金代碼());
        com.Parameters.AddWithValue("@PK_會計年度", gbcVisaDetailAbateDetailVO.getPK_會計年度());
        com.Parameters.AddWithValue("@PK_動支編號", gbcVisaDetailAbateDetailVO.getPK_動支編號());
        com.Parameters.AddWithValue("@PK_種類", gbcVisaDetailAbateDetailVO.getPK_種類());
        com.Parameters.AddWithValue("@PK_次別", gbcVisaDetailAbateDetailVO.getPK_次別());
        com.Parameters.AddWithValue("@PK_明細號", gbcVisaDetailAbateDetailVO.getPK_明細號());
        com.Parameters.AddWithValue("@傳票年度", gbcVisaDetailAbateDetailVO.getF_傳票年度());
        com.Parameters.AddWithValue("@傳票號1", gbcVisaDetailAbateDetailVO.getF_傳票號1());
        com.Parameters.AddWithValue("@傳票明細號1", gbcVisaDetailAbateDetailVO.getF_傳票明細號1());
        com.Parameters.AddWithValue("@製票日期1", gbcVisaDetailAbateDetailVO.getF_製票日期1());
        //com.Parameters.AddWithValue("@傳票號2", gbcVisaDetailAbateDetailVO.getF_傳票號2());
        //com.Parameters.AddWithValue("@傳票明細號2", gbcVisaDetailAbateDetailVO.getF_傳票明細號2());
        //com.Parameters.AddWithValue("@製票日期2", gbcVisaDetailAbateDetailVO.getF_製票日期2());

        com.CommandType = CommandType.Text;
        com.ExecuteNonQuery();

        conn.Close();
    }

    /// <summary>
    /// 回填傳票號2
    /// </summary>
    /// <param name="gbcVisaDetailAbateDetailVO"></param>
    public void UpdateVouNo2(GBCVisaDetailAbateDetailVO gbcVisaDetailAbateDetailVO)
    {
        SqlConnection conn = null;
        SqlCommand com = null;

        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        conn.Open();
        com = new SqlCommand(FILL_VOU_2_STMT, conn);

        com.Parameters.AddWithValue("@基金代碼", gbcVisaDetailAbateDetailVO.get基金代碼());
        com.Parameters.AddWithValue("@PK_會計年度", gbcVisaDetailAbateDetailVO.getPK_會計年度());
        com.Parameters.AddWithValue("@PK_動支編號", gbcVisaDetailAbateDetailVO.getPK_動支編號());
        com.Parameters.AddWithValue("@PK_種類", gbcVisaDetailAbateDetailVO.getPK_種類());
        com.Parameters.AddWithValue("@PK_次別", gbcVisaDetailAbateDetailVO.getPK_次別());
        com.Parameters.AddWithValue("@PK_明細號", gbcVisaDetailAbateDetailVO.getPK_明細號());
        com.Parameters.AddWithValue("@傳票年度", gbcVisaDetailAbateDetailVO.getF_傳票年度());
        com.Parameters.AddWithValue("@傳票號1", gbcVisaDetailAbateDetailVO.getF_傳票號1());
        com.Parameters.AddWithValue("@傳票明細號1", gbcVisaDetailAbateDetailVO.getF_傳票明細號1());
        com.Parameters.AddWithValue("@製票日期1", gbcVisaDetailAbateDetailVO.getF_製票日期1());
        //com.Parameters.AddWithValue("@傳票號2", gbcVisaDetailAbateDetailVO.getF_傳票號2());
        //com.Parameters.AddWithValue("@傳票明細號2", gbcVisaDetailAbateDetailVO.getF_傳票明細號2());
        //com.Parameters.AddWithValue("@製票日期2", gbcVisaDetailAbateDetailVO.getF_製票日期2());

        com.CommandType = CommandType.Text;
        com.ExecuteNonQuery();

        conn.Close();
    }
}