using System;
using System.Data;
using System.Data.SqlClient;
using System.Web.Configuration;

/// <summary>
/// GBCJSONRecordDAO 的摘要描述
/// </summary>
public class GBCJSONRecordDAO : GBCJSONRecord_Interface
{
    private const string DELETE_VOU_JSON1_STMT =
        "delete from  GBCJSONRecord where 基金代碼=@基金代碼 and  PFK_會計年度=@PFK_會計年度 and  PFK_動支編號=@PFK_動支編號 and PFK_種類=@PFK_種類 and PFK_次別=@PFK_次別";

    private const string FIND_JSON1_STMT =
        "select 傳票JSON1 from GBCJSONRecord where 基金代碼=@基金代碼 and PFK_會計年度=@PFK_會計年度 and PFK_動支編號=@PFK_動支編號 and PFK_種類=@PFK_種類 and PFK_次別=@PFK_次別";

    private const string FIND_JSON2_STMT =
        "select 傳票JSON2 from GBCJSONRecord where 基金代碼=@基金代碼 and PFK_會計年度=@PFK_會計年度 and PFK_動支編號=@PFK_動支編號 and PFK_種類=@PFK_種類 and PFK_次別=@PFK_次別";

    private const string INSERT_VOU_JSON1_STMT =
        "insert into GBCJSONRecord(基金代碼, PFK_會計年度, PFK_動支編號, PFK_種類, PFK_次別, 傳票JSON1,是否結案) values(@基金代碼,@PFK_會計年度,@PFK_動支編號,@PFK_種類,@PFK_次別,@傳票JSON1,'0')";

    private const string IS_PASS_STMT =
        "select isNull(是否結案,'') as 是否結案 from GBCJSONRecord where 基金代碼=@基金代碼 and PFK_會計年度=@PFK_會計年度 and PFK_動支編號=@PFK_動支編號 and PFK_種類=@PFK_種類 and PFK_次別=@PFK_次別";

    private const string UPDATE_PASS_STMT =
        "update GBCJSONRecord set 是否結案 = '1' where 基金代碼=@基金代碼 and PFK_會計年度=@PFK_會計年度 and PFK_動支編號=@PFK_動支編號 and PFK_種類=@PFK_種類 and PFK_次別=@PFK_次別";

    private const string UPDATE_VOU_JSON2_STMT =
        "update GBCJSONRecord set 傳票JSON2 = @傳票JSON2 where 基金代碼=@基金代碼 and  PFK_會計年度=@PFK_會計年度 and  PFK_動支編號=@PFK_動支編號 and PFK_種類=@PFK_種類 and PFK_次別=@PFK_次別";

    private const string INSERT_JSON_LOG_STMT =
        "insert into GBCJSONRecordLog (基金代碼,條碼,JSON紀錄,接收時間) values(@基金代碼,@條碼,@JSON紀錄,@接收時間)";

    
    /// <summary>
    /// 移除傳票底稿1
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    public void DeleteJsonRecord1(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        SqlConnection conn = null;
        SqlCommand com = null;

        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        conn.Open();
        com = new SqlCommand(DELETE_VOU_JSON1_STMT, conn);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PFK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PFK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PFK_種類", vw_GBCVisaDetail.PK_種類);
        com.Parameters.AddWithValue("@PFK_次別", vw_GBCVisaDetail.PK_次別);

        com.CommandType = CommandType.Text;
        com.ExecuteNonQuery();

        conn.Close();
    }

    /// <summary>
    /// 找JSON1資料
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <returns></returns>
    public string FindJSON1(Vw_GBCVisaDetail vw_GBCVisaDetail)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(FIND_JSON1_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PFK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PFK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PFK_種類", vw_GBCVisaDetail.PK_種類);
        com.Parameters.AddWithValue("@PFK_次別", vw_GBCVisaDetail.PK_次別);
        con.Open();

        string JSON1 = com.ExecuteScalar().ToString();

        con.Close();

        return JSON1;
    }

    /// <summary>
    /// 找JSON2資料
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <returns></returns>
    public string FindJSON2(string 基金代碼, string 會計年度, string 動支編號, string 種類, string 次別)
    {
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(FIND_JSON2_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", 基金代碼);
        com.Parameters.AddWithValue("@PFK_會計年度", 會計年度);
        com.Parameters.AddWithValue("@PFK_動支編號", 動支編號);
        com.Parameters.AddWithValue("@PFK_種類", 種類);
        com.Parameters.AddWithValue("@PFK_次別", 次別);
        con.Open();

        string JSON2 = com.ExecuteScalar().ToString();

        con.Close();

        return JSON2;
    }

    /// <summary>
    /// 寫入傳票底稿1
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    public void InsertJsonRecord1(Vw_GBCVisaDetail vw_GBCVisaDetail, string vouJoson)
    {
        SqlConnection conn = null;
        SqlCommand com = null;
        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlTransaction transtation;

        conn.Open();
        transtation = conn.BeginTransaction();
        //com.Transaction = transtation;

        try
        {
            
            com = new SqlCommand(INSERT_VOU_JSON1_STMT, conn, transtation);
            com.Parameters.AddWithValue("@傳票JSON1", vouJoson);
            com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
            com.Parameters.AddWithValue("@PFK_會計年度", vw_GBCVisaDetail.PK_會計年度);
            com.Parameters.AddWithValue("@PFK_動支編號", vw_GBCVisaDetail.PK_動支編號);
            com.Parameters.AddWithValue("@PFK_種類", vw_GBCVisaDetail.PK_種類);
            com.Parameters.AddWithValue("@PFK_次別", vw_GBCVisaDetail.PK_次別);

            com.CommandType = CommandType.Text;
            com.ExecuteNonQuery();

            transtation.Commit();

            conn.Close();
        }
        catch (Exception)
        {
            transtation.Rollback();
            conn.Close();

            throw;
        }

    }

    /// <summary>
    /// JSON紀錄是否結案
    /// </summary>
    /// <param name="基金代碼"></param>
    /// <param name="會計年度"></param>
    /// <param name="動支編號"></param>
    /// <param name="種類"></param>
    /// <param name="次別"></param>
    /// <returns></returns>
    public string IsPass(string 基金代碼, string 會計年度, string 動支編號, string 種類, string 次別)
    {
        string isPass = "";
        SqlConnection con = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlCommand com = new SqlCommand(IS_PASS_STMT, con);
        com.Parameters.AddWithValue("@基金代碼", 基金代碼);
        com.Parameters.AddWithValue("@PFK_會計年度", 會計年度);
        com.Parameters.AddWithValue("@PFK_動支編號", 動支編號);
        com.Parameters.AddWithValue("@PFK_種類", 種類);
        com.Parameters.AddWithValue("@PFK_次別", 次別);
        con.Open();
        try
        {
            isPass = com.ExecuteScalar().ToString();
            if (isPass.Length == 0)
            {
                isPass = "0";
            }
        }
        catch (Exception)
        {
            isPass = "0";
        }

        con.Close();

        return isPass;
    }

    /// <summary>
    /// JSON接收LOG
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    /// <param name="vouJoson"></param>
    public void InsertJsonLog(string fundNo, string acmWordNum, string vouJoson)
    {
        SqlConnection conn = null;
        SqlCommand com = null;
        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlTransaction transtation;

        conn.Open();
        transtation = conn.BeginTransaction();
        //com.Transaction = transtation;

        try
        {

            com = new SqlCommand(INSERT_JSON_LOG_STMT, conn, transtation);
            com.Parameters.AddWithValue("@基金代碼", fundNo);
            com.Parameters.AddWithValue("@條碼", acmWordNum);
            com.Parameters.AddWithValue("@JSON紀錄", vouJoson);
            com.Parameters.AddWithValue("@接收時間", DateTime.Now);

            com.CommandType = CommandType.Text;
            com.ExecuteNonQuery();

            transtation.Commit();

            conn.Close();
        }
        catch (Exception)
        {
            transtation.Rollback();
            conn.Close();

            throw;
        }

    }

    /// <summary>
    /// 寫入傳票底稿2
    /// </summary>
    /// <param name="vw_GBCVisaDetail"></param>
    public void UpdateJsonRecord2(Vw_GBCVisaDetail vw_GBCVisaDetail, string vouJoson)
    {
        SqlConnection conn = null;
        SqlCommand com = null;

        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        conn.Open();
        com = new SqlCommand(UPDATE_VOU_JSON2_STMT, conn);
        com.Parameters.AddWithValue("@傳票JSON2", vouJoson);
        com.Parameters.AddWithValue("@基金代碼", vw_GBCVisaDetail.基金代碼);
        com.Parameters.AddWithValue("@PFK_會計年度", vw_GBCVisaDetail.PK_會計年度);
        com.Parameters.AddWithValue("@PFK_動支編號", vw_GBCVisaDetail.PK_動支編號);
        com.Parameters.AddWithValue("@PFK_種類", vw_GBCVisaDetail.PK_種類);
        com.Parameters.AddWithValue("@PFK_次別", vw_GBCVisaDetail.PK_次別);

        com.CommandType = CommandType.Text;
        com.ExecuteNonQuery();

        conn.Close();
    }

    /// <summary>
    /// 標上結案旗標
    /// </summary>
    /// <param name="基金代碼"></param>
    /// <param name="會計年度"></param>
    /// <param name="動支編號"></param>
    /// <param name="種類"></param>
    /// <param name="次別"></param>
    public void UpdatePassFlg(string 基金代碼, string 會計年度, string 動支編號, string 種類, string 次別)
    {
        SqlConnection conn = null;
        SqlCommand com = null;

        conn = new SqlConnection(WebConfigurationManager.ConnectionStrings["SqlDbConnStr"].ConnectionString);
        SqlTransaction transtation;
        conn.Open();
        transtation = conn.BeginTransaction();
        //com.Transaction = transtation;
        try
        {            
            com = new SqlCommand(UPDATE_PASS_STMT, conn, transtation);

            com.Parameters.AddWithValue("@基金代碼", 基金代碼);
            com.Parameters.AddWithValue("@PFK_會計年度", 會計年度);
            com.Parameters.AddWithValue("@PFK_動支編號", 動支編號);
            com.Parameters.AddWithValue("@PFK_種類", 種類);
            com.Parameters.AddWithValue("@PFK_次別", 次別);

            com.CommandType = CommandType.Text;
            com.ExecuteNonQuery();

            transtation.Commit();

            conn.Close();
        }
        catch (Exception)
        {
            transtation.Rollback();
            conn.Close();
            throw;
        }

    }
}

