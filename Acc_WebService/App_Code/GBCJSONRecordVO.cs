/// <summary>
/// Vw_GBCJSONRecordVO 的摘要描述
/// </summary>

public class GBCJSONRecordVO
{
    private string PFK_次別;
    private string PFK_動支編號;
    private string PFK_會計年度;
    private string PFK_種類;
    private string 是否結案;
    private string 基金代碼;
    private string 傳票JSON1;
    private string 傳票JSON2;

    public string getPFK_次別()
    {
        return PFK_次別;
    }

    public string getPFK_會計年度()
    {
        return PFK_會計年度;
    }

    public string getPFK_種類()
    {
        return PFK_種類;
    }

    public string getPK_動支編號()
    {
        return PFK_動支編號;
    }

    public string get是否結案()
    {
        return 是否結案;
    }

    public string get基金代碼()
    {
        return 基金代碼;
    }

    public string get傳票JSON1()
    {
        return 傳票JSON1;
    }

    public string get傳票JSON2()
    {
        return 傳票JSON2;
    }

    public void setPFK_次別(string pFK_次別)
    {
        PFK_次別 = pFK_次別;
    }

    public void setPFK_會計年度(string pFK_會計年度)
    {
        PFK_會計年度 = pFK_會計年度;
    }

    public void setPFK_種類(string pFK_種類)
    {
        PFK_種類 = pFK_種類;
    }

    public void setPK_動支編號(string pFK_動支編號)
    {
        PFK_動支編號 = pFK_動支編號;
    }

    public void set是否結案(string 是否結案)
    {
        this.是否結案 = 是否結案;
    }

    public void set基金代碼(string 基金代碼)
    {
        this.基金代碼 = 基金代碼;
    }

    public void set傳票JSON1(string 傳票JSON1)
    {
        this.傳票JSON1 = 傳票JSON1;
    }

    public void set傳票JSON2(string 傳票JSON2)
    {
        this.傳票JSON2 = 傳票JSON2;
    }
}