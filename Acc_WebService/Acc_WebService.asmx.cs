using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Data;
using Newtonsoft.Json;
using System.Web.Configuration;


namespace Acc_WebService
{
    /// <summary>
    ///Acc_WebService 的摘要描述
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允許使用 ASP.NET AJAX 從指令碼呼叫此 Web 服務，請取消註解下列一行。
    // [System.Web.Script.Services.ScriptService]
    public class Acc_WebService : System.Web.Services.WebService
    {
        //輸入條碼,回傳傳票底稿
        [WebMethod]
        public string GetVw_GBCVisaDetail(string fundNo,string acmWordNum)
        {
            //宣告接收從預控端取得之JSON字串
            string JSONReturn = "";

            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {

            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                JSONReturn = ws.GetVw_GBCVisaDetailJSON(acmWordNum); //呼叫預控的服務,取得此動支編號的view資料
            }
            else
            {
                return "基金代號有誤! 號碼為: " + fundNo;
            }

            最外層 vouTop = null; //宣告輸出JSON格式
            最外層 vouTop2 = null; //宣告輸出JSON2格式
            string JSON1 = null; //宣告回傳JSON1

            Vw_GBCVisaDetail vw_GBCVisaDetail = new Vw_GBCVisaDetail();
            List<Vw_GBCVisaDetail> vwList = new List<Vw_GBCVisaDetail>();
            try
            {
                vwList = JsonConvert.DeserializeObject<List<Vw_GBCVisaDetail>>(JSONReturn);  //反序列化JSON               
                
            }
            catch (Exception)
            {
                return JSONReturn;
            }
            //JSON底稿定義
            //List<最外層> vouTopList = new List<最外層>(); //可能有多筆明細號,所以用集合包起來
            List<傳票明細> vouDtlList = new List<傳票明細>();
            List<傳票受款人> vouPayList = new List<傳票受款人>();
            List<傳票內容> vouCollectionList = new List<傳票內容>();

            //如果會開到第二種傳票時須額外再定義一次
            List<傳票明細> vouDtlList2 = new List<傳票明細>();
            List<傳票受款人> vouPayList2 = new List<傳票受款人>();
            List<傳票內容> vouCollectionList2 = new List<傳票內容>();

            var accKindView = from acckind in vwList select acckind.PK_種類;
            var accSumMoney = (from money in vwList select money.F_核定金額).Sum();
            var vouKindView = from voukind in vwList select voukind.基金代碼;
            string accKind = accKindView.ElementAt(0); //取狀態代碼
            string vouKind = vouKindView.ElementAt(0); //取票種類(主要是用來區分付款憑單(vouKind=5)或是支出傳票(vouKind=2))
            GBCVisaDetailAbateDetailDAO dao = new GBCVisaDetailAbateDetailDAO();
            GBCJSONRecordDAO jsonDAO = new GBCJSONRecordDAO();
            int isPrePay = 0; //有無預付
            int isLog = 0; //有無預付

            if (vouKind == "090") //如果是家防基金,使用憑單
            {
                vouKind = "5";
            }
            else
            {
                vouKind = "2";
            }

            /*
             * 一共有六種狀態,分別為:
             * 1.預付    、2.核銷    、3.估列、
             * 4.估列收回、5.預撥收回、6.核銷收回
             */

            //vwList = vwList.OrderBy(x => x.PK_明細號).ToList();

            /*--------------------------------1.預付作業------------------------------------------*/
            if ("預付".Equals(accKind))
            {
                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;

                    try
                    {
                        isLog = dao.FindLog(vw_GBCVisaDetail);
                        string isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        else if ((isLog > 0) && isPass.Equals("0"))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    傳票明細 vouDtl_D = new 傳票明細()
                    {
                        借貸別 = "借",
                        科目代號 = "1154",
                        科目名稱 = "預付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = "",
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_D);
                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        地址 = "",
                        實付金額 = vw_GBCVisaDetail.F_核定金額,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                var vouPayGroup = from xxx in vouPayList
                                  group xxx by new {xxx.統一編號 ,xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號  = g.Key.銀行代號, 銀行名稱  = g.Key.銀行名稱, 銀行帳號  = g.Key.銀行帳號, 帳戶名稱  = g.Key.帳戶名稱, 實付金額  = g.Sum(xxx => xxx.實付金額)};
                vouPayList = new List<傳票受款人>();
                foreach (var vouPayGroupItem in vouPayGroup)
                {
                    傳票受款人 vouPay = new 傳票受款人();
                    vouPay.統一編號 = vouPayGroupItem.統一編號;
                    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                    vouPay.地址 = vouPayGroupItem.地址;
                    vouPay.實付金額 = vouPayGroupItem.實付金額;
                    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                    vouPayList.Add(vouPay);
                }
                傳票主檔 vouMain = new 傳票主檔()
                {
                    傳票種類 = vouKind,
                    製票日期 = vw_GBCVisaDetail.F_製票日,
                    主摘要 = vw_GBCVisaDetail.F_摘要,
                    交付方式 = "1"
                };
                傳票明細 vouDtl_C = new 傳票明細()
                {
                    借貸別 = "貸",
                    科目代號 = "1112",
                    科目名稱 = "銀行存款",
                    摘要 = vw_GBCVisaDetail.F_摘要,
                    金額 = accSumMoney,
                    計畫代碼 = "",
                    用途別代碼 = "",
                    沖轉字號 = "",
                    對象代碼 = "",
                    對象說明 = ""
                };
                vouDtlList.Add(vouDtl_C);
                傳票內容 vouCollection = new 傳票內容()
                {
                    傳票主檔 = vouMain,
                    傳票明細 = vouDtlList,
                    傳票受款人 = vouPayList
                };

                vouCollectionList.Add(vouCollection);
                vouTop = new 最外層()
                {
                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                    年度 = vw_GBCVisaDetail.PK_會計年度,
                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                    種類 = vw_GBCVisaDetail.PK_種類,
                    次別 = vw_GBCVisaDetail.PK_次別,
                    明細號 = vw_GBCVisaDetail.PK_明細號,
                    傳票內容 = vouCollectionList
                };
            }
            /*--------------------------------3.估列作業------------------------------------------*/
            if ("估列".Equals(accKind))
            {
                foreach (var item in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = item.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = item.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = item.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = item.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = item.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = item.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = item.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = item.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = item.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = item.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = item.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = item.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = item.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = item.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = item.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = item.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = item.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = item.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = item.F_批號;
                    try
                    {
                        isLog = dao.FindLog(vw_GBCVisaDetail);
                        string isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        else if ((isLog > 0) && isPass.Equals("0"))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }
                    傳票明細 vouDtl_D = new 傳票明細()
                    {
                        借貸別 = "借",
                        科目代號 = "5",
                        科目名稱 = "基金用途",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = "",
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人
                    };
                    vouDtlList.Add(vouDtl_D);

                    傳票明細 vouDtl_C = new 傳票明細()
                    {
                        借貸別 = "貸",
                        科目代號 = "2125",
                        科目名稱 = "應付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = "",
                        沖轉字號 = "",
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_C);

                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        地址 = "",
                        實付金額 = vw_GBCVisaDetail.F_核定金額,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);

                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                //var vouPayGroup = from xxx in vouPayList
                //                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                //                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                //vouPayList = new List<傳票受款人>();
                //foreach (var vouPayGroupItem in vouPayGroup)
                //{
                //    傳票受款人 vouPay = new 傳票受款人();
                //    vouPay.統一編號 = vouPayGroupItem.統一編號;
                //    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                //    vouPay.地址 = vouPayGroupItem.地址;
                //    vouPay.實付金額 = vouPayGroupItem.實付金額;
                //    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                //    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                //    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                //    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                //    vouPayList.Add(vouPay);
                //}
                傳票主檔 vouMain = new 傳票主檔()
                {
                    傳票種類 = "3",
                    製票日期 = vw_GBCVisaDetail.F_製票日,
                    主摘要 = vw_GBCVisaDetail.F_摘要,
                    交付方式 = "1"
                };

                傳票內容 vouCollection = new 傳票內容()
                {
                    傳票主檔 = vouMain,
                    傳票明細 = vouDtlList,
                    傳票受款人 = vouPayList
                };

                vouCollectionList.Add(vouCollection);
                vouTop = new 最外層()
                {
                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                    年度 = vw_GBCVisaDetail.PK_會計年度,
                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                    種類 = vw_GBCVisaDetail.PK_種類,
                    次別 = vw_GBCVisaDetail.PK_次別,
                    明細號 = vw_GBCVisaDetail.PK_明細號,
                    傳票內容 = vouCollectionList
                };
            }
            /*--------------------------------5.預撥收回作業------------------------------------------*/
            if ("預撥收回".Equals(accKind))
            {
                int prePayMoney = 0;
                int prePayMoneyAbate = 0;
                int prePayBalance = 0;

                //貸方要沖銷預付
                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;
                    try
                    {
                        isLog = dao.FindLog(vw_GBCVisaDetail);
                        string isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        else if ((isLog > 0) && isPass.Equals("0"))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    //計算預付沖銷餘額
                    List<VouDetailVO> prePayNouNoList = dao.FindPrePayVouNo(vw_GBCVisaDetail);
                    foreach (var prePayVouNo in prePayNouNoList)
                    {
                        prePayMoney = prePayMoney + dao.PrePayMoney(vwListItem.基金代碼, vwListItem.PK_會計年度, prePayVouNo.傳票號);
                        prePayMoneyAbate = prePayMoneyAbate + dao.PrePayMoneyAbate(vwListItem.基金代碼, vwListItem.PK_會計年度, prePayVouNo.傳票號, prePayVouNo.傳票明細號);
                    }

                    //預付沖銷餘額 = 已預付 - 已轉正
                    prePayBalance = prePayMoney - prePayMoneyAbate;

                    if (prePayBalance <= 0)
                    {
                        return "預付沖銷餘額不足!";
                    }

                    //找預付沖轉字號
                    var abatePrePayVouNo = from prevou in prePayNouNoList select prevou.傳票號;
                    var abatePrePayVouDtlNo = from prevou in prePayNouNoList select prevou.傳票明細號;

                    int abateCnt = 0;

                    傳票明細 vouDtl_C = new 傳票明細()
                    {
                        借貸別 = "貸",
                        科目代號 = "1154",
                        科目名稱 = "預付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt),
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_C);

                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        地址 = "",
                        實付金額 = vw_GBCVisaDetail.F_核定金額,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    abateCnt++;

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                var vouPayGroup = from xxx in vouPayList
                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                vouPayList = new List<傳票受款人>();
                foreach (var vouPayGroupItem in vouPayGroup)
                {
                    傳票受款人 vouPay = new 傳票受款人();
                    vouPay.統一編號 = vouPayGroupItem.統一編號;
                    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                    vouPay.地址 = vouPayGroupItem.地址;
                    vouPay.實付金額 = vouPayGroupItem.實付金額;
                    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                    vouPayList.Add(vouPay);
                }
                傳票主檔 vouMain = new 傳票主檔()
                {
                    傳票種類 = "1",
                    製票日期 = vw_GBCVisaDetail.F_製票日,
                    主摘要 = vw_GBCVisaDetail.F_摘要,
                    交付方式 = "1"
                };

                傳票明細 vouDtl_D = new 傳票明細()
                {
                    借貸別 = "借",
                    科目代號 = "1112",
                    科目名稱 = "銀行存款",
                    摘要 = vw_GBCVisaDetail.F_摘要,
                    金額 = accSumMoney,
                    計畫代碼 = "",
                    用途別代碼 = "",
                    沖轉字號 = "",
                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                    對象說明 = vw_GBCVisaDetail.F_受款人
                };
                vouDtlList.Add(vouDtl_D);
                傳票內容 vouCollection = new 傳票內容()
                {
                    傳票主檔 = vouMain,
                    傳票明細 = vouDtlList,
                    傳票受款人 = vouPayList
                };

                vouCollectionList.Add(vouCollection);
                vouTop = new 最外層()
                {
                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                    年度 = vw_GBCVisaDetail.PK_會計年度,
                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                    種類 = vw_GBCVisaDetail.PK_種類,
                    次別 = vw_GBCVisaDetail.PK_次別,
                    明細號 = vw_GBCVisaDetail.PK_明細號,
                    傳票內容 = vouCollectionList
                };
            }
            /*--------------------------------4.估列收回作業------------------------------------------*/
            if ("估列收回".Equals(accKind))
            {
                int estimateMoney = 0;
                int estimateMoneyAbate = 0;
                int estimateBalance = 0;

                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = vwListItem.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = vwListItem.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = vwListItem.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = vwListItem.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = vwListItem.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = vwListItem.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = vwListItem.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = vwListItem.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = vwListItem.F_批號;
                    try
                    {
                        isLog = dao.FindLog(vw_GBCVisaDetail);
                        string isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        else if ((isLog > 0) && isPass.Equals("0"))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    //計算估列沖銷餘額
                    List<VouDetailVO> estimateNouNoList = dao.FindEstimateVouNo(vw_GBCVisaDetail);
                    foreach (var estimateVouNo in estimateNouNoList)
                    {
                        //算已估列總額(此受款人編號)
                        estimateMoney = estimateMoney + dao.EstimateMoney(vwListItem.基金代碼, vwListItem.PK_會計年度, estimateVouNo.傳票號);
                        //算已沖估列總額(此受款人編號)
                        estimateMoneyAbate = estimateMoneyAbate + dao.EstimateMoneyAbate(vwListItem.基金代碼, vwListItem.PK_會計年度, estimateVouNo.傳票號, estimateVouNo.傳票明細號);
                    }

                    //估列沖銷餘額 = 已估列 - 已沖
                    estimateBalance = estimateMoney - estimateMoneyAbate;

                    if (estimateBalance <= 0)
                    {
                        return "估列沖銷餘額不足!";
                    }

                    //找應付沖轉字號
                    var abateEstimateVouNo = from estvou in estimateNouNoList select estvou.傳票號;
                    var abateEstimateVouDtlNo = from estvou in estimateNouNoList select estvou.傳票明細號;

                    int abateCnt = 0;

                    傳票明細 vouDtl_D = new 傳票明細()
                    {
                        借貸別 = "借",
                        科目代號 = "2125",
                        科目名稱 = "應付費用",
                        摘要 = vw_GBCVisaDetail.F_摘要,
                        金額 = vw_GBCVisaDetail.F_核定金額,
                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                        沖轉字號 = abateEstimateVouNo.ElementAt(abateCnt) + "-" + abateEstimateVouDtlNo.ElementAt(abateCnt),
                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                        對象說明 = vw_GBCVisaDetail.F_受款人,
                        明細號 = vw_GBCVisaDetail.PK_明細號
                    };
                    vouDtlList.Add(vouDtl_D);

                    //是否為以前年度
                    if (int.Parse(vw_GBCVisaDetail.PK_動支編號.Substring(0, 3)) < int.Parse(vw_GBCVisaDetail.PK_會計年度))
                    {
                        傳票明細 vouDtl_C = new 傳票明細()
                        {
                            借貸別 = "貸",
                            科目代號 = "4YY",
                            科目名稱 = "雜項收入",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = vw_GBCVisaDetail.F_核定金額,
                            計畫代碼 = "",
                            用途別代碼 = "",
                            沖轉字號 = "",
                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                            對象說明 = vw_GBCVisaDetail.F_受款人
                        };
                        vouDtlList.Add(vouDtl_C);
                    }
                    else
                    {
                        傳票明細 vouDtl_C = new 傳票明細()
                        {
                            借貸別 = "貸",
                            科目代號 = "5",
                            科目名稱 = "基金用途",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = vw_GBCVisaDetail.F_核定金額,
                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                            沖轉字號 = "",
                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                            對象說明 = vw_GBCVisaDetail.F_受款人
                        };
                        vouDtlList.Add(vouDtl_C);
                    }

                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        地址 = "",
                        實付金額 = vw_GBCVisaDetail.F_核定金額,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    abateCnt++;

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                //var vouPayGroup = from xxx in vouPayList
                //                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                //                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                //vouPayList = new List<傳票受款人>();
                //foreach (var vouPayGroupItem in vouPayGroup)
                //{
                //    傳票受款人 vouPay = new 傳票受款人();
                //    vouPay.統一編號 = vouPayGroupItem.統一編號;
                //    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                //    vouPay.地址 = vouPayGroupItem.地址;
                //    vouPay.實付金額 = vouPayGroupItem.實付金額;
                //    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                //    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                //    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                //    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                //    vouPayList.Add(vouPay);
                //}
                傳票主檔 vouMain = new 傳票主檔()
                {
                    傳票種類 = "3",
                    製票日期 = vw_GBCVisaDetail.F_製票日,
                    主摘要 = vw_GBCVisaDetail.F_摘要,
                    交付方式 = "1"
                };
                

                傳票內容 vouCollection = new 傳票內容()
                {
                    傳票主檔 = vouMain,
                    傳票明細 = vouDtlList,
                    傳票受款人 = vouPayList
                };

                vouCollectionList.Add(vouCollection);
                vouTop = new 最外層()
                {
                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                    年度 = vw_GBCVisaDetail.PK_會計年度,
                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                    種類 = vw_GBCVisaDetail.PK_種類,
                    次別 = vw_GBCVisaDetail.PK_次別,
                    明細號 = vw_GBCVisaDetail.PK_明細號,
                    傳票內容 = vouCollectionList
                };
            }
            /*--------------------------------6.核銷收回作業------------------------------------------*/
            if ("核銷收回".Equals(accKind))
            {
                foreach (var item in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = item.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = item.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = item.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = item.PK_種類;
                    vw_GBCVisaDetail.PK_次別 = item.PK_次別;
                    vw_GBCVisaDetail.PK_明細號 = item.PK_明細號;
                    vw_GBCVisaDetail.F_科室代碼 = item.F_科室代碼;
                    vw_GBCVisaDetail.F_用途別代碼 = item.F_用途別代碼;
                    vw_GBCVisaDetail.F_計畫代碼 = item.F_計畫代碼;
                    vw_GBCVisaDetail.F_動支金額 = item.F_動支金額;
                    vw_GBCVisaDetail.F_製票日 = item.F_製票日;
                    vw_GBCVisaDetail.F_是否核定 = item.F_是否核定;
                    vw_GBCVisaDetail.F_核定金額 = item.F_核定金額;
                    vw_GBCVisaDetail.F_核定日期 = item.F_核定日期;
                    vw_GBCVisaDetail.F_摘要 = item.F_摘要;
                    vw_GBCVisaDetail.F_受款人 = item.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = item.F_受款人編號;
                    vw_GBCVisaDetail.F_原動支編號 = item.F_原動支編號;
                    vw_GBCVisaDetail.F_批號 = item.F_批號;
                    try
                    {
                        isLog = dao.FindLog(vw_GBCVisaDetail);
                        string isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                        if ((isLog > 0) && isPass.Equals("1"))
                        {
                            return "此筆資料已轉入過,並且結案。";
                        }
                        else if ((isLog > 0) && isPass.Equals("0"))
                        {
                            dao.Update(vw_GBCVisaDetail);
                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                        }
                        else
                        {
                            dao.Insert(vw_GBCVisaDetail);
                        }
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    //計算估列沖銷餘額
                    List<VouDetailVO> payVouNoList = dao.FindPayVouNo(vw_GBCVisaDetail);

                    //找應付沖轉字號
                    var payVouNo = from payvou in payVouNoList select payvou.傳票號;
                    var payVouDtlNo = from payvou in payVouNoList select payvou.傳票明細號;

                    int abateCnt = 0;

                    //是否為以前年度
                    if (int.Parse(vw_GBCVisaDetail.PK_動支編號.Substring(0, 3)) < int.Parse(vw_GBCVisaDetail.PK_會計年度))
                    {
                        傳票明細 vouDtl_C = new 傳票明細()
                        {
                            借貸別 = "貸",
                            科目代號 = "4YY",
                            科目名稱 = "雜項收入",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = vw_GBCVisaDetail.F_核定金額,
                            計畫代碼 = "",
                            用途別代碼 = "",
                            沖轉字號 = "",//不用沖
                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                            對象說明 = vw_GBCVisaDetail.F_受款人
                        };
                        vouDtlList.Add(vouDtl_C);
                    }
                    else
                    {
                        傳票明細 vouDtl_C = new 傳票明細()
                        {
                            借貸別 = "貸",
                            科目代號 = "5",
                            科目名稱 = "基金用途",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = vw_GBCVisaDetail.F_核定金額,
                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                            沖轉字號 = payVouNo.ElementAt(abateCnt) + "-" + payVouDtlNo.ElementAt(abateCnt),
                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                            對象說明 = vw_GBCVisaDetail.F_受款人,
                            明細號 = vw_GBCVisaDetail.PK_明細號
                        };
                        vouDtlList.Add(vouDtl_C);
                    }
                    傳票受款人 vouPay = new 傳票受款人()
                    {
                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                        地址 = "",
                        實付金額 = vw_GBCVisaDetail.F_核定金額,
                        銀行代號 = "",
                        銀行名稱 = "",
                        銀行帳號 = "",
                        帳戶名稱 = ""
                    };
                    vouPayList.Add(vouPay);

                    abateCnt++;

                    //填傳票明細號1
                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                }
                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                var vouPayGroup = from xxx in vouPayList
                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                vouPayList = new List<傳票受款人>();
                foreach (var vouPayGroupItem in vouPayGroup)
                {
                    傳票受款人 vouPay = new 傳票受款人();
                    vouPay.統一編號 = vouPayGroupItem.統一編號;
                    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                    vouPay.地址 = vouPayGroupItem.地址;
                    vouPay.實付金額 = vouPayGroupItem.實付金額;
                    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                    vouPayList.Add(vouPay);
                }
                傳票主檔 vouMain = new 傳票主檔()
                {
                    傳票種類 = vouKind,
                    製票日期 = vw_GBCVisaDetail.F_製票日,
                    主摘要 = vw_GBCVisaDetail.F_摘要,
                    交付方式 = "1"
                };

                傳票明細 vouDtl_D = new 傳票明細()
                {
                    借貸別 = "借",
                    科目代號 = "1112",
                    科目名稱 = "銀行存款",
                    摘要 = vw_GBCVisaDetail.F_摘要,
                    金額 = accSumMoney,
                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                    沖轉字號 = "",
                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                    對象說明 = vw_GBCVisaDetail.F_受款人
                };
                vouDtlList.Add(vouDtl_D);

                傳票內容 vouCollection = new 傳票內容()
                {
                    傳票主檔 = vouMain,
                    傳票明細 = vouDtlList,
                    傳票受款人 = vouPayList
                };

                vouCollectionList.Add(vouCollection);
                vouTop = new 最外層()
                {
                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                    年度 = vw_GBCVisaDetail.PK_會計年度,
                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                    種類 = vw_GBCVisaDetail.PK_種類,
                    次別 = vw_GBCVisaDetail.PK_次別,
                    明細號 = vw_GBCVisaDetail.PK_明細號,
                    傳票內容 = vouCollectionList
                };
            }
            /*--------------------------------2.核銷作業------------------------------------------*/
            if ("核銷".Equals(accKind))
            {
                foreach (var vwListItem in vwList)
                {
                    vw_GBCVisaDetail.基金代碼 = vwListItem.基金代碼;
                    vw_GBCVisaDetail.PK_會計年度 = vwListItem.PK_會計年度;
                    vw_GBCVisaDetail.PK_動支編號 = vwListItem.PK_動支編號;
                    vw_GBCVisaDetail.PK_種類 = vwListItem.PK_種類;
                    vw_GBCVisaDetail.PK_明細號 = vwListItem.PK_明細號;
                    vw_GBCVisaDetail.PK_次別 = vwListItem.PK_次別;
                    vw_GBCVisaDetail.F_核定金額 = vwListItem.F_核定金額;
                    vw_GBCVisaDetail.F_受款人 = vwListItem.F_受款人;
                    vw_GBCVisaDetail.F_受款人編號 = vwListItem.F_受款人編號;
                    vw_GBCVisaDetail.F_計畫代碼 = vwListItem.F_計畫代碼;
                    
                    //傳票號1有值,而且尚未結案時,直接回傳JSON2
                    string isVouNo1 = dao.FindVouNo(vwListItem.基金代碼, vwListItem.PK_會計年度, vwListItem.PK_動支編號, vwListItem.PK_種類, vwListItem.PK_次別, vwListItem.PK_明細號);
                    string isPass = jsonDAO.IsPass(vwListItem.基金代碼, vwListItem.PK_會計年度, vwListItem.PK_動支編號, vwListItem.PK_種類, vwListItem.PK_次別);

                    if (((isVouNo1.Trim()).Length != 0) && isPass == "0")
                    {
                        return jsonDAO.FindJSON2(vwListItem.基金代碼, vwListItem.PK_會計年度, vwListItem.PK_動支編號, vwListItem.PK_種類, vwListItem.PK_次別);
                    }

                    



                    //是否為暫付及待結轉帳項
                    if (vw_GBCVisaDetail.F_計畫代碼 == "1315")
                    {
                        foreach (var payCash in vwList)
                        {
                            vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                            vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                            vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                            vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                            vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                            vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                            vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                            vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                            vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                            vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                            vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                            vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                            vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                            vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                            vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                            vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                            vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                            vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                            try
                            {
                                isLog = dao.FindLog(vw_GBCVisaDetail);
                                isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                if ((isLog > 0) && isPass.Equals("1"))
                                {
                                    return "此筆資料已轉入過,並且結案。";
                                }
                                else if ((isLog > 0) && isPass.Equals("0"))
                                {
                                    dao.Update(vw_GBCVisaDetail);
                                    jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                }
                                else
                                {
                                    dao.Insert(vw_GBCVisaDetail);
                                }
                            }
                            catch (Exception e)
                            {
                                return e.Message;
                            }

                            傳票明細 vouDtl_D = new 傳票明細()
                            {
                                借貸別 = "借",
                                科目代號 = "1315",
                                科目名稱 = "暫付及待結轉帳項",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = vw_GBCVisaDetail.F_核定金額,
                                計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                沖轉字號 = "",
                                對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                對象說明 = vw_GBCVisaDetail.F_受款人,
                                明細號 = vw_GBCVisaDetail.PK_明細號
                            };
                            vouDtlList.Add(vouDtl_D);
                            傳票受款人 vouPay = new 傳票受款人()
                            {
                                統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                地址 = "",
                                實付金額 = vw_GBCVisaDetail.F_核定金額,
                                銀行代號 = "",
                                銀行名稱 = "",
                                銀行帳號 = "",
                                帳戶名稱 = ""
                            };
                            vouPayList.Add(vouPay);

                            //填傳票明細號1
                            //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                        }
                        //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                        var vouPayGroup = from xxx in vouPayList
                                          group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                          select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                        vouPayList = new List<傳票受款人>();
                        foreach (var vouPayGroupItem in vouPayGroup)
                        {
                            傳票受款人 vouPay = new 傳票受款人();
                            vouPay.統一編號 = vouPayGroupItem.統一編號;
                            vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                            vouPay.地址 = vouPayGroupItem.地址;
                            vouPay.實付金額 = vouPayGroupItem.實付金額;
                            vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                            vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                            vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                            vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                            vouPayList.Add(vouPay);
                        }
                        傳票主檔 vouMain = new 傳票主檔()
                        {
                            傳票種類 = vouKind,
                            製票日期 = vw_GBCVisaDetail.F_製票日,
                            主摘要 = vw_GBCVisaDetail.F_摘要,
                            交付方式 = "1"
                        };
                        傳票明細 vouDtl_C = new 傳票明細()
                        {
                            借貸別 = "貸",
                            科目代號 = "1112",
                            科目名稱 = "銀行存款",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = accSumMoney,
                            計畫代碼 = "",
                            用途別代碼 = "",
                            沖轉字號 = "",
                            對象代碼 = "",
                            對象說明 = ""
                        };
                        vouDtlList.Add(vouDtl_C);
                        傳票內容 vouCollection = new 傳票內容()
                        {
                            傳票主檔 = vouMain,
                            傳票明細 = vouDtlList,
                            傳票受款人 = vouPayList
                        };

                        vouCollectionList.Add(vouCollection);
                        vouTop = new 最外層()
                        {
                            基金代碼 = vw_GBCVisaDetail.基金代碼,
                            年度 = vw_GBCVisaDetail.PK_會計年度,
                            動支編號 = vw_GBCVisaDetail.PK_動支編號,
                            種類 = vw_GBCVisaDetail.PK_種類,
                            次別 = vw_GBCVisaDetail.PK_次別,
                            明細號 = vw_GBCVisaDetail.PK_明細號,
                            傳票內容 = vouCollectionList
                        };

                        //紀錄第一張傳票底稿
                        try
                        {
                            jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                        }
                        catch (Exception e)
                        {
                            return e.Message;
                        }

                        JSON1 = jsonDAO.FindJSON1(vw_GBCVisaDetail);

                        return JSON1;
                    }

                    //判斷是否有預付(基金代號-動支編號-預付-受編)
                    try
                    {
                        isPrePay = dao.FindPrePay(vw_GBCVisaDetail);
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }

                    //有預付
                    if (isPrePay > 0)
                    {
                        //若有預付紀錄:
                        //算出"預付未沖銷數"及"估列未沖銷數"

                        int prePayMoney = 0;
                        int prePayMoneyAbate = 0;
                        int prePayBalance = 0;
                        int estimateMoney = 0;
                        int estimateMoneyAbate = 0;
                        int estimateBalance = 0;

                        //開始計算預付沖銷餘額
                        //找預付傳票號
                        List<VouDetailVO> prePayNouNoList = dao.FindPrePayVouNo(vw_GBCVisaDetail);
                        foreach (var prePayVouNo in prePayNouNoList)
                        {
                            prePayMoney = prePayMoney + dao.PrePayMoney(vwListItem.基金代碼, prePayVouNo.傳票年度, prePayVouNo.傳票號);
                            prePayMoneyAbate = prePayMoneyAbate + dao.PrePayMoneyAbate(vwListItem.基金代碼, prePayVouNo.傳票年度, prePayVouNo.傳票號, prePayVouNo.傳票明細號);
                        }

                        //預付沖銷餘額 = 已預付 - 已轉正
                        prePayBalance = prePayMoney - prePayMoneyAbate;

                        //計算估列沖銷餘額
                        List<VouDetailVO> estimateNouNoList = dao.FindEstimateVouNo(vw_GBCVisaDetail);
                        foreach (var estimateVouNo in estimateNouNoList)
                        {
                            estimateMoney = estimateMoney + dao.EstimateMoney(vwListItem.基金代碼, estimateVouNo.傳票年度, estimateVouNo.傳票號);
                            estimateMoneyAbate = estimateMoneyAbate + dao.EstimateMoneyAbate(vwListItem.基金代碼, estimateVouNo.傳票年度, estimateVouNo.傳票號, estimateVouNo.傳票明細號);
                        }

                        //估列沖銷餘額 = 已估列 - 已轉正
                        estimateBalance = estimateMoney - estimateMoneyAbate;

                        //找預付沖轉字號
                        var abatePrePayVouNo = from prevou in prePayNouNoList select prevou.傳票號;
                        var abatePrePayVouDtlNo = from prevou in prePayNouNoList select prevou.傳票明細號;

                        //找應付沖轉字號
                        var abateEstimateVouNo = from estvou in estimateNouNoList select estvou.傳票號;
                        var abateEstimateVouDtlNo = from estvou in estimateNouNoList select estvou.傳票明細號;

                        int abateCnt = 0;

                        //判斷核銷傳票
                        //狀況1: 預付沖銷餘額<=0 ----實支
                        //狀況2: 預付沖銷餘額>0 && 估列沖銷餘額<=0 -----無估列之轉正
                        //  2-1: 核銷數<預付沖銷餘額----分轉(貸預付 借費用)
                        //  2-2: 核銷數>預付沖銷餘額----分轉(貸預借 借費用) + 支(借費用 貸銀行)
                        //狀況3: 預付沖銷餘額>0 && 估列沖銷餘額>0  -----有估列的轉正
                        //  3-1: 預付沖銷餘額>Y ---分轉(貸預付 借應付 借費用)
                        //       核銷數>預付沖銷餘額 ---增加 支(借費用 貸銀行)
                        //  3-2: 預付沖銷餘額<=Y---分轉(貸預付 借應付)
                        //       核銷數>預付沖銷餘額 ---增加 支(借費用 貸銀行)

                        if (prePayBalance <= 0) //實支
                        {
                            foreach (var payCash in vwList)
                            {
                                vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                try
                                {
                                    isLog = dao.FindLog(vw_GBCVisaDetail);
                                    isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                    if ((isLog > 0) && isPass.Equals("1"))
                                    {
                                        return "此筆資料已轉入過,並且結案。";
                                    }
                                    else if ((isLog > 0) && isPass.Equals("0"))
                                    {
                                        dao.Update(vw_GBCVisaDetail);
                                        jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                    }
                                    else
                                    {
                                        dao.Insert(vw_GBCVisaDetail);
                                    }
                                }
                                catch (Exception e)
                                {
                                    return e.Message;
                                }

                                傳票明細 vouDtl_D = new 傳票明細()
                                {
                                    借貸別 = "借",
                                    科目代號 = "5",
                                    科目名稱 = "基金用途",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.F_核定金額,
                                    計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                    用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                    沖轉字號 = "",
                                    對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                    對象說明 = vw_GBCVisaDetail.F_受款人,
                                    明細號 = vw_GBCVisaDetail.PK_明細號
                                };
                                //是否為以前年度
                                if (int.Parse(vw_GBCVisaDetail.PK_動支編號.Substring(0, 3)) < int.Parse(vw_GBCVisaDetail.PK_會計年度))
                                {
                                    vouDtl_D.用途別代碼 = "91Y";
                                }
                                vouDtlList.Add(vouDtl_D);

                                傳票受款人 vouPay = new 傳票受款人()
                                {
                                    統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                    受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                    地址 = "",
                                    實付金額 = vw_GBCVisaDetail.F_核定金額,
                                    銀行代號 = "",
                                    銀行名稱 = "",
                                    銀行帳號 = "",
                                    帳戶名稱 = ""
                                };
                                vouPayList.Add(vouPay);

                                abateCnt++;

                                //填傳票明細號1
                                //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                            }
                            //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                            var vouPayGroup = from xxx in vouPayList
                                              group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                              select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                            vouPayList = new List<傳票受款人>();
                            foreach (var vouPayGroupItem in vouPayGroup)
                            {
                                傳票受款人 vouPay = new 傳票受款人();
                                vouPay.統一編號 = vouPayGroupItem.統一編號;
                                vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                vouPay.地址 = vouPayGroupItem.地址;
                                vouPay.實付金額 = vouPayGroupItem.實付金額;
                                vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                vouPayList.Add(vouPay);
                            }
                            傳票主檔 vouMain = new 傳票主檔()
                            {
                                傳票種類 = vouKind,
                                製票日期 = vw_GBCVisaDetail.F_製票日,
                                主摘要 = vw_GBCVisaDetail.F_摘要,
                                交付方式 = "1"
                            };
                            傳票明細 vouDtl_C = new 傳票明細()
                            {
                                借貸別 = "貸",
                                科目代號 = "1112",
                                科目名稱 = "銀行存款",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = accSumMoney,
                                計畫代碼 = "",
                                用途別代碼 = "",
                                沖轉字號 = "",
                                對象代碼 = "",
                                對象說明 = ""
                            };
                            vouDtlList.Add(vouDtl_C);
                            傳票內容 vouCollection = new 傳票內容()
                            {
                                傳票主檔 = vouMain,
                                傳票明細 = vouDtlList,
                                傳票受款人 = vouPayList
                            };

                            vouCollectionList.Add(vouCollection);
                            vouTop = new 最外層()
                            {
                                基金代碼 = vw_GBCVisaDetail.基金代碼,
                                年度 = vw_GBCVisaDetail.PK_會計年度,
                                動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                種類 = vw_GBCVisaDetail.PK_種類,
                                次別 = vw_GBCVisaDetail.PK_次別,
                                明細號 = vw_GBCVisaDetail.PK_明細號,
                                傳票內容 = vouCollectionList
                            };

                            //紀錄第一張傳票底稿
                            try
                            {
                                jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                            }
                            catch (Exception e)
                            {
                                return e.Message;
                            }

                            JSON1 = jsonDAO.FindJSON1(vw_GBCVisaDetail);

                            return JSON1;
                        }
                        else if (prePayBalance > 0 && estimateBalance <= 0) //無估列之轉正
                        {
                            if (vw_GBCVisaDetail.F_核定金額 <= prePayBalance) //分轉(貸預付 借費用) only
                            {                                
                                foreach (var payCash in vwList)
                                {
                                    vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                    vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                    vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                    vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                    vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                    vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                    vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                    vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                    vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                    vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                    vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                    vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                    vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                    vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                    vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                    vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                    vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                    vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                    try
                                    {
                                        isLog = dao.FindLog(vw_GBCVisaDetail);
                                        isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                        if ((isLog > 0) && isPass.Equals("1"))
                                        {
                                            return "此筆資料已轉入過,並且結案。";
                                        }
                                        else if ((isLog > 0) && isPass.Equals("0"))
                                        {
                                            dao.Update(vw_GBCVisaDetail);
                                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                        }
                                        else
                                        {
                                            dao.Insert(vw_GBCVisaDetail);
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    傳票明細 vouDtl_C = new 傳票明細()
                                    {
                                        借貸別 = "貸",
                                        科目代號 = "1154",
                                        科目名稱 = "預付費用",
                                        摘要 = vw_GBCVisaDetail.F_摘要,
                                        金額 = vw_GBCVisaDetail.F_核定金額,
                                        計畫代碼 = "",
                                        用途別代碼 = "",
                                        沖轉字號 = abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt), //沖轉支出傳票 from prePayNouNoList
                                        對象代碼 = "",
                                        對象說明 = ""
                                    };
                                    vouDtlList.Add(vouDtl_C);
                                    傳票明細 vouDtl_D = new 傳票明細()
                                    {
                                        借貸別 = "借",
                                        科目代號 = "5",
                                        科目名稱 = "基金用途",
                                        摘要 = vw_GBCVisaDetail.F_摘要,
                                        金額 = vw_GBCVisaDetail.F_核定金額,
                                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                        沖轉字號 = "",
                                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                        對象說明 = vw_GBCVisaDetail.F_受款人,
                                        明細號 = vw_GBCVisaDetail.PK_明細號
                                    };
                                    vouDtlList.Add(vouDtl_D);
                                    傳票受款人 vouPay = new 傳票受款人()
                                    {
                                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                        地址 = "",
                                        實付金額 = vw_GBCVisaDetail.F_核定金額,
                                        銀行代號 = "",
                                        銀行名稱 = "",
                                        銀行帳號 = "",
                                        帳戶名稱 = ""
                                    };
                                    vouPayList.Add(vouPay);

                                    abateCnt++;

                                    //填傳票明細號1
                                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                }
                                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                var vouPayGroup = from xxx in vouPayList
                                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                vouPayList = new List<傳票受款人>();
                                foreach (var vouPayGroupItem in vouPayGroup)
                                {
                                    傳票受款人 vouPay = new 傳票受款人();
                                    vouPay.統一編號 = vouPayGroupItem.統一編號;
                                    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                    vouPay.地址 = vouPayGroupItem.地址;
                                    vouPay.實付金額 = vouPayGroupItem.實付金額;
                                    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                    vouPayList.Add(vouPay);
                                }
                                傳票主檔 vouMain = new 傳票主檔()
                                {
                                    傳票種類 = "4",
                                    製票日期 = vw_GBCVisaDetail.F_製票日,
                                    主摘要 = vw_GBCVisaDetail.F_摘要,
                                    交付方式 = "1"
                                };

                                傳票內容 vouCollection = new 傳票內容()
                                {
                                    傳票主檔 = vouMain,
                                    傳票明細 = vouDtlList,
                                    傳票受款人 = vouPayList
                                };

                                vouCollectionList.Add(vouCollection);
                                vouTop = new 最外層()
                                {
                                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                                    年度 = vw_GBCVisaDetail.PK_會計年度,
                                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                    種類 = vw_GBCVisaDetail.PK_種類,
                                    次別 = vw_GBCVisaDetail.PK_次別,
                                    明細號 = vw_GBCVisaDetail.PK_明細號,
                                    傳票內容 = vouCollectionList
                                };

                                //紀錄第一張傳票底稿
                                try
                                {
                                    jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                                }
                                catch (Exception e)
                                {
                                    return e.Message;
                                }

                                //回傳第一張傳票底稿
                                JSON1 = jsonDAO.FindJSON1(vw_GBCVisaDetail);

                                return JSON1;
                            }
                            else //分轉(貸預付 借費用) + 支(借費用 貸銀行)
                            {
                                //------分轉-------
                                foreach (var payCash in vwList)
                                {
                                    vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                    vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                    vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                    vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                    vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                    vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                    vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                    vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                    vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                    vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                    vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                    vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                    vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                    vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                    vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                    vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                    vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                    vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                    try
                                    {
                                        isLog = dao.FindLog(vw_GBCVisaDetail);
                                        isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                        if ((isLog > 0) && isPass.Equals("1"))
                                        {
                                            return "此筆資料已轉入過,並且結案。";
                                        }
                                        else if ((isLog > 0) && isPass.Equals("0"))
                                        {
                                            dao.Update(vw_GBCVisaDetail);
                                            jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                        }
                                        else
                                        {
                                            dao.Insert(vw_GBCVisaDetail);
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    傳票明細 vouDtl_C = new 傳票明細()
                                    {
                                        借貸別 = "貸",
                                        科目代號 = "1154",
                                        科目名稱 = "預付費用",
                                        摘要 = vw_GBCVisaDetail.F_摘要,
                                        金額 = prePayBalance,
                                        計畫代碼 = "",
                                        用途別代碼 = "",
                                        沖轉字號 = abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt), //沖轉支出傳票 from prePayNouNoList
                                        對象代碼 = "",
                                        對象說明 = ""
                                    };
                                    vouDtlList.Add(vouDtl_C);
                                    傳票明細 vouDtl_D = new 傳票明細()
                                    {
                                        借貸別 = "借",
                                        科目代號 = "5",
                                        科目名稱 = "基金用途",
                                        摘要 = vw_GBCVisaDetail.F_摘要,
                                        金額 = prePayBalance,
                                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                        沖轉字號 = "",
                                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                        對象說明 = vw_GBCVisaDetail.F_受款人,
                                        明細號 = vw_GBCVisaDetail.PK_明細號
                                    };
                                    vouDtlList.Add(vouDtl_D);
                                    傳票受款人 vouPay = new 傳票受款人()
                                    {
                                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                        地址 = "",
                                        實付金額 = prePayBalance,
                                        銀行代號 = "",
                                        銀行名稱 = "",
                                        銀行帳號 = "",
                                        帳戶名稱 = ""
                                    };
                                    vouPayList.Add(vouPay);

                                    abateCnt++;

                                    //填傳票明細號1
                                    //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                }
                                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                var vouPayGroup = from xxx in vouPayList
                                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                vouPayList = new List<傳票受款人>();
                                foreach (var vouPayGroupItem in vouPayGroup)
                                {
                                    傳票受款人 vouPay = new 傳票受款人();
                                    vouPay.統一編號 = vouPayGroupItem.統一編號;
                                    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                    vouPay.地址 = vouPayGroupItem.地址;
                                    vouPay.實付金額 = vouPayGroupItem.實付金額;
                                    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                    vouPayList.Add(vouPay);
                                }
                                傳票主檔 vouMain = new 傳票主檔()
                                {
                                    傳票種類 = "4",
                                    製票日期 = vw_GBCVisaDetail.F_製票日,
                                    主摘要 = vw_GBCVisaDetail.F_摘要,
                                    交付方式 = "1"
                                };

                                傳票內容 vouCollection = new 傳票內容()
                                {
                                    傳票主檔 = vouMain,
                                    傳票明細 = vouDtlList,
                                    傳票受款人 = vouPayList
                                };

                                vouCollectionList.Add(vouCollection);
                                vouTop = new 最外層()
                                {
                                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                                    年度 = vw_GBCVisaDetail.PK_會計年度,
                                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                    種類 = vw_GBCVisaDetail.PK_種類,
                                    次別 = vw_GBCVisaDetail.PK_次別,
                                    明細號 = vw_GBCVisaDetail.PK_明細號,
                                    傳票內容 = vouCollectionList
                                };

                                //------支出傳票------
                                foreach (var payCash in vwList)
                                {
                                    vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                    vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                    vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                    vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                    vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                    vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                    vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                    vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                    vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                    vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                    vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                    vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                    vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                    vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                    vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                    vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                    vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                    vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                    傳票明細 vouDtl_D2 = new 傳票明細()
                                    {
                                        借貸別 = "借",
                                        科目代號 = "5",
                                        科目名稱 = "基金用途",
                                        摘要 = vw_GBCVisaDetail.F_摘要,
                                        金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                        計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                        用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                        沖轉字號 = "",
                                        對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                        對象說明 = vw_GBCVisaDetail.F_受款人,
                                        明細號 = vw_GBCVisaDetail.PK_明細號
                                    };
                                    vouDtlList2.Add(vouDtl_D2);
                                    傳票受款人 vouPay2 = new 傳票受款人()
                                    {
                                        統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                        受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                        地址 = "",
                                        實付金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                        銀行代號 = "",
                                        銀行名稱 = "",
                                        銀行帳號 = "",
                                        帳戶名稱 = ""
                                    };
                                    vouPayList2.Add(vouPay2);

                                    //填傳票明細號2
                                    dao.FillVouDtl2(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                }
                                //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                var vouPayGroup2 = from xxx in vouPayList2
                                                  group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                  select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                vouPayList2 = new List<傳票受款人>();
                                foreach (var vouPayGroupItem in vouPayGroup2)
                                {
                                    傳票受款人 vouPay = new 傳票受款人();
                                    vouPay.統一編號 = vouPayGroupItem.統一編號;
                                    vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                    vouPay.地址 = vouPayGroupItem.地址;
                                    vouPay.實付金額 = vouPayGroupItem.實付金額;
                                    vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                    vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                    vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                    vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                    vouPayList2.Add(vouPay);
                                }
                                傳票主檔 vouMain2 = new 傳票主檔()
                                {
                                    傳票種類 = vouKind,
                                    製票日期 = vw_GBCVisaDetail.F_製票日,
                                    主摘要 = vw_GBCVisaDetail.F_摘要,
                                    交付方式 = "1"
                                };
                                傳票明細 vouDtl_C2 = new 傳票明細()
                                {
                                    借貸別 = "貸",
                                    科目代號 = "1112",
                                    科目名稱 = "銀行存款",
                                    摘要 = vw_GBCVisaDetail.F_摘要,
                                    金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                    計畫代碼 = "",
                                    用途別代碼 = "",
                                    沖轉字號 = "",
                                    對象代碼 = "",
                                    對象說明 = ""
                                };
                                vouDtlList2.Add(vouDtl_C2);
                                傳票內容 vouCollection2 = new 傳票內容()
                                {
                                    傳票主檔 = vouMain2,
                                    傳票明細 = vouDtlList2,
                                    傳票受款人 = vouPayList2
                                };

                                vouCollectionList2.Add(vouCollection2);

                                vouTop2 = new 最外層()
                                {
                                    基金代碼 = vw_GBCVisaDetail.基金代碼,
                                    年度 = vw_GBCVisaDetail.PK_會計年度,
                                    動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                    種類 = vw_GBCVisaDetail.PK_種類,
                                    次別 = vw_GBCVisaDetail.PK_次別,
                                    明細號 = vw_GBCVisaDetail.PK_明細號,
                                    傳票內容 = vouCollectionList2
                                };

                                //紀錄第一張傳票底稿
                                try
                                {
                                    jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                                }
                                catch (Exception e)
                                {
                                    return e.Message;
                                }

                                //紀錄第二張傳票底稿
                                try
                                {
                                    jsonDAO.UpdateJsonRecord2(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop2));
                                }
                                catch (Exception e)
                                {
                                    return e.Message;
                                }

                                return JsonConvert.SerializeObject(vouTop);
                            }
                        }
                        else if (prePayBalance > 0 && estimateBalance > 0) //有估列的轉正
                        {
                            if (prePayBalance > estimateBalance) //分轉(貸預付 借應付 借費用)
                            {
                                if (vw_GBCVisaDetail.F_核定金額 < prePayBalance) //不用加開支出傳票
                                {
                                    //----------分轉(貸預付 借應付 借費用)-----------
                                    foreach (var payCash in vwList)
                                    {
                                        vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                        vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                        vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                        vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                        vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                        vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                        vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                        vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                        vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                        vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                        vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                        vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                        vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                        vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                        vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                        vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                        vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                        vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                        try
                                        {
                                            isLog = dao.FindLog(vw_GBCVisaDetail);
                                            isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                            if ((isLog > 0) && isPass.Equals("1"))
                                            {
                                                return "此筆資料已轉入過,並且結案。";
                                            }
                                            else if ((isLog > 0) && isPass.Equals("0"))
                                            {
                                                dao.Update(vw_GBCVisaDetail);
                                                jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                            }
                                            else
                                            {
                                                dao.Insert(vw_GBCVisaDetail);
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            return e.Message;
                                        }

                                        傳票明細 vouDtl_C = new 傳票明細()
                                        {
                                            借貸別 = "貸",
                                            科目代號 = "1154",
                                            科目名稱 = "預付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = vw_GBCVisaDetail.F_核定金額,
                                            計畫代碼 = "",
                                            用途別代碼 = "",
                                            沖轉字號 = abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt), //沖轉支出傳票 from prePayNouNoList
                                            對象代碼 = "",
                                            對象說明 = ""
                                        };
                                        vouDtlList.Add(vouDtl_C);
                                        傳票明細 vouDtl_D = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "2125",
                                            科目名稱 = "應付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = estimateBalance,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = abateEstimateVouNo.ElementAt(abateCnt) + "-" + abateEstimateVouDtlNo.ElementAt(abateCnt),
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList.Add(vouDtl_D);
                                        傳票明細 vouDtl_D2 = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "5",
                                            科目名稱 = "基金用途",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = vw_GBCVisaDetail.F_核定金額 - estimateBalance,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = "",
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList.Add(vouDtl_D2);
                                        傳票受款人 vouPay = new 傳票受款人()
                                        {
                                            統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                            受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                            地址 = "",
                                            實付金額 = vw_GBCVisaDetail.F_核定金額,
                                            銀行代號 = "",
                                            銀行名稱 = "",
                                            銀行帳號 = "",
                                            帳戶名稱 = ""
                                        };
                                        vouPayList.Add(vouPay);

                                        abateCnt++;

                                        //填傳票明細號1
                                        //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                    }
                                    //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                    var vouPayGroup = from xxx in vouPayList
                                                      group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                      select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                    vouPayList = new List<傳票受款人>();
                                    foreach (var vouPayGroupItem in vouPayGroup)
                                    {
                                        傳票受款人 vouPay = new 傳票受款人();
                                        vouPay.統一編號 = vouPayGroupItem.統一編號;
                                        vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                        vouPay.地址 = vouPayGroupItem.地址;
                                        vouPay.實付金額 = vouPayGroupItem.實付金額;
                                        vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                        vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                        vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                        vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                        vouPayList.Add(vouPay);
                                    }
                                    傳票主檔 vouMain = new 傳票主檔()
                                    {
                                        傳票種類 = "4",
                                        製票日期 = vw_GBCVisaDetail.F_製票日,
                                        主摘要 = vw_GBCVisaDetail.F_摘要,
                                        交付方式 = "1"
                                    };

                                    傳票內容 vouCollection = new 傳票內容()
                                    {
                                        傳票主檔 = vouMain,
                                        傳票明細 = vouDtlList,
                                        傳票受款人 = vouPayList
                                    };

                                    vouCollectionList.Add(vouCollection);
                                    vouTop = new 最外層()
                                    {
                                        基金代碼 = vw_GBCVisaDetail.基金代碼,
                                        年度 = vw_GBCVisaDetail.PK_會計年度,
                                        動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                        種類 = vw_GBCVisaDetail.PK_種類,
                                        次別 = vw_GBCVisaDetail.PK_次別,
                                        明細號 = vw_GBCVisaDetail.PK_明細號,
                                        傳票內容 = vouCollectionList
                                    };

                                    //紀錄第一張傳票底稿
                                    try
                                    {
                                        jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    //回傳第一張傳票底稿
                                    JSON1 = jsonDAO.FindJSON1(vw_GBCVisaDetail);

                                    return JSON1;
                                }
                                else //分轉(貸預付 借應付 借費用) + 支(借費用 貸銀行)
                                {
                                    foreach (var payCash in vwList)
                                    {
                                        vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                        vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                        vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                        vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                        vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                        vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                        vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                        vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                        vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                        vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                        vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                        vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                        vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                        vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                        vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                        vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                        vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                        vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                        try
                                        {
                                            isLog = dao.FindLog(vw_GBCVisaDetail);
                                            isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                            if ((isLog > 0) && isPass.Equals("1"))
                                            {
                                                return "此筆資料已轉入過,並且結案。";
                                            }
                                            else if ((isLog > 0) && isPass.Equals("0"))
                                            {
                                                dao.Update(vw_GBCVisaDetail);
                                                jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                            }
                                            else
                                            {
                                                dao.Insert(vw_GBCVisaDetail);
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            return e.Message;
                                        }

                                        傳票明細 vouDtl_C = new 傳票明細()
                                        {
                                            借貸別 = "貸",
                                            科目代號 = "1154",
                                            科目名稱 = "預付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = prePayBalance,
                                            計畫代碼 = "",
                                            用途別代碼 = "",
                                            沖轉字號 = abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt), //沖轉支出傳票 from prePayNouNoList
                                            對象代碼 = "",
                                            對象說明 = ""
                                        };
                                        vouDtlList.Add(vouDtl_C);
                                        傳票明細 vouDtl_D = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "2125",
                                            科目名稱 = "應付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = estimateBalance,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = abateEstimateVouNo.ElementAt(abateCnt) + "-" + abateEstimateVouDtlNo.ElementAt(abateCnt),
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList.Add(vouDtl_D);
                                        傳票明細 vouDtl_D2 = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "5",
                                            科目名稱 = "基金用途",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = prePayBalance - estimateBalance,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = "",
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList.Add(vouDtl_D2);
                                        傳票受款人 vouPay = new 傳票受款人()
                                        {
                                            統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                            受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                            地址 = "",
                                            實付金額 = prePayBalance,
                                            銀行代號 = "",
                                            銀行名稱 = "",
                                            銀行帳號 = "",
                                            帳戶名稱 = ""
                                        };
                                        vouPayList.Add(vouPay);

                                        abateCnt++;

                                        //填傳票明細號1
                                        //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                    }
                                    //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                    var vouPayGroup = from xxx in vouPayList
                                                      group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                      select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                    vouPayList = new List<傳票受款人>();
                                    foreach (var vouPayGroupItem in vouPayGroup)
                                    {
                                        傳票受款人 vouPay = new 傳票受款人();
                                        vouPay.統一編號 = vouPayGroupItem.統一編號;
                                        vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                        vouPay.地址 = vouPayGroupItem.地址;
                                        vouPay.實付金額 = vouPayGroupItem.實付金額;
                                        vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                        vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                        vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                        vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                        vouPayList.Add(vouPay);
                                    }
                                    傳票主檔 vouMain = new 傳票主檔()
                                    {
                                        傳票種類 = "4",
                                        製票日期 = vw_GBCVisaDetail.F_製票日,
                                        主摘要 = vw_GBCVisaDetail.F_摘要,
                                        交付方式 = "1"
                                    };

                                    傳票內容 vouCollection = new 傳票內容()
                                    {
                                        傳票主檔 = vouMain,
                                        傳票明細 = vouDtlList,
                                        傳票受款人 = vouPayList
                                    };

                                    vouCollectionList.Add(vouCollection);
                                    vouTop = new 最外層()
                                    {
                                        基金代碼 = vw_GBCVisaDetail.基金代碼,
                                        年度 = vw_GBCVisaDetail.PK_會計年度,
                                        動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                        種類 = vw_GBCVisaDetail.PK_種類,
                                        次別 = vw_GBCVisaDetail.PK_次別,
                                        明細號 = vw_GBCVisaDetail.PK_明細號,
                                        傳票內容 = vouCollectionList
                                    };
                                    //------支出傳票------
                                    foreach (var payCash in vwList)
                                    {
                                        vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                        vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                        vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                        vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                        vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                        vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                        vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                        vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                        vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                        vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                        vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                        vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                        vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                        vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                        vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                        vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                        vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                        vw_GBCVisaDetail.F_批號 = payCash.F_批號;
                                        傳票明細 vouDtl_D2 = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "5",
                                            科目名稱 = "基金用途",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = "",
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList2.Add(vouDtl_D2);
                                        傳票受款人 vouPay2 = new 傳票受款人()
                                        {
                                            統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                            受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                            地址 = "",
                                            實付金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                            銀行代號 = "",
                                            銀行名稱 = "",
                                            銀行帳號 = "",
                                            帳戶名稱 = ""
                                        };
                                        vouPayList2.Add(vouPay2);

                                        //填傳票明細號2
                                        dao.FillVouDtl2(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                    }
                                    //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                    var vouPayGroup2 = from xxx in vouPayList2
                                                      group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                      select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                    vouPayList2 = new List<傳票受款人>();
                                    foreach (var vouPayGroupItem in vouPayGroup2)
                                    {
                                        傳票受款人 vouPay = new 傳票受款人();
                                        vouPay.統一編號 = vouPayGroupItem.統一編號;
                                        vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                        vouPay.地址 = vouPayGroupItem.地址;
                                        vouPay.實付金額 = vouPayGroupItem.實付金額;
                                        vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                        vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                        vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                        vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                        vouPayList2.Add(vouPay);
                                    }
                                    傳票主檔 vouMain2 = new 傳票主檔()
                                    {
                                        傳票種類 = vouKind,
                                        製票日期 = vw_GBCVisaDetail.F_製票日,
                                        主摘要 = vw_GBCVisaDetail.F_摘要,
                                        交付方式 = "1"
                                    };
                                    傳票明細 vouDtl_C2 = new 傳票明細()
                                    {
                                        借貸別 = "貸",
                                        科目代號 = "1112",
                                        科目名稱 = "銀行存款",
                                        摘要 = vw_GBCVisaDetail.F_摘要,
                                        金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                        計畫代碼 = "",
                                        用途別代碼 = "",
                                        沖轉字號 = "",
                                        對象代碼 = "",
                                        對象說明 = ""
                                    };
                                    vouDtlList2.Add(vouDtl_C2);
                                    傳票內容 vouCollection2 = new 傳票內容()
                                    {
                                        傳票主檔 = vouMain2,
                                        傳票明細 = vouDtlList2,
                                        傳票受款人 = vouPayList2
                                    };
                                    vouCollectionList2.Add(vouCollection2);

                                    vouTop2 = new 最外層()
                                    {
                                        基金代碼 = vw_GBCVisaDetail.基金代碼,
                                        年度 = vw_GBCVisaDetail.PK_會計年度,
                                        動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                        種類 = vw_GBCVisaDetail.PK_種類,
                                        次別 = vw_GBCVisaDetail.PK_次別,
                                        明細號 = vw_GBCVisaDetail.PK_明細號,
                                        傳票內容 = vouCollectionList2
                                    };

                                    //紀錄第一張傳票底稿
                                    try
                                    {
                                        jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    //紀錄第二張傳票底稿
                                    try
                                    {
                                        jsonDAO.UpdateJsonRecord2(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop2));
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    return JsonConvert.SerializeObject(vouTop2);
                                }
                            }
                            else //分轉(貸預付 借應付) 無借方費用
                            {
                                if (vw_GBCVisaDetail.F_核定金額 < prePayBalance) //不用加開支出傳票
                                {
                                    //------分轉-------
                                    foreach (var payCash in vwList)
                                    {
                                        vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                        vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                        vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                        vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                        vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                        vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                        vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                        vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                        vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                        vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                        vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                        vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                        vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                        vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                        vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                        vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                        vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                        vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                        try
                                        {
                                            isLog = dao.FindLog(vw_GBCVisaDetail);
                                            isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                            if ((isLog > 0) && isPass.Equals("1"))
                                            {
                                                return "此筆資料已轉入過,並且結案。";
                                            }
                                            else if ((isLog > 0) && isPass.Equals("0"))
                                            {
                                                dao.Update(vw_GBCVisaDetail);
                                                jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                            }
                                            else
                                            {
                                                dao.Insert(vw_GBCVisaDetail);
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            return e.Message;
                                        }

                                        傳票明細 vouDtl_C = new 傳票明細()
                                        {
                                            借貸別 = "貸",
                                            科目代號 = "1154",
                                            科目名稱 = "預付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = vw_GBCVisaDetail.F_核定金額,
                                            計畫代碼 = "",
                                            用途別代碼 = "",
                                            沖轉字號 = abatePrePayVouNo.ElementAt(abateCnt) + "-" + abatePrePayVouDtlNo.ElementAt(abateCnt), //沖轉支出傳票 from prePayNouNoList
                                            對象代碼 = "",
                                            對象說明 = ""
                                        };
                                        vouDtlList.Add(vouDtl_C);
                                        傳票明細 vouDtl_D = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "2125",
                                            科目名稱 = "應付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = vw_GBCVisaDetail.F_核定金額,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = abateEstimateVouNo.ElementAt(abateCnt) + "-" + abateEstimateVouDtlNo.ElementAt(abateCnt),
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList.Add(vouDtl_D);
                                        傳票受款人 vouPay = new 傳票受款人()
                                        {
                                            統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                            受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                            地址 = "",
                                            實付金額 = vw_GBCVisaDetail.F_核定金額,
                                            銀行代號 = "",
                                            銀行名稱 = "",
                                            銀行帳號 = "",
                                            帳戶名稱 = ""
                                        };
                                        vouPayList.Add(vouPay);

                                        abateCnt++;

                                        //填傳票明細號1
                                        //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                    }
                                    //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                    var vouPayGroup = from xxx in vouPayList
                                                      group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                      select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                    vouPayList = new List<傳票受款人>();
                                    foreach (var vouPayGroupItem in vouPayGroup)
                                    {
                                        傳票受款人 vouPay = new 傳票受款人();
                                        vouPay.統一編號 = vouPayGroupItem.統一編號;
                                        vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                        vouPay.地址 = vouPayGroupItem.地址;
                                        vouPay.實付金額 = vouPayGroupItem.實付金額;
                                        vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                        vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                        vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                        vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                        vouPayList.Add(vouPay);
                                    }
                                    傳票主檔 vouMain = new 傳票主檔()
                                    {
                                        傳票種類 = "4",
                                        製票日期 = vw_GBCVisaDetail.F_製票日,
                                        主摘要 = vw_GBCVisaDetail.F_摘要,
                                        交付方式 = "1"
                                    };

                                    傳票內容 vouCollection = new 傳票內容()
                                    {
                                        傳票主檔 = vouMain,
                                        傳票明細 = vouDtlList,
                                        傳票受款人 = vouPayList
                                    };

                                    vouCollectionList.Add(vouCollection);
                                    vouTop = new 最外層()
                                    {
                                        基金代碼 = vw_GBCVisaDetail.基金代碼,
                                        年度 = vw_GBCVisaDetail.PK_會計年度,
                                        動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                        種類 = vw_GBCVisaDetail.PK_種類,
                                        次別 = vw_GBCVisaDetail.PK_次別,
                                        明細號 = vw_GBCVisaDetail.PK_明細號,
                                        傳票內容 = vouCollectionList
                                    };

                                    //紀錄第一張傳票底稿
                                    try
                                    {
                                        jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    //回傳第一張傳票底稿
                                    JSON1 = jsonDAO.FindJSON1(vw_GBCVisaDetail);

                                    return JSON1;
                                }
                                else //要支出傳票
                                {
                                    //------分轉-------
                                    foreach (var payCash in vwList)
                                    {
                                        vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                        vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                        vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                        vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                        vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                        vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                        vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                        vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                        vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                        vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                        vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                        vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                        vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                        vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                        vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                        vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                        vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                        vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                                        try
                                        {
                                            isLog = dao.FindLog(vw_GBCVisaDetail);
                                            isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                            if ((isLog > 0) && isPass.Equals("1"))
                                            {
                                                return "此筆資料已轉入過,並且結案。";
                                            }
                                            else if ((isLog > 0) && isPass.Equals("0"))
                                            {
                                                dao.Update(vw_GBCVisaDetail);
                                                jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                            }
                                            else
                                            {
                                                dao.Insert(vw_GBCVisaDetail);
                                            }
                                        }
                                        catch (Exception e)
                                        {
                                            return e.Message;
                                        }

                                        傳票明細 vouDtl_C = new 傳票明細()
                                        {
                                            借貸別 = "貸",
                                            科目代號 = "1154",
                                            科目名稱 = "預付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = prePayBalance,
                                            計畫代碼 = "",
                                            用途別代碼 = "",
                                            沖轉字號 = "",
                                            對象代碼 = "",
                                            對象說明 = ""
                                        };
                                        vouDtlList.Add(vouDtl_C);
                                        傳票明細 vouDtl_D = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "2125",
                                            科目名稱 = "應付費用",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = prePayBalance,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = "",
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList.Add(vouDtl_D);
                                        傳票受款人 vouPay = new 傳票受款人()
                                        {
                                            統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                            受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                            地址 = "",
                                            實付金額 = prePayBalance,
                                            銀行代號 = "",
                                            銀行名稱 = "",
                                            銀行帳號 = "",
                                            帳戶名稱 = ""
                                        };
                                        vouPayList.Add(vouPay);

                                        //填傳票明細號1
                                        //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                    }
                                    //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                    var vouPayGroup = from xxx in vouPayList
                                                      group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                      select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                    vouPayList = new List<傳票受款人>();
                                    foreach (var vouPayGroupItem in vouPayGroup)
                                    {
                                        傳票受款人 vouPay = new 傳票受款人();
                                        vouPay.統一編號 = vouPayGroupItem.統一編號;
                                        vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                        vouPay.地址 = vouPayGroupItem.地址;
                                        vouPay.實付金額 = vouPayGroupItem.實付金額;
                                        vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                        vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                        vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                        vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                        vouPayList.Add(vouPay);
                                    }
                                    傳票主檔 vouMain = new 傳票主檔()
                                    {
                                        傳票種類 = "4",
                                        製票日期 = vw_GBCVisaDetail.F_製票日,
                                        主摘要 = vw_GBCVisaDetail.F_摘要,
                                        交付方式 = "1"
                                    };

                                    傳票內容 vouCollection = new 傳票內容()
                                    {
                                        傳票主檔 = vouMain,
                                        傳票明細 = vouDtlList,
                                        傳票受款人 = vouPayList
                                    };

                                    vouCollectionList.Add(vouCollection);
                                    vouTop = new 最外層()
                                    {
                                        基金代碼 = vw_GBCVisaDetail.基金代碼,
                                        年度 = vw_GBCVisaDetail.PK_會計年度,
                                        動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                        種類 = vw_GBCVisaDetail.PK_種類,
                                        次別 = vw_GBCVisaDetail.PK_次別,
                                        明細號 = vw_GBCVisaDetail.PK_明細號,
                                        傳票內容 = vouCollectionList
                                    };
                                    //------支出傳票------
                                    foreach (var payCash in vwList)
                                    {
                                        vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                                        vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                                        vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                                        vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                                        vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                                        vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                                        vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                                        vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                                        vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                                        vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                                        vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                                        vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                                        vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                                        vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                                        vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                                        vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                                        vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                                        vw_GBCVisaDetail.F_批號 = payCash.F_批號;
                                        傳票明細 vouDtl_D2 = new 傳票明細()
                                        {
                                            借貸別 = "借",
                                            科目代號 = "5",
                                            科目名稱 = "基金用途",
                                            摘要 = vw_GBCVisaDetail.F_摘要,
                                            金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                            計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                            用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                            沖轉字號 = "",
                                            對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                            對象說明 = vw_GBCVisaDetail.F_受款人,
                                            明細號 = vw_GBCVisaDetail.PK_明細號
                                        };
                                        vouDtlList2.Add(vouDtl_D2);
                                        傳票受款人 vouPay2 = new 傳票受款人()
                                        {
                                            統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                            受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                            地址 = "",
                                            實付金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                            銀行代號 = "",
                                            銀行名稱 = "",
                                            銀行帳號 = "",
                                            帳戶名稱 = ""
                                        };
                                        vouPayList2.Add(vouPay2);

                                        //填傳票明細號2
                                        dao.FillVouDtl2(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                                    }
                                    //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                                    var vouPayGroup2 = from xxx in vouPayList2
                                                      group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                                      select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                                    vouPayList2 = new List<傳票受款人>();
                                    foreach (var vouPayGroupItem in vouPayGroup2)
                                    {
                                        傳票受款人 vouPay = new 傳票受款人();
                                        vouPay.統一編號 = vouPayGroupItem.統一編號;
                                        vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                                        vouPay.地址 = vouPayGroupItem.地址;
                                        vouPay.實付金額 = vouPayGroupItem.實付金額;
                                        vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                                        vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                                        vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                                        vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                                        vouPayList2.Add(vouPay);
                                    }
                                    傳票主檔 vouMain2 = new 傳票主檔()
                                    {
                                        傳票種類 = vouKind,
                                        製票日期 = vw_GBCVisaDetail.F_製票日,
                                        主摘要 = vw_GBCVisaDetail.F_摘要,
                                        交付方式 = "1"
                                    };
                                    傳票明細 vouDtl_C2 = new 傳票明細()
                                    {
                                        借貸別 = "貸",
                                        科目代號 = "1112",
                                        科目名稱 = "銀行存款",
                                        摘要 = vw_GBCVisaDetail.F_摘要,
                                        金額 = vw_GBCVisaDetail.F_核定金額 - prePayBalance,
                                        計畫代碼 = "",
                                        用途別代碼 = "",
                                        沖轉字號 = "",
                                        對象代碼 = "",
                                        對象說明 = ""
                                    };
                                    vouDtlList2.Add(vouDtl_C2);
                                    傳票內容 vouCollection2 = new 傳票內容()
                                    {
                                        傳票主檔 = vouMain2,
                                        傳票明細 = vouDtlList2,
                                        傳票受款人 = vouPayList2
                                    };
                                    vouCollectionList2.Add(vouCollection2);

                                    vouTop2 = new 最外層()
                                    {
                                        基金代碼 = vw_GBCVisaDetail.基金代碼,
                                        年度 = vw_GBCVisaDetail.PK_會計年度,
                                        動支編號 = vw_GBCVisaDetail.PK_動支編號,
                                        種類 = vw_GBCVisaDetail.PK_種類,
                                        次別 = vw_GBCVisaDetail.PK_次別,
                                        明細號 = vw_GBCVisaDetail.PK_明細號,
                                        傳票內容 = vouCollectionList2
                                    };

                                    //紀錄第一張傳票底稿
                                    try
                                    {
                                        jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    //紀錄第二張傳票底稿
                                    try
                                    {
                                        jsonDAO.UpdateJsonRecord2(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop2));
                                    }
                                    catch (Exception e)
                                    {
                                        return e.Message;
                                    }

                                    return JsonConvert.SerializeObject(vouTop2);
                                }
                            }
                        }
                        return JsonConvert.SerializeObject("xxx");
                    }
                    //無預付立沖(執行實支)
                    else
                    {
                        foreach (var payCash in vwList)
                        {
                            vw_GBCVisaDetail.PK_會計年度 = payCash.PK_會計年度;
                            vw_GBCVisaDetail.PK_動支編號 = payCash.PK_動支編號;
                            vw_GBCVisaDetail.PK_種類 = payCash.PK_種類;
                            vw_GBCVisaDetail.PK_次別 = payCash.PK_次別;
                            vw_GBCVisaDetail.PK_明細號 = payCash.PK_明細號;
                            vw_GBCVisaDetail.F_科室代碼 = payCash.F_科室代碼;
                            vw_GBCVisaDetail.F_用途別代碼 = payCash.F_用途別代碼;
                            vw_GBCVisaDetail.F_計畫代碼 = payCash.F_計畫代碼;
                            vw_GBCVisaDetail.F_動支金額 = payCash.F_動支金額;
                            vw_GBCVisaDetail.F_製票日 = payCash.F_製票日;
                            vw_GBCVisaDetail.F_是否核定 = payCash.F_是否核定;
                            vw_GBCVisaDetail.F_核定金額 = payCash.F_核定金額;
                            vw_GBCVisaDetail.F_核定日期 = payCash.F_核定日期;
                            vw_GBCVisaDetail.F_摘要 = payCash.F_摘要;
                            vw_GBCVisaDetail.F_受款人 = payCash.F_受款人;
                            vw_GBCVisaDetail.F_受款人編號 = payCash.F_受款人編號;
                            vw_GBCVisaDetail.F_原動支編號 = payCash.F_原動支編號;
                            vw_GBCVisaDetail.F_批號 = payCash.F_批號;

                            try
                            {
                                isLog = dao.FindLog(vw_GBCVisaDetail);
                                isPass = jsonDAO.IsPass(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別);
                                if ((isLog > 0) && isPass.Equals("1"))
                                {
                                    return "此筆資料已轉入過,並且結案。";
                                }
                                else if ((isLog > 0) && isPass.Equals("0"))
                                {
                                    dao.Update(vw_GBCVisaDetail);
                                    jsonDAO.DeleteJsonRecord1(vw_GBCVisaDetail);
                                }
                                else
                                {
                                    dao.Insert(vw_GBCVisaDetail);
                                }
                            }
                            catch (Exception e)
                            {
                                return e.Message;
                            }

                            傳票明細 vouDtl_D = new 傳票明細()
                            {
                                借貸別 = "借",
                                科目代號 = "5",
                                科目名稱 = "基金用途",
                                摘要 = vw_GBCVisaDetail.F_摘要,
                                金額 = vw_GBCVisaDetail.F_核定金額,
                                計畫代碼 = vw_GBCVisaDetail.F_計畫代碼,
                                用途別代碼 = vw_GBCVisaDetail.F_用途別代碼,
                                沖轉字號 = "",
                                對象代碼 = vw_GBCVisaDetail.F_受款人編號,
                                對象說明 = vw_GBCVisaDetail.F_受款人,
                                明細號 = vw_GBCVisaDetail.PK_明細號
                            };

                            //是否為以前年度
                            if (int.Parse(vw_GBCVisaDetail.PK_動支編號.Substring(0, 3)) < int.Parse(vw_GBCVisaDetail.PK_會計年度))
                            {
                                vouDtl_D.用途別代碼 = "91Y";
                            }
                            vouDtlList.Add(vouDtl_D);
                            傳票受款人 vouPay = new 傳票受款人()
                            {
                                統一編號 = vw_GBCVisaDetail.F_受款人編號,
                                受款人名稱 = vw_GBCVisaDetail.F_受款人,
                                地址 = "",
                                實付金額 = vw_GBCVisaDetail.F_核定金額,
                                銀行代號 = "",
                                銀行名稱 = "",
                                銀行帳號 = "",
                                帳戶名稱 = ""
                            };
                            vouPayList.Add(vouPay);

                            //填傳票明細號1
                            //dao.FillVouDtl1(vw_GBCVisaDetail.基金代碼, vw_GBCVisaDetail.PK_會計年度, vw_GBCVisaDetail.PK_動支編號, vw_GBCVisaDetail.PK_種類, vw_GBCVisaDetail.PK_次別, vw_GBCVisaDetail.PK_明細號, vouDtlList.Count);
                        }
                        //重新處理受款人清單,如果有重複受款人名稱,則金額加總
                        var vouPayGroup = from xxx in vouPayList
                                          group xxx by new { xxx.統一編號, xxx.受款人名稱, xxx.地址, xxx.銀行代號, xxx.銀行名稱, xxx.銀行帳號, xxx.帳戶名稱 } into g
                                          select new { 統一編號 = g.Key.統一編號, 受款人名稱 = g.Key.受款人名稱, 地址 = g.Key.地址, 銀行代號 = g.Key.銀行代號, 銀行名稱 = g.Key.銀行名稱, 銀行帳號 = g.Key.銀行帳號, 帳戶名稱 = g.Key.帳戶名稱, 實付金額 = g.Sum(xxx => xxx.實付金額) };
                        vouPayList = new List<傳票受款人>();
                        foreach (var vouPayGroupItem in vouPayGroup)
                        {
                            傳票受款人 vouPay = new 傳票受款人();
                            vouPay.統一編號 = vouPayGroupItem.統一編號;
                            vouPay.受款人名稱 = vouPayGroupItem.受款人名稱;
                            vouPay.地址 = vouPayGroupItem.地址;
                            vouPay.實付金額 = vouPayGroupItem.實付金額;
                            vouPay.銀行代號 = vouPayGroupItem.銀行代號;
                            vouPay.銀行名稱 = vouPayGroupItem.銀行名稱;
                            vouPay.銀行帳號 = vouPayGroupItem.銀行帳號;
                            vouPay.帳戶名稱 = vouPayGroupItem.帳戶名稱;
                            vouPayList.Add(vouPay);
                        }
                        傳票主檔 vouMain = new 傳票主檔()
                        {
                            傳票種類 = vouKind,
                            製票日期 = vw_GBCVisaDetail.F_製票日,
                            主摘要 = vw_GBCVisaDetail.F_摘要,
                            交付方式 = "1"
                        };
                        傳票明細 vouDtl_C = new 傳票明細()
                        {
                            借貸別 = "貸",
                            科目代號 = "1112",
                            科目名稱 = "銀行存款",
                            摘要 = vw_GBCVisaDetail.F_摘要,
                            金額 = accSumMoney,
                            計畫代碼 = "",
                            用途別代碼 = "",
                            沖轉字號 = "",
                            對象代碼 = "",
                            對象說明 = ""
                        };
                        vouDtlList.Add(vouDtl_C);
                        傳票內容 vouCollection = new 傳票內容()
                        {
                            傳票主檔 = vouMain,
                            傳票明細 = vouDtlList,
                            傳票受款人 = vouPayList
                        };

                        vouCollectionList.Add(vouCollection);
                        vouTop = new 最外層()
                        {
                            基金代碼 = vw_GBCVisaDetail.基金代碼,
                            年度 = vw_GBCVisaDetail.PK_會計年度,
                            動支編號 = vw_GBCVisaDetail.PK_動支編號,
                            種類 = vw_GBCVisaDetail.PK_種類,
                            次別 = vw_GBCVisaDetail.PK_次別,
                            明細號 = vw_GBCVisaDetail.PK_明細號,
                            傳票內容 = vouCollectionList
                        };
                    }
                    //紀錄第一張傳票底稿
                    try
                    {
                        jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
                    }
                    catch (Exception e)
                    {
                        return e.Message;
                    }
                    //回傳第一張傳票底稿
                    JSON1 = jsonDAO.FindJSON1(vw_GBCVisaDetail);

                    return JSON1;
                }

            }

            //紀錄第一張傳票底稿
            try
            {
                jsonDAO.InsertJsonRecord1(vw_GBCVisaDetail, JsonConvert.SerializeObject(vouTop));
            }
            catch (Exception e)
            {
                return e.Message;
            }
            //回傳第一張傳票底稿
            JSON1 = jsonDAO.FindJSON1(vw_GBCVisaDetail);

            return JSON1;
        }

        [WebMethod]
        public string FillVouNo(string vouNoJSON)
        {
            GBCVisaDetailAbateDetailDAO dao = new GBCVisaDetailAbateDetailDAO();
            GBCJSONRecordDAO jsonDAO = new GBCJSONRecordDAO();
            FillVouScript fillVouScript = new FillVouScript();
            GBCJSONRecordVO gbcJSONRecordVO = new GBCJSONRecordVO();
            GBCVisaDetailAbateDetailVO gbcVisaDetailAbateDetailVO = new GBCVisaDetailAbateDetailVO();

            string isVouNo1 = "";
            string isJSON2 = "";
            string isPass = "";
            int count = 0;

            try
            {
                fillVouScript = JsonConvert.DeserializeObject<FillVouScript>(vouNoJSON);  //反序列化JSON
            }
            catch (Exception e)
            {
                return e.StackTrace;
            }
           
            isPass = jsonDAO.IsPass(fillVouScript.基金代碼, fillVouScript.會計年度, fillVouScript.動支編號, fillVouScript.種類, fillVouScript.次別);
            isJSON2 = jsonDAO.FindJSON2(fillVouScript.基金代碼, fillVouScript.會計年度, fillVouScript.動支編號, fillVouScript.種類, fillVouScript.次別);

            gbcVisaDetailAbateDetailVO.set基金代碼(fillVouScript.基金代碼);
            gbcVisaDetailAbateDetailVO.setPK_會計年度(fillVouScript.會計年度);
            gbcVisaDetailAbateDetailVO.setPK_動支編號(fillVouScript.動支編號);
            gbcVisaDetailAbateDetailVO.setPK_種類(fillVouScript.種類);
            gbcVisaDetailAbateDetailVO.setPK_次別(fillVouScript.次別);
            gbcVisaDetailAbateDetailVO.setF_傳票年度(fillVouScript.傳票年度);
            gbcVisaDetailAbateDetailVO.setF_傳票號1(fillVouScript.傳票號);
            gbcVisaDetailAbateDetailVO.setF_製票日期1(fillVouScript.製票日期);

            foreach (var 傳票明細Item in fillVouScript.傳票明細)
            {
                isVouNo1 = dao.FindVouNo(fillVouScript.基金代碼, fillVouScript.會計年度, fillVouScript.動支編號, fillVouScript.種類, fillVouScript.次別, 傳票明細Item.傳票明細號);
                gbcVisaDetailAbateDetailVO.setPK_明細號(傳票明細Item.明細號);
                gbcVisaDetailAbateDetailVO.setF_傳票明細號1(傳票明細Item.傳票明細號);

                if (((isVouNo1.Trim()).Length == 0) && (isPass == "0")) //傳票1未回填 AND 未結案 --回填至傳票1
                {
                    dao.UpdateVouNo1(gbcVisaDetailAbateDetailVO);
                    count++;
                    if ((isJSON2.Trim().Length == 0) && (count == fillVouScript.傳票明細.Count))
                    {
                        jsonDAO.UpdatePassFlg(fillVouScript.基金代碼, fillVouScript.會計年度, fillVouScript.動支編號, fillVouScript.種類, fillVouScript.次別);
                    }
                }
                else if (((isVouNo1.Trim()).Length != 0) && (isPass == "0"))//傳票1已回填 AND 未結案 --回填至傳票2
                {
                    dao.UpdateVouNo2(gbcVisaDetailAbateDetailVO);
                    jsonDAO.UpdatePassFlg(fillVouScript.基金代碼, fillVouScript.會計年度, fillVouScript.動支編號, fillVouScript.種類, fillVouScript.次別);
                }
                else
                {
                    return fillVouScript.動支編號 + "-" + fillVouScript.種類 + "-" + fillVouScript.次別 + "...回填失敗!  請確認是否已回填完畢。";
                }

                //傳票號回寫至預控系統
                //由Web.Config來開關是否回填至預控系統
                string isFillToGBC = WebConfigurationManager.AppSettings["isFillToGBC"];

                if ((isFillToGBC.Trim()).Equals("1"))
                {
                    //判斷基金代號,回填至對應的預控系統(GBC)
                    if (gbcVisaDetailAbateDetailVO.get基金代碼() == "010")//醫發服務參考
                    {
                        GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetailVO.getPK_會計年度(), gbcVisaDetailAbateDetailVO.getPK_動支編號(), gbcVisaDetailAbateDetailVO.getPK_種類(), gbcVisaDetailAbateDetailVO.getPK_次別(), gbcVisaDetailAbateDetailVO.getPK_明細號(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1());
                    }
                    else if (gbcVisaDetailAbateDetailVO.get基金代碼() == "040")//菸害****尚未加入服務參考****
                    {

                    }
                    else if (gbcVisaDetailAbateDetailVO.get基金代碼() == "090")//家防服務參考
                    {
                        DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetailVO.getPK_會計年度(), gbcVisaDetailAbateDetailVO.getPK_動支編號(), gbcVisaDetailAbateDetailVO.getPK_種類(), gbcVisaDetailAbateDetailVO.getPK_次別(), gbcVisaDetailAbateDetailVO.getPK_明細號(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1());
                    }
                    else if (gbcVisaDetailAbateDetailVO.get基金代碼() == "100")//長照
                    {
                        LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetailVO.getPK_會計年度(), gbcVisaDetailAbateDetailVO.getPK_動支編號(), gbcVisaDetailAbateDetailVO.getPK_種類(), gbcVisaDetailAbateDetailVO.getPK_次別(), gbcVisaDetailAbateDetailVO.getPK_明細號(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1());
                    }
                    else if (gbcVisaDetailAbateDetailVO.get基金代碼() == "110")//生產
                    {
                        BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                        ws.FillVouNo(gbcVisaDetailAbateDetailVO.getPK_會計年度(), gbcVisaDetailAbateDetailVO.getPK_動支編號(), gbcVisaDetailAbateDetailVO.getPK_種類(), gbcVisaDetailAbateDetailVO.getPK_次別(), gbcVisaDetailAbateDetailVO.getPK_明細號(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1(), gbcVisaDetailAbateDetailVO.getF_傳票號1(), gbcVisaDetailAbateDetailVO.getF_製票日期1());
                    }
                }
            }
         
            return "回填完畢";
        }
         
        //傳票號碼回填至GBC(輸入條碼)
        public string FillVouNo2(string fundNo, string acmWordNum, string vouYear, string makeVouNo, string makeVouDate)
        {
            GBCVisaDetailAbateDetailDAO dao = new GBCVisaDetailAbateDetailDAO();
            GBCJSONRecordDAO jsonDAO = new GBCJSONRecordDAO();
            List<GBCVisaDetailAbateDetailVO> gbcList = new List<GBCVisaDetailAbateDetailVO>();
            string isVouNo1 = "";
            string isJSON2 = "";
            string isPass = "";
            int count = 0;
            int vouDtlCnt = 1;

            gbcList = dao.FindGBCReco(fundNo, acmWordNum);

            foreach (var gbcListItem in gbcList)
            {
                isVouNo1 = dao.FindVouNo(gbcListItem.get基金代碼(), gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別(), gbcListItem.getPK_明細號());
                isPass = jsonDAO.IsPass(gbcListItem.get基金代碼(), gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別());
                isJSON2 = jsonDAO.FindJSON2(gbcListItem.get基金代碼(), gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別());

                gbcListItem.setF_傳票年度(vouYear);
                gbcListItem.setF_傳票號1(makeVouNo);
                //gbcListItem.setF_傳票明細號1(vouDtlCnt.ToString());
                gbcListItem.setF_製票日期1(makeVouDate);

                vouDtlCnt++;

                if (((isVouNo1.Trim()).Length == 0) && (isPass == "0")) //傳票1未回填 AND 未結案 --回填至傳票1
                {
                    dao.UpdateVouNo1(gbcListItem);
                    count++;
                    if ((isJSON2.Trim().Length == 0) && (count == gbcList.Count))
                    {
                        jsonDAO.UpdatePassFlg(gbcListItem.get基金代碼(), gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別());
                    }
                }
                else if (((isVouNo1.Trim()).Length != 0) && (isPass == "0"))//傳票1已回填 AND 未結案 --回填至傳票2
                {
                    dao.UpdateVouNo2(gbcListItem);
                    jsonDAO.UpdatePassFlg(gbcListItem.get基金代碼(), gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別());
                }
                else
                {
                    return gbcListItem.getPK_動支編號() + "-" + gbcListItem.getPK_種類() + "-" + gbcListItem.getPK_次別() + "...回填失敗!  請確認是否已回填。";
                }

            }

            //由Web.Config來開關是否回填至預控系統
            string isFillToGBC = WebConfigurationManager.AppSettings["isFillToGBC"];

            if ((isFillToGBC.Trim()).Equals("1"))
            {
                //判斷基金代號,回填至對應的預控系統(GBC)
                if (fundNo == "010")//醫發服務參考
                {
                    GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                    foreach (var gbcListItem in gbcList)
                    {
                        ws.FillVouNo(gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別(), gbcListItem.getPK_明細號(), gbcListItem.getF_傳票號1(), gbcListItem.getF_製票日期1(), gbcListItem.getF_傳票號1(), gbcListItem.getF_製票日期1());
                    }

                }
                else if (fundNo == "040")//菸害****尚未加入服務參考****
                {

                }
                else if (fundNo == "090")//家防服務參考
                {
                    DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                    foreach (var gbcListItem in gbcList)
                    {
                        ws.FillVouNo(gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別(), gbcListItem.getPK_明細號(), gbcListItem.getF_傳票號1(), gbcListItem.getF_製票日期1(), gbcListItem.getF_傳票號1(), gbcListItem.getF_製票日期1());
                    }

                }
                else if (fundNo == "100")//長照****尚未加入服務參考****
                {

                }
                else if (fundNo == "110")//生產****尚未加入服務參考****
                {
                    BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                    foreach (var gbcListItem in gbcList)
                    {
                        ws.FillVouNo(gbcListItem.getPK_會計年度(), gbcListItem.getPK_動支編號(), gbcListItem.getPK_種類(), gbcListItem.getPK_次別(), gbcListItem.getPK_明細號(), gbcListItem.getF_傳票號1(), gbcListItem.getF_製票日期1(), gbcListItem.getF_傳票號1(), gbcListItem.getF_製票日期1());
                    }
                }
            }
            
            return "回填完畢";
        }

        //==================手動搜尋功能===================
        //1.public List<string> GetYear(string fundNo)
        //2.public List<string> GetAcmWordNum(string fundNo, string accYear)
        //3.public List<string> GetAccKind(string fundNo, string accYear, string acmWordNum)
        //4.public List<string> GetAccCount(string fundNo, string accYear, string acmWordNum, string accKind)
        //5.public List<string> GetAccDetail(string fundNo, string accYear, string acmWordNum, string accKind, string accCount)
        //6.public string GetByPrimaryKey(string fundNo, string accYear, string acmWordNum, string accKind, string accCount, string accDetail)

        [WebMethod]
        //取年度
        public List<string> GetYear(string fundNo)
        {

            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                List<string> yearList = new List<string>(ws.GetYear());

                return yearList;
            }
            else
            {
                return null;
            }

        }

        [WebMethod]
        //取動支號
        public List<string> GetAcmWordNum(string fundNo, string accYear)
        {
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                List<string> acmNoList = new List<string>(
                    ws.GetAcmWordNum(accYear));

                return acmNoList;
            }
            else
            {
                return null;
            }

        }

        [WebMethod]
        //取種類
        public List<string> GetAccKind(string fundNo, string accYear, string acmWordNum)
        {
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                List<string> accKindList = new List<string>(
                    ws.GetAccKind(accYear, acmWordNum));
                return accKindList;
            }
            else
            {
                return null;
            }

        }

        [WebMethod]
        //取次數
        public List<string> GetAccCount(string fundNo, string accYear, string acmWordNum, string accKind)
        {
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccCount(accYear, acmWordNum, accKind));
                return accDetailList;
            }
            else
            {
                return null;
            }

        }

        [WebMethod]
        //取明細號
        public List<string> GetAccDetail(string fundNo, string accYear, string acmWordNum, string accKind, string accCount)
        {
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                List<string> accDetailList = new List<string>(
                    ws.GetAccDetail(accYear, acmWordNum, accKind, accCount));
                return accDetailList;
            }
            else
            {
                return null;
            }

        }

        [WebMethod]
        //依據KEY值取View
        public string GetByPrimaryKey(string fundNo, string accYear, string acmWordNum, string accKind, string accCount, string accDetail)
        {
            //先判斷基金代號
            if (fundNo == "010")//醫發服務參考
            {
                GBC_WebService.GBCWebService ws = new GBC_WebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else if (fundNo == "040")//菸害****尚未加入服務參考****
            {
                return null;
            }
            else if (fundNo == "090")//家防服務參考
            {
                DVGBC_WebService.GBCWebService ws = new DVGBC_WebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else if (fundNo == "100")//長照****尚未加入服務參考****
            {
                LCGBC_WebService.GBCWebService ws = new LCGBC_WebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else if (fundNo == "110")//生產****尚未加入服務參考****
            {
                BAGBC_WebService.GBCWebService ws = new BAGBC_WebService.GBCWebService();
                string getGBCVisaDetail = ws.GetByPrimaryKey(accYear, acmWordNum, accKind, accCount, accDetail);
                return getGBCVisaDetail;
            }
            else
            {
                return null;
            }

        }

    }
}
