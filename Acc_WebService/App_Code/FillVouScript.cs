﻿using System.Collections.Generic;
/// <summary>
/// FillVouScript 的摘要描述
/// </summary>
public class FillVouScript
{
    public class 回填明細
    {
        public string 明細號 { get; set; }
        public string 傳票明細號 { get; set; }
    }

    public string 傳票年度 { get; set; }
    public string 傳票號 { get; set; }
    public string 製票日期 { get; set; }
    public List<回填明細> 傳票明細 { get; set; }

}