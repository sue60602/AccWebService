﻿//------------------------------------------------------------------------------
// <auto-generated>
//     這段程式碼是由工具產生的。
//     執行階段版本:4.0.30319.42000
//
//     對這個檔案所做的變更可能會造成錯誤的行為，而且如果重新產生程式碼，
//     變更將會遺失。
// </auto-generated>
//------------------------------------------------------------------------------

// 
// 原始程式碼已由 Microsoft.VSDesigner 自動產生，版本 4.0.30319.42000。
// 
#pragma warning disable 1591

namespace Acc_WebService.GBC_WebService {
    using System;
    using System.Web.Services;
    using System.Diagnostics;
    using System.Web.Services.Protocols;
    using System.Xml.Serialization;
    using System.ComponentModel;
    
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Web.Services.WebServiceBindingAttribute(Name="GBCWebServiceSoap", Namespace="http://tempuri.org/")]
    public partial class GBCWebService : System.Web.Services.Protocols.SoapHttpClientProtocol {
        
        private System.Threading.SendOrPostCallback getVw_GBCVisaDetailOperationCompleted;
        
        private System.Threading.SendOrPostCallback getVw_GBCVisaDetailXMLOperationCompleted;
        
        private System.Threading.SendOrPostCallback getVw_GBCVisaDetailJSONOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetYearOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAcmWordNumOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAccKindOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAccCountOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetAccDetailOperationCompleted;
        
        private System.Threading.SendOrPostCallback GetByPrimaryKeyOperationCompleted;
        
        private bool useDefaultCredentialsSetExplicitly;
        
        /// <remarks/>
        public GBCWebService() {
            this.Url = global::Acc_WebService.Properties.Settings.Default.Acc_WebService_GBC_WebService_GBCWebService;
            if ((this.IsLocalFileSystemWebService(this.Url) == true)) {
                this.UseDefaultCredentials = true;
                this.useDefaultCredentialsSetExplicitly = false;
            }
            else {
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        public new string Url {
            get {
                return base.Url;
            }
            set {
                if ((((this.IsLocalFileSystemWebService(base.Url) == true) 
                            && (this.useDefaultCredentialsSetExplicitly == false)) 
                            && (this.IsLocalFileSystemWebService(value) == false))) {
                    base.UseDefaultCredentials = false;
                }
                base.Url = value;
            }
        }
        
        public new bool UseDefaultCredentials {
            get {
                return base.UseDefaultCredentials;
            }
            set {
                base.UseDefaultCredentials = value;
                this.useDefaultCredentialsSetExplicitly = true;
            }
        }
        
        /// <remarks/>
        public event getVw_GBCVisaDetailCompletedEventHandler getVw_GBCVisaDetailCompleted;
        
        /// <remarks/>
        public event getVw_GBCVisaDetailXMLCompletedEventHandler getVw_GBCVisaDetailXMLCompleted;
        
        /// <remarks/>
        public event getVw_GBCVisaDetailJSONCompletedEventHandler getVw_GBCVisaDetailJSONCompleted;
        
        /// <remarks/>
        public event GetYearCompletedEventHandler GetYearCompleted;
        
        /// <remarks/>
        public event GetAcmWordNumCompletedEventHandler GetAcmWordNumCompleted;
        
        /// <remarks/>
        public event GetAccKindCompletedEventHandler GetAccKindCompleted;
        
        /// <remarks/>
        public event GetAccCountCompletedEventHandler GetAccCountCompleted;
        
        /// <remarks/>
        public event GetAccDetailCompletedEventHandler GetAccDetailCompleted;
        
        /// <remarks/>
        public event GetByPrimaryKeyCompletedEventHandler GetByPrimaryKeyCompleted;
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/getVw_GBCVisaDetail", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string getVw_GBCVisaDetail(string acmWordNum) {
            object[] results = this.Invoke("getVw_GBCVisaDetail", new object[] {
                        acmWordNum});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void getVw_GBCVisaDetailAsync(string acmWordNum) {
            this.getVw_GBCVisaDetailAsync(acmWordNum, null);
        }
        
        /// <remarks/>
        public void getVw_GBCVisaDetailAsync(string acmWordNum, object userState) {
            if ((this.getVw_GBCVisaDetailOperationCompleted == null)) {
                this.getVw_GBCVisaDetailOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetVw_GBCVisaDetailOperationCompleted);
            }
            this.InvokeAsync("getVw_GBCVisaDetail", new object[] {
                        acmWordNum}, this.getVw_GBCVisaDetailOperationCompleted, userState);
        }
        
        private void OngetVw_GBCVisaDetailOperationCompleted(object arg) {
            if ((this.getVw_GBCVisaDetailCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getVw_GBCVisaDetailCompleted(this, new getVw_GBCVisaDetailCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/getVw_GBCVisaDetailXML", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public Vw_GBCVisaDetail[] getVw_GBCVisaDetailXML(string acmWordNum) {
            object[] results = this.Invoke("getVw_GBCVisaDetailXML", new object[] {
                        acmWordNum});
            return ((Vw_GBCVisaDetail[])(results[0]));
        }
        
        /// <remarks/>
        public void getVw_GBCVisaDetailXMLAsync(string acmWordNum) {
            this.getVw_GBCVisaDetailXMLAsync(acmWordNum, null);
        }
        
        /// <remarks/>
        public void getVw_GBCVisaDetailXMLAsync(string acmWordNum, object userState) {
            if ((this.getVw_GBCVisaDetailXMLOperationCompleted == null)) {
                this.getVw_GBCVisaDetailXMLOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetVw_GBCVisaDetailXMLOperationCompleted);
            }
            this.InvokeAsync("getVw_GBCVisaDetailXML", new object[] {
                        acmWordNum}, this.getVw_GBCVisaDetailXMLOperationCompleted, userState);
        }
        
        private void OngetVw_GBCVisaDetailXMLOperationCompleted(object arg) {
            if ((this.getVw_GBCVisaDetailXMLCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getVw_GBCVisaDetailXMLCompleted(this, new getVw_GBCVisaDetailXMLCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/getVw_GBCVisaDetailJSON", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string getVw_GBCVisaDetailJSON(string acmWordNum) {
            object[] results = this.Invoke("getVw_GBCVisaDetailJSON", new object[] {
                        acmWordNum});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void getVw_GBCVisaDetailJSONAsync(string acmWordNum) {
            this.getVw_GBCVisaDetailJSONAsync(acmWordNum, null);
        }
        
        /// <remarks/>
        public void getVw_GBCVisaDetailJSONAsync(string acmWordNum, object userState) {
            if ((this.getVw_GBCVisaDetailJSONOperationCompleted == null)) {
                this.getVw_GBCVisaDetailJSONOperationCompleted = new System.Threading.SendOrPostCallback(this.OngetVw_GBCVisaDetailJSONOperationCompleted);
            }
            this.InvokeAsync("getVw_GBCVisaDetailJSON", new object[] {
                        acmWordNum}, this.getVw_GBCVisaDetailJSONOperationCompleted, userState);
        }
        
        private void OngetVw_GBCVisaDetailJSONOperationCompleted(object arg) {
            if ((this.getVw_GBCVisaDetailJSONCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.getVw_GBCVisaDetailJSONCompleted(this, new getVw_GBCVisaDetailJSONCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetYear", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetYear() {
            object[] results = this.Invoke("GetYear", new object[0]);
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetYearAsync() {
            this.GetYearAsync(null);
        }
        
        /// <remarks/>
        public void GetYearAsync(object userState) {
            if ((this.GetYearOperationCompleted == null)) {
                this.GetYearOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetYearOperationCompleted);
            }
            this.InvokeAsync("GetYear", new object[0], this.GetYearOperationCompleted, userState);
        }
        
        private void OnGetYearOperationCompleted(object arg) {
            if ((this.GetYearCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetYearCompleted(this, new GetYearCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAcmWordNum", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAcmWordNum(string accYear) {
            object[] results = this.Invoke("GetAcmWordNum", new object[] {
                        accYear});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAcmWordNumAsync(string accYear) {
            this.GetAcmWordNumAsync(accYear, null);
        }
        
        /// <remarks/>
        public void GetAcmWordNumAsync(string accYear, object userState) {
            if ((this.GetAcmWordNumOperationCompleted == null)) {
                this.GetAcmWordNumOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAcmWordNumOperationCompleted);
            }
            this.InvokeAsync("GetAcmWordNum", new object[] {
                        accYear}, this.GetAcmWordNumOperationCompleted, userState);
        }
        
        private void OnGetAcmWordNumOperationCompleted(object arg) {
            if ((this.GetAcmWordNumCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAcmWordNumCompleted(this, new GetAcmWordNumCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAccKind", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAccKind(string accYear, string acmWordNum) {
            object[] results = this.Invoke("GetAccKind", new object[] {
                        accYear,
                        acmWordNum});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAccKindAsync(string accYear, string acmWordNum) {
            this.GetAccKindAsync(accYear, acmWordNum, null);
        }
        
        /// <remarks/>
        public void GetAccKindAsync(string accYear, string acmWordNum, object userState) {
            if ((this.GetAccKindOperationCompleted == null)) {
                this.GetAccKindOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAccKindOperationCompleted);
            }
            this.InvokeAsync("GetAccKind", new object[] {
                        accYear,
                        acmWordNum}, this.GetAccKindOperationCompleted, userState);
        }
        
        private void OnGetAccKindOperationCompleted(object arg) {
            if ((this.GetAccKindCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAccKindCompleted(this, new GetAccKindCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAccCount", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAccCount(string accYear, string acmWordNum, string accKind) {
            object[] results = this.Invoke("GetAccCount", new object[] {
                        accYear,
                        acmWordNum,
                        accKind});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAccCountAsync(string accYear, string acmWordNum, string accKind) {
            this.GetAccCountAsync(accYear, acmWordNum, accKind, null);
        }
        
        /// <remarks/>
        public void GetAccCountAsync(string accYear, string acmWordNum, string accKind, object userState) {
            if ((this.GetAccCountOperationCompleted == null)) {
                this.GetAccCountOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAccCountOperationCompleted);
            }
            this.InvokeAsync("GetAccCount", new object[] {
                        accYear,
                        acmWordNum,
                        accKind}, this.GetAccCountOperationCompleted, userState);
        }
        
        private void OnGetAccCountOperationCompleted(object arg) {
            if ((this.GetAccCountCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAccCountCompleted(this, new GetAccCountCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetAccDetail", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string[] GetAccDetail(string accYear, string acmWordNum, string accKind, string accCount) {
            object[] results = this.Invoke("GetAccDetail", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount});
            return ((string[])(results[0]));
        }
        
        /// <remarks/>
        public void GetAccDetailAsync(string accYear, string acmWordNum, string accKind, string accCount) {
            this.GetAccDetailAsync(accYear, acmWordNum, accKind, accCount, null);
        }
        
        /// <remarks/>
        public void GetAccDetailAsync(string accYear, string acmWordNum, string accKind, string accCount, object userState) {
            if ((this.GetAccDetailOperationCompleted == null)) {
                this.GetAccDetailOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetAccDetailOperationCompleted);
            }
            this.InvokeAsync("GetAccDetail", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount}, this.GetAccDetailOperationCompleted, userState);
        }
        
        private void OnGetAccDetailOperationCompleted(object arg) {
            if ((this.GetAccDetailCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetAccDetailCompleted(this, new GetAccDetailCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        [System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/GetByPrimaryKey", RequestNamespace="http://tempuri.org/", ResponseNamespace="http://tempuri.org/", Use=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)]
        public string GetByPrimaryKey(string accYear, string acmWordNum, string accKind, string accCount, string accDetail) {
            object[] results = this.Invoke("GetByPrimaryKey", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount,
                        accDetail});
            return ((string)(results[0]));
        }
        
        /// <remarks/>
        public void GetByPrimaryKeyAsync(string accYear, string acmWordNum, string accKind, string accCount, string accDetail) {
            this.GetByPrimaryKeyAsync(accYear, acmWordNum, accKind, accCount, accDetail, null);
        }
        
        /// <remarks/>
        public void GetByPrimaryKeyAsync(string accYear, string acmWordNum, string accKind, string accCount, string accDetail, object userState) {
            if ((this.GetByPrimaryKeyOperationCompleted == null)) {
                this.GetByPrimaryKeyOperationCompleted = new System.Threading.SendOrPostCallback(this.OnGetByPrimaryKeyOperationCompleted);
            }
            this.InvokeAsync("GetByPrimaryKey", new object[] {
                        accYear,
                        acmWordNum,
                        accKind,
                        accCount,
                        accDetail}, this.GetByPrimaryKeyOperationCompleted, userState);
        }
        
        private void OnGetByPrimaryKeyOperationCompleted(object arg) {
            if ((this.GetByPrimaryKeyCompleted != null)) {
                System.Web.Services.Protocols.InvokeCompletedEventArgs invokeArgs = ((System.Web.Services.Protocols.InvokeCompletedEventArgs)(arg));
                this.GetByPrimaryKeyCompleted(this, new GetByPrimaryKeyCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState));
            }
        }
        
        /// <remarks/>
        public new void CancelAsync(object userState) {
            base.CancelAsync(userState);
        }
        
        private bool IsLocalFileSystemWebService(string url) {
            if (((url == null) 
                        || (url == string.Empty))) {
                return false;
            }
            System.Uri wsUri = new System.Uri(url);
            if (((wsUri.Port >= 1024) 
                        && (string.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) == 0))) {
                return true;
            }
            return false;
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Xml", "4.6.1055.0")]
    [System.SerializableAttribute()]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(Namespace="http://tempuri.org/")]
    public partial class Vw_GBCVisaDetail {
        
        private string 基金代碼Field;
        
        private string pK_會計年度Field;
        
        private string pK_動支編號Field;
        
        private string pK_種類Field;
        
        private string pK_次別Field;
        
        private string pK_明細號Field;
        
        private string f_科室代碼Field;
        
        private string f_用途別代碼Field;
        
        private string f_計畫代碼Field;
        
        private string f_動支金額Field;
        
        private string f_製票日Field;
        
        private string f_是否核定Field;
        
        private string f_核定金額Field;
        
        private string f_核定日期Field;
        
        private string f_摘要Field;
        
        private string f_受款人Field;
        
        private string f_受款人編號Field;
        
        private string f_原動支編號Field;
        
        private string f_批號Field;
        
        /// <remarks/>
        public string 基金代碼 {
            get {
                return this.基金代碼Field;
            }
            set {
                this.基金代碼Field = value;
            }
        }
        
        /// <remarks/>
        public string PK_會計年度 {
            get {
                return this.pK_會計年度Field;
            }
            set {
                this.pK_會計年度Field = value;
            }
        }
        
        /// <remarks/>
        public string PK_動支編號 {
            get {
                return this.pK_動支編號Field;
            }
            set {
                this.pK_動支編號Field = value;
            }
        }
        
        /// <remarks/>
        public string PK_種類 {
            get {
                return this.pK_種類Field;
            }
            set {
                this.pK_種類Field = value;
            }
        }
        
        /// <remarks/>
        public string PK_次別 {
            get {
                return this.pK_次別Field;
            }
            set {
                this.pK_次別Field = value;
            }
        }
        
        /// <remarks/>
        public string PK_明細號 {
            get {
                return this.pK_明細號Field;
            }
            set {
                this.pK_明細號Field = value;
            }
        }
        
        /// <remarks/>
        public string F_科室代碼 {
            get {
                return this.f_科室代碼Field;
            }
            set {
                this.f_科室代碼Field = value;
            }
        }
        
        /// <remarks/>
        public string F_用途別代碼 {
            get {
                return this.f_用途別代碼Field;
            }
            set {
                this.f_用途別代碼Field = value;
            }
        }
        
        /// <remarks/>
        public string F_計畫代碼 {
            get {
                return this.f_計畫代碼Field;
            }
            set {
                this.f_計畫代碼Field = value;
            }
        }
        
        /// <remarks/>
        public string F_動支金額 {
            get {
                return this.f_動支金額Field;
            }
            set {
                this.f_動支金額Field = value;
            }
        }
        
        /// <remarks/>
        public string F_製票日 {
            get {
                return this.f_製票日Field;
            }
            set {
                this.f_製票日Field = value;
            }
        }
        
        /// <remarks/>
        public string F_是否核定 {
            get {
                return this.f_是否核定Field;
            }
            set {
                this.f_是否核定Field = value;
            }
        }
        
        /// <remarks/>
        public string F_核定金額 {
            get {
                return this.f_核定金額Field;
            }
            set {
                this.f_核定金額Field = value;
            }
        }
        
        /// <remarks/>
        public string F_核定日期 {
            get {
                return this.f_核定日期Field;
            }
            set {
                this.f_核定日期Field = value;
            }
        }
        
        /// <remarks/>
        public string F_摘要 {
            get {
                return this.f_摘要Field;
            }
            set {
                this.f_摘要Field = value;
            }
        }
        
        /// <remarks/>
        public string F_受款人 {
            get {
                return this.f_受款人Field;
            }
            set {
                this.f_受款人Field = value;
            }
        }
        
        /// <remarks/>
        public string F_受款人編號 {
            get {
                return this.f_受款人編號Field;
            }
            set {
                this.f_受款人編號Field = value;
            }
        }
        
        /// <remarks/>
        public string F_原動支編號 {
            get {
                return this.f_原動支編號Field;
            }
            set {
                this.f_原動支編號Field = value;
            }
        }
        
        /// <remarks/>
        public string F_批號 {
            get {
                return this.f_批號Field;
            }
            set {
                this.f_批號Field = value;
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void getVw_GBCVisaDetailCompletedEventHandler(object sender, getVw_GBCVisaDetailCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getVw_GBCVisaDetailCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getVw_GBCVisaDetailCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void getVw_GBCVisaDetailXMLCompletedEventHandler(object sender, getVw_GBCVisaDetailXMLCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getVw_GBCVisaDetailXMLCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getVw_GBCVisaDetailXMLCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public Vw_GBCVisaDetail[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((Vw_GBCVisaDetail[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void getVw_GBCVisaDetailJSONCompletedEventHandler(object sender, getVw_GBCVisaDetailJSONCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class getVw_GBCVisaDetailJSONCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal getVw_GBCVisaDetailJSONCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void GetYearCompletedEventHandler(object sender, GetYearCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetYearCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetYearCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void GetAcmWordNumCompletedEventHandler(object sender, GetAcmWordNumCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAcmWordNumCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAcmWordNumCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void GetAccKindCompletedEventHandler(object sender, GetAccKindCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAccKindCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAccKindCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void GetAccCountCompletedEventHandler(object sender, GetAccCountCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAccCountCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAccCountCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void GetAccDetailCompletedEventHandler(object sender, GetAccDetailCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetAccDetailCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetAccDetailCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string[] Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string[])(this.results[0]));
            }
        }
    }
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    public delegate void GetByPrimaryKeyCompletedEventHandler(object sender, GetByPrimaryKeyCompletedEventArgs e);
    
    /// <remarks/>
    [System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.6.1055.0")]
    [System.Diagnostics.DebuggerStepThroughAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    public partial class GetByPrimaryKeyCompletedEventArgs : System.ComponentModel.AsyncCompletedEventArgs {
        
        private object[] results;
        
        internal GetByPrimaryKeyCompletedEventArgs(object[] results, System.Exception exception, bool cancelled, object userState) : 
                base(exception, cancelled, userState) {
            this.results = results;
        }
        
        /// <remarks/>
        public string Result {
            get {
                this.RaiseExceptionIfNecessary();
                return ((string)(this.results[0]));
            }
        }
    }
}

#pragma warning restore 1591