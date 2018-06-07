using System;
using System.Windows.Forms;
using System.Net;
using System.Text;
using System.Threading;
using System.Security.Permissions;
using System.Collections.Generic;
using System.Collections;
using HtmlAgilityPack;
using System.Data.OleDb;
using System.Data;

namespace pc
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]//　　注意： 类定义前需要加上下面两行，否则调用失败！
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            webBr.Url = new Uri("https://www.tianyancha.com/search?key=京东");
            //getExcel("E:\\公司名称.xls");


            //WebClient webClient = new WebClient();
            //String url = txtInput.Text.Trim();
            //if (url == null || url == "") return;
            //byte[] result = webClient.DownloadData(url);
            //txtShow.Text = Encoding.UTF8.GetString(result);


        }
        /// <summary>
        /// 从excel获得公司名称
        /// </summary>
        /// <param name="Path">excel路径</param>
        /// <returns></returns>
        public DataSet getExcel(String Path) {
                string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + Path + ";" + "Extended Properties=Excel 8.0;";
                OleDbConnection conn = new OleDbConnection(strConn);
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                DataSet ds = null;
                strExcel = "select * from [sheet1$]";
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                ds = new DataSet();
                myCommand.Fill(ds, "table1");
                return ds;
        }
        private void webBr_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            System.Windows.Forms.HtmlDocument document = webBr.Document;
            //htmlAgilityPack扩展库
            HtmlAgilityPack.HtmlDocument agHtmlDocument = new HtmlAgilityPack.HtmlDocument();
            agHtmlDocument.LoadHtml(webBr.DocumentText);
            HtmlNodeCollection htmlNodeName = agHtmlDocument.DocumentNode.SelectNodes("//div[@data-id]//a[@title]");
            HtmlNodeCollection htmlNodePhone = agHtmlDocument.DocumentNode.SelectNodes("//div[@data-id]/div[last()]/div[last()-1]/div[last()-1]//span[@class='sec-c3']");///div//span[@class='sec-c3']//following-sibling::span[1]
            List<String> dataId = new List<string>();
            foreach (HtmlNode item in htmlNodePhone)
            {
                String a = item.InnerHtml;

            }


            //if (webBr.Document.GetElementById("home-main-search") == null)
            //{
            //    return;
            //}
            //webBr.Document.GetElementById("home-main-search").InnerText = "百度";
            //Thread.Sleep(2000);
            ////调用js函数
            //webBr.Document.InvokeScript("header.search(true,'#home-main-search') ");

        }
    }
}
