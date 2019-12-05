using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Globalization;

namespace Inventory
{
    public partial class Form1 : Form
    {
        object misvalue = System.Reflection.Missing.Value;
        string connectionString;
        SqlConnection cnn;

        public class INVENTORY
        {
            public string category;
            public List<ITEM> item = new List<ITEM>();
        }
        public class ITEM
        {
            public string itemcode;
            public int sales30;
            public List<Stock> stock = new List<Stock>();
        }
        public class Stock
        {
            public string wh;
            public int sa;
            public int di;
            public int n;
            public int size;
        }
        public List<INVENTORY> inventory = new List<INVENTORY>();

        public class NModel
        {
            public int min;
            public int max;
            public int size;
            public int n;
        }
        public List<NModel> nModels = new List<NModel>();

        public class OTB
        {
            public DateTime LogDate;
            public string Category;
            public string Brand;
            public float Ratio;
            public string CeilingExpire;
            public float Ceiling;
            public float Target;
            public float BudgetCost;
            public string SalesCost;
            public string RefundCost;
            public float POCost;
            public string Remark;
            public DateTime Date;
            public string PONo;
            public string TotalInventory;
            public string OnOrder;
        }
        public OTB otb = new OTB();
        List<OTB> newOTB = new List<OTB>();

        public Form1()
        {
            InitializeComponent();
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {

            try
            {
                connectionString = @"Data Source=localhost;Initial Catalog=Databases;database=Example;integrated security=SSPI";
                cnn = new SqlConnection(connectionString);
                cnn.Open();

                string sql;
                sql = "select * from SAAll order by [category], itemcode";
                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            INVENTORY inv = new INVENTORY();
                            ITEM item = new ITEM();
                            Stock stock = new Stock();
                            var a = Convert.ToDateTime(reader["date"].ToString());
                            DateTime b = Convert.ToDateTime(date1.Text);
                            if (a != b)
                            {
                                continue;
                            }

                            if (inventory.Count == 0 || inventory[inventory.Count - 1].category != reader["category"].ToString())
                            {
                                inv.category = reader["category"].ToString();
                                inventory.Add(inv);
                            }
                            item.itemcode = reader["itemcode"].ToString();
                            stock.wh = "London";
                            stock.size = 3;
                            stock.sa = Convert.ToInt32(reader["101 London"].ToString());
                            item.stock.Add(stock);

                            Stock stock2 = new Stock();
                            stock2.wh = "Brampton";
                            stock2.size = 4;
                            stock2.sa = Convert.ToInt32(reader["103 Brampton"].ToString());
                            item.stock.Add(stock2);

                            Stock stock3 = new Stock();
                            stock3.wh = "Waterloo";
                            stock3.size = 3;
                            stock3.sa = Convert.ToInt32(reader["104 Waterloo"].ToString());
                            item.stock.Add(stock3);

                            Stock stock4 = new Stock();
                            stock4.wh = "Richmond Hill";
                            stock4.size = 2;
                            stock4.sa = Convert.ToInt32(reader["105 Richmond Hill"].ToString());
                            item.stock.Add(stock4);

                            Stock stock5 = new Stock();
                            stock5.wh = "Barrie";
                            stock5.size = 2;
                            stock5.sa = Convert.ToInt32(reader["106 Barrie"].ToString());
                            item.stock.Add(stock5);

                            Stock stock6 = new Stock();
                            stock6.wh = "Toronto Down Town 284";
                            stock6.size = 3;
                            stock6.sa = Convert.ToInt32(reader["107 Toronto Down Town 284"].ToString());
                            item.stock.Add(stock6);

                            Stock stock7 = new Stock();
                            stock7.wh = "Toronto Down Town 366";
                            stock7.size = 3;
                            stock7.sa = Convert.ToInt32(reader["108 Toronto Down Town 366"].ToString());
                            item.stock.Add(stock7);

                            Stock stock8 = new Stock();
                            stock8.wh = "Vaughan";
                            stock8.size = 3;
                            stock8.sa = Convert.ToInt32(reader["110 Vaughan"].ToString());
                            item.stock.Add(stock8);

                            Stock stock9 = new Stock();
                            stock9.wh = "Newmarket";
                            stock9.size = 2;
                            stock9.sa = Convert.ToInt32(reader["111 Newmarket"].ToString());
                            item.stock.Add(stock9);

                            Stock stock10 = new Stock();
                            stock10.wh = "Mississauga";
                            stock10.size = 4;
                            stock10.sa = Convert.ToInt32(reader["112 Mississauga"].ToString());
                            item.stock.Add(stock10);

                            Stock stock11 = new Stock();
                            stock11.wh = "Ajax";
                            stock11.size = 2;
                            stock11.sa = Convert.ToInt32(reader["114 Ajax"].ToString());
                            item.stock.Add(stock11);

                            Stock stock12 = new Stock();
                            stock12.wh = "Kingston";
                            stock12.size = 2;
                            stock12.sa = Convert.ToInt32(reader["115 Kingston"].ToString());
                            item.stock.Add(stock12);

                            Stock stock13 = new Stock();
                            stock13.wh = "Ottawa Merivale";
                            stock13.size = 3;
                            stock13.sa = Convert.ToInt32(reader["117 Ottawa Merivale"].ToString());
                            item.stock.Add(stock13);

                            Stock stock14 = new Stock();
                            stock14.wh = "Ottawa Orleans";
                            stock14.size = 2;
                            stock14.sa = Convert.ToInt32(reader["118 Ottawa Orleans"].ToString());
                            item.stock.Add(stock14);

                            Stock stock15 = new Stock();
                            stock15.wh = "Toronto Kennedy";
                            stock15.size = 3;
                            stock15.sa = Convert.ToInt32(reader["119 Toronto Kennedy"].ToString());
                            item.stock.Add(stock15);

                            Stock stock16 = new Stock();
                            stock16.wh = "Hamilton";
                            stock16.size = 3;
                            stock16.sa = Convert.ToInt32(reader["120 Hamilton"].ToString());
                            item.stock.Add(stock16);

                            Stock stock17 = new Stock();
                            stock17.wh = "St Catharines";
                            stock17.size = 1;
                            stock17.sa = Convert.ToInt32(reader["121 St Catharines"].ToString());
                            item.stock.Add(stock17);

                            Stock stock18 = new Stock();
                            stock18.wh = "Head Office";
                            stock18.size = 5;
                            stock18.sa = Convert.ToInt32(reader["123 Head Office"].ToString());
                            item.stock.Add(stock18);

                            Stock stock19 = new Stock();
                            stock19.wh = "Markham Unionville";
                            stock19.size = 4;
                            stock19.sa = Convert.ToInt32(reader["124 Markham Unionville"].ToString());
                            item.stock.Add(stock19);

                            Stock stock20 = new Stock();
                            stock20.wh = "Whitby";
                            stock20.size = 2;
                            stock20.sa = Convert.ToInt32(reader["126 Whitby"].ToString());
                            item.stock.Add(stock20);

                            Stock stock21 = new Stock();
                            stock21.wh = "Laval";
                            stock21.size = 2;
                            stock21.sa = Convert.ToInt32(reader["127 Laval"].ToString());
                            item.stock.Add(stock21);

                            Stock stock22 = new Stock();
                            stock22.wh = "Kanata";
                            stock22.size = 2;
                            stock22.sa = Convert.ToInt32(reader["128 Kanata"].ToString());
                            item.stock.Add(stock22);

                            Stock stock23 = new Stock();
                            stock23.wh = "Burlington";
                            stock23.size = 2;
                            stock23.sa = Convert.ToInt32(reader["129 Burlington"].ToString());
                            item.stock.Add(stock23);

                            Stock stock24 = new Stock();
                            stock24.wh = "West Island";
                            stock24.size = 1;
                            stock24.sa = Convert.ToInt32(reader["130 West Island"].ToString());
                            item.stock.Add(stock24);

                            Stock stock25 = new Stock();
                            stock25.wh = "Toronto Mid Town";
                            stock25.size = 2;
                            stock25.sa = Convert.ToInt32(reader["132 Toronto Mid Town"].ToString());
                            item.stock.Add(stock25);

                            Stock stock26 = new Stock();
                            stock26.wh = "Oshawa";
                            stock26.size = 1;
                            stock26.sa = Convert.ToInt32(reader["134 Oshawa"].ToString());
                            item.stock.Add(stock26);

                            Stock stock27 = new Stock();
                            stock27.wh = "Milton";
                            stock27.size = 1;
                            stock27.sa = Convert.ToInt32(reader["135 Milton"].ToString());
                            item.stock.Add(stock27);

                            Stock stock28 = new Stock();
                            stock28.wh = "Etobicoke";
                            stock28.size = 3;
                            stock28.sa = Convert.ToInt32(reader["136 Etobicoke"].ToString());
                            item.stock.Add(stock28);

                            Stock stock29 = new Stock();
                            stock29.wh = "Greenfield Park";
                            stock29.size = 2;
                            stock29.sa = Convert.ToInt32(reader["137 Greenfield Park"].ToString());
                            item.stock.Add(stock29);

                            Stock stock30 = new Stock();
                            stock30.wh = "Ottawa Downtown";
                            stock30.size = 2;
                            stock30.sa = Convert.ToInt32(reader["138 Ottawa Downtown"].ToString());
                            item.stock.Add(stock30);

                            Stock stock31 = new Stock();
                            stock31.wh = "Montreal";
                            stock31.size = 1;
                            stock31.sa = Convert.ToInt32(reader["139 Montreal"].ToString());
                            item.stock.Add(stock31);

                            Stock stock32 = new Stock();
                            stock32.wh = "Vancouver";
                            stock32.size = 1;
                            stock32.sa = Convert.ToInt32(reader["140 Vancouver"].ToString());
                            item.stock.Add(stock32);

                            Stock stock33 = new Stock();
                            stock33.wh = "Richmond";
                            stock33.size = 2;
                            stock33.sa = Convert.ToInt32(reader["141 Richmond"].ToString());
                            item.stock.Add(stock33);

                            Stock stock34 = new Stock();
                            stock34.wh = "Coquitlam";
                            stock34.size = 1;
                            stock34.sa = Convert.ToInt32(reader["142 Coquitlam"].ToString());
                            item.stock.Add(stock34);

                            Stock stock35 = new Stock();
                            stock35.wh = "Burnaby";
                            stock35.size = 2;
                            stock35.sa = Convert.ToInt32(reader["143 Burnaby"].ToString());
                            item.stock.Add(stock35);

                            Stock stock36 = new Stock();
                            stock36.wh = "Gatineau";
                            stock36.size = 2;
                            stock36.sa = Convert.ToInt32(reader["144 Gatineau"].ToString());
                            item.stock.Add(stock36);

                            Stock stock37 = new Stock();
                            stock37.wh = "Grandview";
                            stock37.size = 2;
                            stock37.sa = Convert.ToInt32(reader["145 Grandview"].ToString());
                            item.stock.Add(stock37);
                            inventory[inventory.Count - 1].item.Add(item);
                        }
                    }
                }

                sql = "select * from DIAll order by [category], itemcode";
                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            INVENTORY inv = new INVENTORY();
                            ITEM item = new ITEM();
                            Stock stock = new Stock();

                            var a = Convert.ToDateTime(reader["date"].ToString());
                            DateTime b = Convert.ToDateTime(date1.Text);
                            if (a != b)
                            {
                                continue;
                            }
                            inv = inventory.Find(x => x.category == reader["category"].ToString());
                            item = inv.item.Find(x => x.itemcode == reader["itemcode"].ToString());
                            stock = item.stock.Find(x => x.wh == "London");
                            stock.di = Convert.ToInt32(reader["101 London"].ToString());

                            stock = item.stock.Find(x => x.wh == "Brampton");
                            stock.di = Convert.ToInt32(reader["103 Brampton"].ToString());

                            stock = item.stock.Find(x => x.wh == "Waterloo");
                            stock.di = Convert.ToInt32(reader["104 Waterloo"].ToString());

                            stock = item.stock.Find(x => x.wh == "Richmond Hill");
                            stock.di = Convert.ToInt32(reader["105 Richmond Hill"].ToString());

                            stock = item.stock.Find(x => x.wh == "Barrie");
                            stock.di = Convert.ToInt32(reader["106 Barrie"].ToString());

                            stock = item.stock.Find(x => x.wh == "Toronto Down Town 284");
                            stock.di = Convert.ToInt32(reader["107 Toronto Down Town 284"].ToString());

                            stock = item.stock.Find(x => x.wh == "Toronto Down Town 366");
                            stock.di = Convert.ToInt32(reader["108 Toronto Down Town 366"].ToString());

                            stock = item.stock.Find(x => x.wh == "Vaughan");
                            stock.di = Convert.ToInt32(reader["110 Vaughan"].ToString());

                            stock = item.stock.Find(x => x.wh == "Newmarket");
                            stock.di = Convert.ToInt32(reader["111 Newmarket"].ToString());

                            stock = item.stock.Find(x => x.wh == "Mississauga");
                            stock.di = Convert.ToInt32(reader["112 Mississauga"].ToString());

                            stock = item.stock.Find(x => x.wh == "Ajax");
                            stock.di = Convert.ToInt32(reader["114 Ajax"].ToString());

                            stock = item.stock.Find(x => x.wh == "Kingston");
                            stock.di = Convert.ToInt32(reader["115 Kingston"].ToString());

                            stock = item.stock.Find(x => x.wh == "Ottawa Merivale");
                            stock.di = Convert.ToInt32(reader["117 Ottawa Merivale"].ToString());

                            stock = item.stock.Find(x => x.wh == "Ottawa Orleans");
                            stock.di = Convert.ToInt32(reader["118 Ottawa Orleans"].ToString());

                            stock = item.stock.Find(x => x.wh == "Toronto Kennedy");
                            stock.di = Convert.ToInt32(reader["119 Toronto Kennedy"].ToString());

                            stock = item.stock.Find(x => x.wh == "Hamilton");
                            stock.di = Convert.ToInt32(reader["120 Hamilton"].ToString());

                            stock = item.stock.Find(x => x.wh == "St Catharines");
                            stock.di = Convert.ToInt32(reader["121 St Catharines"].ToString());

                            stock = item.stock.Find(x => x.wh == "Head Office");
                            stock.di = Convert.ToInt32(reader["123 Head Office"].ToString());

                            stock = item.stock.Find(x => x.wh == "Markham Unionville");
                            stock.di = Convert.ToInt32(reader["124 Markham Unionville"].ToString());

                            stock = item.stock.Find(x => x.wh == "Whitby");
                            stock.di = Convert.ToInt32(reader["126 Whitby"].ToString());

                            stock = item.stock.Find(x => x.wh == "Laval");
                            stock.di = Convert.ToInt32(reader["127 Laval"].ToString());

                            stock = item.stock.Find(x => x.wh == "Kanata");
                            stock.di = Convert.ToInt32(reader["128 Kanata"].ToString());

                            stock = item.stock.Find(x => x.wh == "Burlington");
                            stock.di = Convert.ToInt32(reader["129 Burlington"].ToString());

                            stock = item.stock.Find(x => x.wh == "West Island");
                            stock.di = Convert.ToInt32(reader["130 West Island"].ToString());

                            stock = item.stock.Find(x => x.wh == "Toronto Mid Town");
                            stock.di = Convert.ToInt32(reader["132 Toronto Mid Town"].ToString());

                            stock = item.stock.Find(x => x.wh == "Oshawa");
                            stock.di = Convert.ToInt32(reader["134 Oshawa"].ToString());

                            stock = item.stock.Find(x => x.wh == "Milton");
                            stock.di = Convert.ToInt32(reader["135 Milton"].ToString());

                            stock = item.stock.Find(x => x.wh == "Etobicoke");
                            stock.di = Convert.ToInt32(reader["136 Etobicoke"].ToString());

                            stock = item.stock.Find(x => x.wh == "Greenfield Park");
                            stock.di = Convert.ToInt32(reader["137 Greenfield Park"].ToString());

                            stock = item.stock.Find(x => x.wh == "Ottawa Downtown");
                            stock.di = Convert.ToInt32(reader["138 Ottawa Downtown"].ToString());

                            stock = item.stock.Find(x => x.wh == "Montreal");
                            stock.di = Convert.ToInt32(reader["139 Montreal"].ToString());

                            stock = item.stock.Find(x => x.wh == "Vancouver");
                            stock.di = Convert.ToInt32(reader["140 Vancouver"].ToString());

                            stock = item.stock.Find(x => x.wh == "Richmond");
                            stock.di = Convert.ToInt32(reader["141 Richmond"].ToString());

                            stock = item.stock.Find(x => x.wh == "Coquitlam");
                            stock.di = Convert.ToInt32(reader["142 Coquitlam"].ToString());

                            stock = item.stock.Find(x => x.wh == "Burnaby");
                            stock.di = Convert.ToInt32(reader["143 Burnaby"].ToString());

                            stock = item.stock.Find(x => x.wh == "Gatineau");
                            stock.di = Convert.ToInt32(reader["144 Gatineau"].ToString());

                            stock = item.stock.Find(x => x.wh == "Grandview");
                            stock.di = Convert.ToInt32(reader["145 Grandview"].ToString());

                        }
                    }
                }

                //sql = "select * from nmodel where itemsize=1";
                //using (SqlCommand scmd = new SqlCommand(sql, cnn))
                //{
                //    using (SqlDataReader reader = scmd.ExecuteReader())
                //    {
                //        while (reader.Read())
                //        {
                //            NModel tmpNModel = new NModel();
                //            tmpNModel.min = Convert.ToInt32(reader["runratemin"].ToString());
                //            tmpNModel.max = Convert.ToInt32(reader["runratemax"].ToString());
                //            tmpNModel.size = Convert.ToInt32(reader["warehousesize"].ToString());

                //            tmpNModel.n = Convert.ToInt32(reader["nqty"].ToString());
                //            nModels.Add(tmpNModel);
                //        }
                //    }
                //}

                //sql = "select t1.*, t2.topcategory from sales30 t1, saall t2 where t1.itemcode=t2.itemcode and t2.date='" + Convert.ToDateTime(date1.Text) + "'";
                //using (SqlCommand scmd = new SqlCommand(sql, cnn))
                //{
                //    using (SqlDataReader reader = scmd.ExecuteReader())
                //    {
                //        while (reader.Read())
                //        {
                //            INVENTORY inv = new INVENTORY();
                //            ITEM item = new ITEM();

                //            inv = inventory.Find(x => x.category == reader["topcategory"].ToString());
                //            item = inv.item.Find(x => x.itemcode == reader["itemcode"].ToString());
                //            if (item != null)
                //            {
                //                item.sales30 = Convert.ToInt32(reader["sku"].ToString());
                //                foreach(Stock tmpStock in item.stock)
                //                {
                //                    foreach(NModel tmpNModel in nModels)
                //                    {
                //                        if (item.sales30 > tmpNModel.min && item.sales30 <= tmpNModel.max)
                //                        {
                //                            if (tmpNModel.size == 6 && (tmpStock.wh == "Grandview" || tmpStock.wh == "Greenfield Park" || tmpStock.wh == "Ottawa Downtown") && tmpStock.n == 0)
                //                            {
                //                                tmpStock.n = tmpNModel.n;

                //                            }
                //                            else if (tmpStock.size == tmpNModel.size && tmpNModel.n > 0)
                //                            {
                //                                tmpStock.n = tmpNModel.n;

                //                            }
                //                        }
                //                    }
                //                }
                //            }
                //        }
                //    }
                //}
                //sql = "if OBJECT_ID('SADI') is not null Drop table SADI; create table SADI(Category nvarchar(255), ItemCode nvarchar(255), WH nvarchar(255), SA int, DI int, N int); ";
                //using (SqlCommand scmd = new SqlCommand(sql, cnn))
                //{
                //    using (SqlDataReader reader = scmd.ExecuteReader()) { }
                //}
                DateTime c = Convert.ToDateTime(date1.Text);
                foreach (INVENTORY tmpInv in inventory)
                {
                    foreach (ITEM tmpItem in tmpInv.item)
                    {
                        foreach (Stock tmpStock in tmpItem.stock)
                        {
                            sql = "INSERT INTO SADI (Category, ItemCode, WH, SA, DI, N, Date) VALUES ("
                                    + "'" + tmpInv.category + "',"
                                    + "'" + tmpItem.itemcode + "',"
                                    + "'" + tmpStock.wh + "',"
                                    + "'" + tmpStock.sa + "',"
                                    + "'" + tmpStock.di + "'," + "'" + tmpStock.n + "'," + "'" + c + "')";
                            using (SqlCommand scmd = new SqlCommand(sql, cnn))
                            {
                                using (SqlDataReader reader = scmd.ExecuteReader()) { }
                            }
                        }
                    }
                }
            }


            catch (SqlException ie)
            {
                MessageBox.Show(ie.ToString());
            }
            MessageBox.Show("Complete!");
        }

        private void btnAutoPOConvert_Click(object sender, EventArgs e)
        {
            try
            {
                connectionString = @"Data Source=localhost;Initial Catalog=Databases;database=Example;integrated security=SSPI";
                cnn = new SqlConnection(connectionString);
                cnn.Open();

                string sql;
                sql = "if OBJECT_ID('AutoConvertRate') is not null Drop table AutoConvertRate; ";
                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader()) { }
                }

                sql = "select Category, date, " +
                        "case when AutoSKU is null then 0 else autosku end AutoSKU,  " +
                        "case when ConvertedSKU is null then 0 else convertedsku end ConvertedSKU,  " +
                        "case when manualsku is null then 0 else manualsku end ManualSKU,  " +
                        "case when[ManualSKU(AutoSKU)] is null then 0 else [ManualSKU(AutoSKU)] end 'ManualSKU(AutoSKU)',  " +
                        "case when autosku is null then 0 else convert(float, convertedsku) / convert(float, autosku) end ConvertRatio1, " +
                           "case when autosku is null then 1 else (convert(float, convertedsku) + convert(float, manualsku)) / convert(float, autosku) end TotalRatio1, " +
                                 "case when AutoQty is null then 0 else AutoQty end AutoQty,  " +
                        "case when ConvertedQty is null then 0 else ConvertedQty end ConvertedQty,  " +
                        "case when ManualQty is null then 0 else ManualQty end ManualQty,  " +
                        "case when autoqty is null then 0 else convert(float, convertedqty) / convert(float, autoqty) end ConvertRatio2, " +
                           "case when autoqty is null then 1 else (convert(float, ConvertedQty) + convert(float, ManualQty)) / convert(float, autoqty) end TotalRatio2, " +
                               "case when AutoCost is null then 0 else AutoCost end AutoCost,  " +
                        "case when ConvertedCost is null then 0 else ConvertedCost end ConvertedCost,  " +
                        "case when ManualCost is null then 0 else ManualCost end ManualCost,  " +
                        "case when AutoCost is null then 0 else convert(float, convertedcost) / convert(float, autocost) end ConvertRatio3, " +
                           "case when autocost is null then 1 else (convert(float, convertedcost) + convert(float, manualcost)) / convert(float, autocost) end TotalRatio3 " +
                        "into AutoConvertRate " +
                        "from " +
                            "(select t3.category, t3.date, AutoSKU, " +

                            "case when convertedsku is null then 0 else convertedsku end ConvertedSKU, " +

                            "case when ManualSKU is null then 0 else ManualSKU end ManualSKU, " +

                            "case when t2.[ManualSKU(AutoSKU)] is null then 0 else t2.[ManualSKU(AutoSKU)] end 'ManualSKU(AutoSKU)', " +
                        "AutoQty, " +
                            "case when ConvertedQty is null then 0 else ConvertedQty end ConvertedQty,  " +
                            "case when ManualQty is null then 0 else ManualQty end ManualQty,  " +
                            "AutoCost,  " +
                            "case when ConvertedCost is null then 0 else ConvertedCost end ConvertedCost,  " +
                            "case when ManualCost is null then 0 else ManualCost end ManualCost " +
                            "from " +
                                "(select category, date from category, " +
                                    "(select date from autopo where date between '" + Convert.ToDateTime(dateTimeAutoPOStart.Text) + "' and '" + Convert.ToDateTime(dateTimeAutoPOEnd.Text) + "' " +
                                    "union " +
                                    "select podate date from po where podate between '" + Convert.ToDateTime(dateTimeAutoPOStart.Text) + "' and '" + Convert.ToDateTime(dateTimeAutoPOEnd.Text) + "' " + "group by podate) t1 " +
                                ")t3 " +
                    "full outer join " +
                    "(select t1.category, t1.date, AutoSKU, AutoQty, AutoCost, ConvertedSKU, ConvertedQty, " +
                                "case when ConvertedCost is null then 0 else ConvertedCost end ConvertedCost from " +
                                    "(select category, date, count(*) AutoSKU, sum(qty) AutoQty, sum(Cost) AutoCost from " +
                                        "(select t2.category, t1.date, t1.itemcode, sum(ordqty) qty, sum(ordqty* ConvertCAD) Cost from autopo t1, itemlist t2 " +
                                        "where t1.itemcode= t2.itemcode and t2.Phaseout= 0 " +
                                        "group by t2.category, date, t1.itemcode, ordqty, convertcad) t1 " +
                                    "group by category, date) t1 full outer join " +
                                    "(select category, date, count(*) ConvertedSKU, sum(qty) ConvertedQty, sum(cost) ConvertedCost from " +
                                        "(select t3.category, t1.itemcode, (orderqty-voidqty) qty, cost* currency*(orderqty-voidqty) Cost, t2.date " +
                                         "from po t1, itemlist t3, " +
                                            "(select itemcode, date, popno from autopo " +
                                            "group by itemcode, date, popno) t2 " +
                                        "where t1.itemcode=t2.itemcode and t1.popno = t2.popno and t1.itemcode=t3.itemcode and t3.phaseout= 0 " +
                                        ") t1 " +
                                    "group by category, date) t2 " +
                                "on t1.Category=t2.category and t1.Date=t2.Date) t1 on t3.date=t1.date and t3.category = t1.Category " +
                                "full outer join " +
                                "(select category, date, " +
                                "sum(case when mdate is not null then 1 else 0 end) 'ManualSKU(AutoSKU)', " +
                                "sum(manualqty) ManualQty, " +
                                "sum(manualcost) ManualCost, " +
                                "count(*) ManualSKU " +
                                "from " +
                                    "(select category, t1.date, t1.itemcode, t2.date mdate, sum(qty) ManualQty, sum(cost) ManualCost " +
                                    "from " +
                                        "(select t2.category, podate date, t1.itemcode, (orderqty-voidqty) qty, " +
                                        "cost* currency*(orderqty-voidqty) Cost from po t1, itemlist t2 " +
                                         "where t1.itemcode=t2.itemcode and type='manual' and podate> '" + Convert.ToDateTime(dateTimeAutoPOStart.Text) + "' and t1.itemcode in " +
                                             "(select itemcode from itemlist " +
                                             "where phaseout= 0 and categoryid in (select categoryid from category) " +
                                            "and itemcode not in  " +
                                                "(select itemcode from itemlist t1, ExcludedBrand t2 " +
                                                "where t1.categoryid=t2.categoryid and t1.BrandName=t2.brandname) " +
                                            ") " +
                                        "group by t2.category, podate, t1.itemcode, currency, cost, orderqty, voidqty) t1 left join " +
                                        "(select date, itemcode from autopo group by date, itemcode) t2 " +
                                    "on t1.date=t2.date and t1.itemcode=t2.itemcode " +
                                    "group by category, t1.date, t1.itemcode, t2.date)t1 " +
                                "group by category, date) t2 on t2.category = t3.Category and t2.date=t3.date " +
                            ") t1 " +
                        "where date between '" + Convert.ToDateTime(dateTimeAutoPOStart.Text) + "' and '" + Convert.ToDateTime(dateTimeAutoPOEnd.Text) + "' " +
                        "and (autosku !=0 or manualsku !=0) " +
                        "order by category, date";

                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader()) { }
                }

                sql = "if OBJECT_ID('AutoConvertRate2') is not null Drop table AutoConvertRate2; ";
                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader()) { }
                }

                sql = "select *, " +
                    "case when(convert(float, Excellent1) + convert(float, Excellent2) + convert(float, Excellent3) + convert(float, Good1) + convert(float, Good2) + convert(float, Good3) + convert(float, Bad1) + convert(float, Bad2) + convert(float, Bad3) + convert(float, NoAutoPO)) = 0 then 'No' " +
                         "when(convert(float, Excellent1) + convert(float, Excellent2) + convert(float, Excellent3)) / (convert(float, Excellent1) + convert(float, Excellent2) + convert(float, Excellent3) + convert(float, Good1) + convert(float, Good2) + convert(float, Good3) + convert(float, Bad1) + convert(float, Bad2) + convert(float, Bad3) + convert(float, NoAutoPO)) >= 0.5 then 'Yes' " +
                         "when(convert(float, Excellent1) + convert(float, Excellent2) + convert(float, Excellent3)) / (convert(float, Excellent1) + convert(float, Excellent2) + convert(float, Excellent3) + convert(float, Good1) + convert(float, Good2) + convert(float, Good3) + convert(float, Bad1) + convert(float, Bad2) + convert(float, Bad3) + convert(float, NoAutoPO)) < 0.5 and " +
                                                                    "(convert(float, Excellent1) + convert(float, Excellent2) + convert(float, Excellent3) + convert(float, Good1) + convert(float, Good2) + convert(float, Good3)) / (convert(float, Excellent1) + convert(float, Excellent2) + convert(float, Excellent3) + convert(float, Good1) + convert(float, Good2) + convert(float, Good3) + convert(float, Bad1) + convert(float, Bad2) + convert(float, Bad3) + convert(float, NoAutoPO)) >= 0.5 then 'Moderate' " +
                         "else 'No' end Suitable " +
                    "into AutoConvertRate2 from " +
                        "(select category, " +
                        "sum(case when autosku != 0 and convertratio1 > 0.05 and(totalratio1 - convertratio1) < 0.1 then 1 else 0 end) Excellent1, " +
                        "sum(case when autosku != 0 and convertratio1 > 0.05 and(totalratio1 - convertratio1) <= 0.5 and(totalratio1 - convertratio1) >= 0.1 then 1 else 0 end) Good1, " +
                        "sum(case when autosku != 0 and((convertratio1 > 0.05 and(totalratio1 - convertratio1) > 0.5) or " +
                           "(convertratio1 <= 0.05 and(totalratio1 - convertratio1) > 0.5)) then 1 else 0 end) Bad1, " +
                        "sum(case when convertratio1 <= 0.05 and(totalratio1 - convertratio1) <= 0.5 then 1 else 0 end) Invaid1, " +
                        "sum(case when autosku != 0 and convertratio2 > 0.05 and(totalratio2 - convertratio2) < 0.1 then 1 else 0 end) Excellent2, " +
                        "sum(case when autosku != 0 and convertratio2 > 0.05 and(totalratio2 - convertratio2) <= 0.5 and(totalratio2 - convertratio2) >= 0.1 then 1 else 0 end) Good2, " +
                        "sum(case when autosku != 0 and((convertratio2 > 0.05 and(totalratio2 - convertratio2) > 0.5) or " +
                           "(convertratio2 <= 0.05 and(totalratio2 - convertratio2) > 0.5)) then 1 else 0 end) Bad2, " +
                        "sum(case when convertratio2 <= 0.05 and(totalratio2 - convertratio2) <= 0.5 then 1 else 0 end) Invaid2, " +
                        "sum(case when autosku != 0 and convertratio3 > 0.05 and(totalratio3 - convertratio3) < 0.1 then 1 else 0 end) Excellent3, " +
                        "sum(case when autosku != 0 and convertratio3 > 0.05 and(totalratio3 - convertratio3) <= 0.5 and(totalratio3 - convertratio3) >= 0.1 then 1 else 0 end) Good3, " +
                        "sum(case when autosku != 0 and((convertratio3 > 0.05 and(totalratio3 - convertratio3) > 0.5) or " +
                           "(convertratio3 <= 0.05 and(totalratio3 - convertratio3) > 0.5)) then 1 else 0 end) Bad3, " +
                        "sum(case when convertratio3 <= 0.05 and(totalratio3 - convertratio3) <= 0.5 then 1 else 0 end) Invaid3, " +
                        "sum(case when autosku = 0 then 1 else 0 end) NoAutoPO " +
                          "from autoconvertrate " +
                          "group by category)t1 " +
                        "order by category";

                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader()) { }
                }

            }


            catch (SqlException ie)
            {
                MessageBox.Show(ie.ToString());
            }
            MessageBox.Show("Complete!");
        }

        private void BtnOTB_Click(object sender, EventArgs e)
        {
            try
            {
                connectionString = @"Data Source=localhost;Initial Catalog=Databases;database=Example;integrated security=SSPI";
                cnn = new SqlConnection(connectionString);
                cnn.Open();

                string sql;
                DateTime lastday = new DateTime();
                sql = "select top 1 date from otb group by date order by date desc";
                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            lastday = Convert.ToDateTime(reader["Date"].ToString());
                        }
                    }
                }

                sql = "select * from (select * from (select *, ROW_NUMBER() OVER(PARTITION BY category, date order by Category, logdate desc, budgetcost) AS row from otb) t1 "
                        + "where row = 1" 
                        + " union "
                        + "select * from (select *, ROW_NUMBER() OVER(PARTITION BY category, date order by Category, logdate desc, budgetcost) AS row from otb" 
                        + " where convert(varchar,convert(datetime, logdate, 111),8)<'04:00:00') t1 "
                        + " where row = 1)t1  order by category,logdate ";

                using (SqlCommand scmd = new SqlCommand(sql, cnn))
                {
                    using (SqlDataReader reader = scmd.ExecuteReader())
                    {
                        
                        while (reader.Read())
                        {
                            OTB tmpOTB = new OTB();

                            if (otb.Category == null)
                            {
                                otb.LogDate = Convert.ToDateTime(reader["LogDate"].ToString());
                                otb.Category = reader["Category"].ToString();
                                otb.Brand = reader["Brand"].ToString();
                                otb.CeilingExpire = reader["CeilingExpire"].ToString();
                                otb.SalesCost = reader["SalesCost"].ToString();
                                otb.RefundCost = reader["RefundCost"].ToString();
                                otb.Remark = reader["Remark"].ToString();
                                otb.PONo = reader["PONo"].ToString();
                                otb.Date = Convert.ToDateTime(reader["Date"].ToString());
                                otb.BudgetCost = float.Parse(reader["BudgetCost"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                otb.Ceiling = float.Parse(reader["Ceiling"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                otb.Target = float.Parse(reader["Target"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                otb.Ratio = float.Parse(reader["Ratio"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                otb.POCost = float.Parse(reader["POCost"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                otb.TotalInventory = reader["TotalInventory"].ToString();
                                otb.OnOrder = reader["OnOrder"].ToString();

                            }
                            else
                            {
                                tmpOTB.LogDate = Convert.ToDateTime(reader["LogDate"].ToString());
                                tmpOTB.Category = reader["Category"].ToString();
                                tmpOTB.Brand = reader["Brand"].ToString();
                                tmpOTB.CeilingExpire = reader["CeilingExpire"].ToString();
                                tmpOTB.SalesCost = reader["SalesCost"].ToString();
                                tmpOTB.RefundCost = reader["RefundCost"].ToString();
                                tmpOTB.Remark = reader["Remark"].ToString();
                                tmpOTB.PONo = reader["PONo"].ToString();
                                tmpOTB.Date = Convert.ToDateTime(reader["Date"].ToString());
                                tmpOTB.BudgetCost = float.Parse(reader["BudgetCost"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                tmpOTB.Ceiling = float.Parse(reader["Ceiling"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                tmpOTB.Target = float.Parse(reader["Target"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                tmpOTB.Ratio = float.Parse(reader["Ratio"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                tmpOTB.POCost = float.Parse(reader["POCost"].ToString(), CultureInfo.InvariantCulture.NumberFormat);
                                tmpOTB.TotalInventory = reader["TotalInventory"].ToString();
                                tmpOTB.OnOrder = reader["OnOrder"].ToString();

                                if (tmpOTB.Category != otb.Category)
                                {
                                    while(otb.Date < lastday)
                                    {
                                        OTB newotb = new OTB();
                                        otb.LogDate = otb.LogDate.AddDays(1);
                                        otb.Date = otb.Date.AddDays(1);
                                        newotb.Brand = otb.Brand;
                                        newotb.BudgetCost = otb.BudgetCost;
                                        newotb.Category = otb.Category;
                                        newotb.Ceiling = otb.Ceiling;
                                        newotb.CeilingExpire = otb.CeilingExpire;
                                        newotb.Date = otb.Date;
                                        newotb.LogDate = otb.LogDate;
                                        newotb.POCost = otb.POCost;
                                        newotb.PONo = otb.PONo;
                                        newotb.Ratio = otb.Ratio;
                                        newotb.RefundCost = otb.RefundCost;
                                        newotb.Remark = otb.Remark;
                                        newotb.SalesCost = otb.SalesCost;
                                        newotb.Target = otb.Target;
                                        newotb.TotalInventory = otb.TotalInventory;
                                        newotb.OnOrder = otb.OnOrder;
                                        newOTB.Add(newotb);
                                    }
                                    otb = tmpOTB;
                                }
                                else
                                {
                                    //if (tmpOTB.Category== "Office Equipment" && tmpOTB.Date==Convert.ToDateTime("2019/11/13"))
                                    //{
                                    //    int i = 0;
                                    //}
                                    while ((otb.Date.AddDays(1) < tmpOTB.Date) || (otb.Date.AddDays(1) == tmpOTB.Date && tmpOTB.LogDate.Hour > 4))
                                    {
                                        OTB newotb = new OTB();
                                        otb.LogDate = otb.LogDate.AddDays(1);
                                        otb.Date = otb.Date.AddDays(1);
                                        newotb.Brand = otb.Brand;
                                        newotb.BudgetCost = otb.BudgetCost;
                                        newotb.Category = otb.Category;
                                        newotb.Ceiling = otb.Ceiling;
                                        newotb.CeilingExpire = otb.CeilingExpire;
                                        newotb.Date = otb.Date;
                                        newotb.LogDate = otb.LogDate;
                                        newotb.POCost = otb.POCost;
                                        newotb.PONo = otb.PONo;
                                        newotb.Ratio = otb.Ratio;
                                        newotb.RefundCost = otb.RefundCost;
                                        newotb.Remark = otb.Remark;
                                        newotb.SalesCost = otb.SalesCost;
                                        newotb.Target = otb.Target;
                                        newotb.TotalInventory = otb.TotalInventory;
                                        newotb.OnOrder = otb.OnOrder;
                                        newOTB.Add(newotb);
                                    }
                                    otb = tmpOTB;
                                }
                            }
                        }
                    }
                }
                if (newOTB.Count > 0)
                {
                    foreach (OTB otb in newOTB)
                    {
                        string logdate = otb.LogDate.ToString("yyyy/MM/dd HH:mm:ss");
                        logdate = logdate.Remove(11);
                        logdate = logdate.Insert(11, "01:00:00");
                        sql = $"insert into otb (LogDate,Category,Brand,Ratio,Ceiling,CeilingExpire,Target,BudgetCost,SalesCost,RefundCost,POCost,Remark,Date,PONo,TotalInventory,OnOrder) "
                            + $"values('{logdate}','{otb.Category}','{otb.Brand}',{otb.Ratio},{otb.Ceiling},'{otb.CeilingExpire}',{otb.Target},"
                            + $"{otb.BudgetCost},'{otb.SalesCost}','{otb.RefundCost}',{otb.POCost},'Auto Insert','{otb.Date}','{otb.PONo}','{otb.TotalInventory}','{otb.OnOrder}')";
                        using (SqlCommand scmd = new SqlCommand(sql, cnn))
                        {
                            using (SqlDataReader reader = scmd.ExecuteReader())
                            {
                            }
                        }
                    }
                }
            }
            catch (SqlException ie)
            {
                MessageBox.Show(ie.ToString());
            }
        }
    }
}
