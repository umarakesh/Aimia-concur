using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace BOS_PO_FROM_CSV_ConCur
{
    static class Global
    {

        static public string BPERR = "";
        static public string ERR = "";
        static public DateTime globaltime1;
        static public string globaltime;
        static public DataTable dt;

        public static string getpodocnum(SAPbobsCOM.Company oCompany)
        {
            string itemcode = "";
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            ors.DoQuery("Select DocNum From OPOR Where DocEntry=(Select MAX(DocEntry) From OPOR)");
            itemcode = ors.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return itemcode;
        }

        public static string getItemCode(SAPbobsCOM.Company oCompany, string Col63)
        {
            string itemcode = "";
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            ors.DoQuery("Select Left([@BOS_CCACT].U_Account,8) From [@BOS_CCACT] Where [@BOS_CCACT].U_Category = '"+ Col63 + "'");
            itemcode = ors.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return itemcode; 
        }

          public static int getPOCount(SAPbobsCOM.Company oCompany,string ponumber)
        {
            int pocount = 0;
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            ors.DoQuery("Select Count(U_concurpo) from OPOR Where U_concurpo='"+ ponumber +"'");
            pocount = Int32.Parse(ors.Fields.Item(0).Value.ToString());
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return pocount;
        }

        public static string getAccountCode(SAPbobsCOM.Company oCompany, string Col63)
        {
            string itemcode = "";
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            ors.DoQuery("Select [@BOS_CCACT].U_Account From [@BOS_CCACT] Where [@BOS_CCACT].U_Category = '" + Col63 + "'");
            itemcode = ors.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return itemcode.Replace("ZZ","");
        }

        public static string getDepartment(SAPbobsCOM.Company oCompany, string Col6, string Col7)
        {
            string retvalue = "";
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string empname = Col7 + " " + Col6;
            ors.DoQuery("select [@BOS_CCDIM].U_Department From [@BOS_CCDIM] Where [@BOS_CCDIM].U_Employee ='"+ empname +"'");
            retvalue = ors.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return retvalue;
        }

        public static string getPartner(SAPbobsCOM.Company oCompany, string Col6, string Col7)
        {
            string retvalue = "";
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string empname = Col7 + " " + Col6;
            ors.DoQuery("select [@BOS_CCDIM].U_Partner From [@BOS_CCDIM] Where [@BOS_CCDIM].U_Employee ='" + empname + "'");
            retvalue = ors.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return retvalue;
        }

        public static string getMngUnit(SAPbobsCOM.Company oCompany, string Col6, string Col7)
        {
            string retvalue = "";
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string empname = Col7 + " " + Col6;
            ors.DoQuery("select [@BOS_CCDIM].U_MngUnit From [@BOS_CCDIM] Where [@BOS_CCDIM].U_Employee ='" + empname + "'");
            retvalue = ors.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return retvalue;
        }


        public static string getProduct(SAPbobsCOM.Company oCompany, string Col6, string Col7)
        {
            string retvalue = "";
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string empname = Col7 + " " + Col6;
            ors.DoQuery("select[@BOS_CCDIM].U_Product From [@BOS_CCDIM] Where [@BOS_CCDIM].U_Employee ='" + empname + "'");
            retvalue = ors.Fields.Item(0).Value.ToString();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ors);
            ors = null;
            GC.Collect();
            return retvalue;
        }

     
        public static DataTable GetDistinctRecords(DataTable dt, string[] Columns)
        {
            DataTable dtUniqRecords = new DataTable();
            dtUniqRecords = dt.DefaultView.ToTable(true, Columns);
            return dtUniqRecords;
        }

        static public void datatablecreate()
        {
            dt = new DataTable();
            dt.Columns.Add("Col1", typeof(string));
            dt.Columns.Add("Col2", typeof(string));
            dt.Columns.Add("Col3", typeof(string));
            dt.Columns.Add("Col4", typeof(string));
            dt.Columns.Add("Col5", typeof(string));
            dt.Columns.Add("Col6", typeof(string));
            dt.Columns.Add("Col7", typeof(string));
            dt.Columns.Add("Col8", typeof(string));
            dt.Columns.Add("Col9", typeof(string));
            dt.Columns.Add("Col10", typeof(string));
            dt.Columns.Add("Col11", typeof(string));
            dt.Columns.Add("Col12", typeof(string));
            dt.Columns.Add("Col13", typeof(string));
            dt.Columns.Add("Col14", typeof(string));
            dt.Columns.Add("Col15", typeof(string));
            dt.Columns.Add("Col16", typeof(string));
            dt.Columns.Add("Col17", typeof(string));
            dt.Columns.Add("Col18", typeof(string));
            dt.Columns.Add("Col19", typeof(string));
            dt.Columns.Add("Col20", typeof(string));
            dt.Columns.Add("Col21", typeof(string));
            dt.Columns.Add("Col22", typeof(string));
            dt.Columns.Add("Col23", typeof(string));
            dt.Columns.Add("Col24", typeof(string));
            dt.Columns.Add("Col25", typeof(string));
            dt.Columns.Add("Col26", typeof(string));
            dt.Columns.Add("Col27", typeof(string));
            dt.Columns.Add("Col28", typeof(string));
            dt.Columns.Add("Col29", typeof(string));
            dt.Columns.Add("Col30", typeof(string));
            dt.Columns.Add("Col31", typeof(string));
            dt.Columns.Add("Col32", typeof(string));
            dt.Columns.Add("Col33", typeof(string));
            dt.Columns.Add("Col34", typeof(string));
            dt.Columns.Add("Col35", typeof(string));
            dt.Columns.Add("Col36", typeof(string));
            dt.Columns.Add("Col37", typeof(string));
            dt.Columns.Add("Col38", typeof(string));
            dt.Columns.Add("Col39", typeof(string));
            dt.Columns.Add("Col40", typeof(string));
            dt.Columns.Add("Col41", typeof(string));
            dt.Columns.Add("Col42", typeof(string));
            dt.Columns.Add("Col43", typeof(string));
            dt.Columns.Add("Col44", typeof(string));
            dt.Columns.Add("Col45", typeof(string));
            dt.Columns.Add("Col46", typeof(string));
            dt.Columns.Add("Col47", typeof(string));
            dt.Columns.Add("Col48", typeof(string));
            dt.Columns.Add("Col49", typeof(string));
            dt.Columns.Add("Col50", typeof(string));
            dt.Columns.Add("Col51", typeof(string));
            dt.Columns.Add("Col52", typeof(string));
            dt.Columns.Add("Col53", typeof(string));
            dt.Columns.Add("Col54", typeof(string));
            dt.Columns.Add("Col55", typeof(string));
            dt.Columns.Add("Col56", typeof(string));
            dt.Columns.Add("Col57", typeof(string));
            dt.Columns.Add("Col58", typeof(string));
            dt.Columns.Add("Col59", typeof(string));
            dt.Columns.Add("Col60", typeof(string));
            dt.Columns.Add("Col61", typeof(string));
            dt.Columns.Add("Col62", typeof(string));
            dt.Columns.Add("Col63", typeof(string));
            dt.Columns.Add("Col64", typeof(string));
            dt.Columns.Add("Col65", typeof(string));
            dt.Columns.Add("Col66", typeof(string));
            dt.Columns.Add("Col67", typeof(string));
            dt.Columns.Add("Col68", typeof(string));
            dt.Columns.Add("Col69", typeof(string));
            dt.Columns.Add("Col70", typeof(string));
            dt.Columns.Add("Col71", typeof(string));
            dt.Columns.Add("Col72", typeof(string));
            dt.Columns.Add("Col73", typeof(string));
            dt.Columns.Add("Col74", typeof(string));
            dt.Columns.Add("Col75", typeof(string));
            dt.Columns.Add("Col76", typeof(string));
            dt.Columns.Add("Col77", typeof(string));
            dt.Columns.Add("Col78", typeof(string));
            dt.Columns.Add("Col79", typeof(string));
            dt.Columns.Add("Col80", typeof(string));
            dt.Columns.Add("Col81", typeof(string));
            dt.Columns.Add("Col82", typeof(string));
            dt.Columns.Add("Col83", typeof(string));
            dt.Columns.Add("Col84", typeof(string));
            dt.Columns.Add("Col85", typeof(string));
            dt.Columns.Add("Col86", typeof(string));
            dt.Columns.Add("Col87", typeof(string));
            dt.Columns.Add("Col88", typeof(string));
            dt.Columns.Add("Col89", typeof(string));
            dt.Columns.Add("Col90", typeof(string));
            dt.Columns.Add("Col91", typeof(string));
            dt.Columns.Add("Col92", typeof(string));
            dt.Columns.Add("Col93", typeof(string));
            dt.Columns.Add("Col94", typeof(string));
            dt.Columns.Add("Col95", typeof(string));
            dt.Columns.Add("Col96", typeof(string));
            dt.Columns.Add("Col97", typeof(string));
            dt.Columns.Add("Col98", typeof(string));
            dt.Columns.Add("Col99", typeof(string));
            dt.Columns.Add("Col100", typeof(string));
            dt.Columns.Add("Col101", typeof(string));
            dt.Columns.Add("Col102", typeof(string));
            dt.Columns.Add("Col103", typeof(string));
            dt.Columns.Add("Col104", typeof(string));
            dt.Columns.Add("Col105", typeof(string));
            dt.Columns.Add("Col106", typeof(string));
            dt.Columns.Add("Col107", typeof(string));
            dt.Columns.Add("Col108", typeof(string));
            dt.Columns.Add("Col109", typeof(string));
            dt.Columns.Add("Col110", typeof(string));
            dt.Columns.Add("Col111", typeof(string));
            dt.Columns.Add("Col112", typeof(string));
            dt.Columns.Add("Col113", typeof(string));
            dt.Columns.Add("Col114", typeof(string));
            dt.Columns.Add("Col115", typeof(string));
            dt.Columns.Add("Col116", typeof(string));
            dt.Columns.Add("Col117", typeof(string));
            dt.Columns.Add("Col118", typeof(string));
            dt.Columns.Add("Col119", typeof(string));
            dt.Columns.Add("Col120", typeof(string));
            dt.Columns.Add("Col121", typeof(string));
            dt.Columns.Add("Col122", typeof(string));
            dt.Columns.Add("Col123", typeof(string));
            dt.Columns.Add("Col124", typeof(string));
            dt.Columns.Add("Col125", typeof(string));
            dt.Columns.Add("Col126", typeof(string));
            dt.Columns.Add("Col127", typeof(string));
            dt.Columns.Add("Col128", typeof(string));
            dt.Columns.Add("Col129", typeof(string));
            dt.Columns.Add("Col130", typeof(string));
            dt.Columns.Add("Col131", typeof(string));
            dt.Columns.Add("Col132", typeof(string));
            dt.Columns.Add("Col133", typeof(string));
            dt.Columns.Add("Col134", typeof(string));
            dt.Columns.Add("Col135", typeof(string));
            dt.Columns.Add("Col136", typeof(string));
            dt.Columns.Add("Col137", typeof(string));
            dt.Columns.Add("Col138", typeof(string));
            dt.Columns.Add("Col139", typeof(string));
            dt.Columns.Add("Col140", typeof(string));
            dt.Columns.Add("Col141", typeof(string));
            dt.Columns.Add("Col142", typeof(string));
            dt.Columns.Add("Col143", typeof(string));
            dt.Columns.Add("Col144", typeof(string));
            dt.Columns.Add("Col145", typeof(string));
            dt.Columns.Add("Col146", typeof(string));
            dt.Columns.Add("Col147", typeof(string));
            dt.Columns.Add("Col148", typeof(string));
            dt.Columns.Add("Col149", typeof(string));
            dt.Columns.Add("Col150", typeof(string));
            dt.Columns.Add("Col151", typeof(string));
            dt.Columns.Add("Col152", typeof(string));
            dt.Columns.Add("Col153", typeof(string));
            dt.Columns.Add("Col154", typeof(string));
            dt.Columns.Add("Col155", typeof(string));
            dt.Columns.Add("Col156", typeof(string));
            dt.Columns.Add("Col157", typeof(string));
            dt.Columns.Add("Col158", typeof(string));
            dt.Columns.Add("Col159", typeof(string));
            dt.Columns.Add("Col160", typeof(string));
            dt.Columns.Add("Col161", typeof(string));
            dt.Columns.Add("Col162", typeof(string));
            dt.Columns.Add("Col163", typeof(string));
            dt.Columns.Add("Col164", typeof(string));
            dt.Columns.Add("Col165", typeof(string));
            dt.Columns.Add("Col166", typeof(string));
            dt.Columns.Add("Col167", typeof(string));
            dt.Columns.Add("Col168", typeof(string));
            dt.Columns.Add("Col169", typeof(string));
            dt.Columns.Add("Col170", typeof(string));
            dt.Columns.Add("Col171", typeof(string));
            dt.Columns.Add("Col172", typeof(string));
            dt.Columns.Add("Col173", typeof(string));
            dt.Columns.Add("Col174", typeof(string));
            dt.Columns.Add("Col175", typeof(string));
            dt.Columns.Add("Col176", typeof(string));
            dt.Columns.Add("Col177", typeof(string));
            dt.Columns.Add("Col178", typeof(string));
            dt.Columns.Add("Col179", typeof(string));
            dt.Columns.Add("Col180", typeof(string));
            dt.Columns.Add("Col181", typeof(string));
            dt.Columns.Add("Col182", typeof(string));
            dt.Columns.Add("Col183", typeof(string));
            dt.Columns.Add("Col184", typeof(string));
            dt.Columns.Add("Col185", typeof(string));
            dt.Columns.Add("Col186", typeof(string));
            dt.Columns.Add("Col187", typeof(string));
            dt.Columns.Add("Col188", typeof(string));
            dt.Columns.Add("Col189", typeof(string));
            dt.Columns.Add("Col190", typeof(string));
            dt.Columns.Add("Col191", typeof(string));
            dt.Columns.Add("Col192", typeof(string));
            dt.Columns.Add("Col193", typeof(string));
            dt.Columns.Add("Col194", typeof(string));
            dt.Columns.Add("Col195", typeof(string));
            dt.Columns.Add("Col196", typeof(string));
            dt.Columns.Add("Col197", typeof(string));
            dt.Columns.Add("Col198", typeof(string));
            dt.Columns.Add("Col199", typeof(string));
            dt.Columns.Add("Col200", typeof(string));
            dt.Columns.Add("Col201", typeof(string));
            dt.Columns.Add("Col202", typeof(string));
            dt.Columns.Add("Col203", typeof(string));
            dt.Columns.Add("Col204", typeof(string));
            dt.Columns.Add("Col205", typeof(string));
            dt.Columns.Add("Col206", typeof(string));
            dt.Columns.Add("Col207", typeof(string));
            dt.Columns.Add("Col208", typeof(string));
            dt.Columns.Add("Col209", typeof(string));
            dt.Columns.Add("Col210", typeof(string));
            dt.Columns.Add("Col211", typeof(string));
            dt.Columns.Add("Col212", typeof(string));
            dt.Columns.Add("Col213", typeof(string));
            dt.Columns.Add("Col214", typeof(string));
            dt.Columns.Add("Col215", typeof(string));
            dt.Columns.Add("Col216", typeof(string));
            dt.Columns.Add("Col217", typeof(string));
            dt.Columns.Add("Col218", typeof(string));
            dt.Columns.Add("Col219", typeof(string));
            dt.Columns.Add("Col220", typeof(string));
            dt.Columns.Add("Col221", typeof(string));
            dt.Columns.Add("Col222", typeof(string));
            dt.Columns.Add("Col223", typeof(string));
            dt.Columns.Add("Col224", typeof(string));
            dt.Columns.Add("Col225", typeof(string));
            dt.Columns.Add("Col226", typeof(string));
            dt.Columns.Add("Col227", typeof(string));
            dt.Columns.Add("Col228", typeof(string));
            dt.Columns.Add("Col229", typeof(string));
            dt.Columns.Add("Col230", typeof(string));
            dt.Columns.Add("Col231", typeof(string));
            dt.Columns.Add("Col232", typeof(string));
            dt.Columns.Add("Col233", typeof(string));
            dt.Columns.Add("Col234", typeof(string));
            dt.Columns.Add("Col235", typeof(string));
            dt.Columns.Add("Col236", typeof(string));
            dt.Columns.Add("Col237", typeof(string));
            dt.Columns.Add("Col238", typeof(string));
            dt.Columns.Add("Col239", typeof(string));
            dt.Columns.Add("Col240", typeof(string));
            dt.Columns.Add("Col241", typeof(string));
            dt.Columns.Add("Col242", typeof(string));
            dt.Columns.Add("Col243", typeof(string));
            dt.Columns.Add("Col244", typeof(string));
            dt.Columns.Add("Col245", typeof(string));
            dt.Columns.Add("Col246", typeof(string));
            dt.Columns.Add("Col247", typeof(string));
            dt.Columns.Add("Col248", typeof(string));
            dt.Columns.Add("Col249", typeof(string));
            dt.Columns.Add("Col250", typeof(string));
            dt.Columns.Add("Col251", typeof(string));
            dt.Columns.Add("Col252", typeof(string));
            dt.Columns.Add("Col253", typeof(string));
            dt.Columns.Add("Col254", typeof(string));
            dt.Columns.Add("Col255", typeof(string));
            dt.Columns.Add("Col256", typeof(string));

        }

    }

    
}
