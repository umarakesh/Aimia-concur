using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Collections;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Globalization;
using System.Data;

namespace BOS_PO_FROM_CSV_ConCur
{
    internal class pofromcsv_cls
    {
        
        private SAPbobsCOM.Company oCompany;
        private SAPbouiCOM.Application SBO_Application;

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;

            SboGuiApi = new SAPbouiCOM.SboGuiApi();

            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication(-1);

        }

        private int SetConnectionContext()
        {
            int setConnectionContextReturn = 0;

            string sCookie = null;
            string sConnectionContext = null;
            int lRetCode = 0;

            oCompany = new SAPbobsCOM.Company();

            sCookie = oCompany.GetContextCookie();
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);

            if (oCompany.Connected == true)
            {
                oCompany.Disconnect();
            }
            setConnectionContextReturn = oCompany.SetSboLoginContext(sConnectionContext);

            return setConnectionContextReturn;
        }
        private int ConnectToCompany()
        {
            int connectToCompanyReturn = 0;
            connectToCompanyReturn = oCompany.Connect();
            return connectToCompanyReturn;
        }

        public pofromcsv_cls()
            : base()
        {

            SetApplication();
            if (!(SetConnectionContext() == 0))
            {
                SBO_Application.StatusBar.SetText("Failed setting a connection to DI API", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                System.Environment.Exit(0); //  Terminating the Add-On Application
            }
            if (!(ConnectToCompany() == 0))
            {
                SBO_Application.StatusBar.SetText("Failed connecting to the company's Data Base", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                System.Environment.Exit(0); //  Terminating the Add-On Application
            }

            SBO_Application.StatusBar.SetText("DI Connected To: " + oCompany.CompanyName, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            
            createUDF();
            AddMenuItems();

            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            if (EventType == SAPbouiCOM.BoAppEventTypes.aet_ShutDown || EventType == SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition || EventType == SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged)
            {
                try
                {
                    System.Windows.Forms.Application.Exit();

                }
                catch (Exception)
                {

                    //    throw;
                }

            }
        }
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            if (pVal.MenuUID == "BO_POCSV")
            {
                LoadFromXML("csvpo_frm.srf");
            }
            BubbleEvent = true;
        }
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            try
            {
                #region Clear path
                if (pVal.FormUID == "csvpo_frm" && pVal.ItemUID == "cpath_btn" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Form ofrom;
                        SAPbouiCOM.EditText path;

                        ofrom = SBO_Application.Forms.GetForm("csvpo_frm", -1);
                        path = (SAPbouiCOM.EditText)ofrom.Items.Item("path_txt").Specific;

                        path.Value = "";
                    }
                }
                #endregion

                #region read text file and load po to matrix
                if (pVal.FormUID == "csvpo_frm" && pVal.ItemUID == "load_btn" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Form ofrom;
                        SAPbouiCOM.EditText path;

                        Global.datatablecreate();
                        Global.dt.Clear();
                        ofrom = SBO_Application.Forms.GetForm("csvpo_frm", -1);
                        path = (SAPbouiCOM.EditText)ofrom.Items.Item("path_txt").Specific;

                        StreamReader f = new StreamReader(path.Value.ToString());

                        String line;
                        String Col1, Col2, Col3, Col4, Col5, Col6, Col7, Col8, Col9, Col10, Col11, Col12, Col13, Col14, Col15, Col16, Col17, Col18, Col19, Col20, Col21, Col22, Col23, Col24, Col25, Col26, Col27, Col28, Col29, Col30, Col31, Col32, Col33, Col34, Col35, Col36, Col37, Col38, Col39, Col40, Col41, Col42, Col43, Col44, Col45, Col46, Col47, Col48, Col49, Col50, Col51, Col52, Col53, Col54, Col55, Col56, Col57, Col58, Col59, Col60, Col61, Col62, Col63, Col64, Col65, Col66, Col67, Col68, Col69, Col70, Col71, Col72, Col73, Col74, Col75, Col76, Col77, Col78, Col79, Col80, Col81, Col82, Col83, Col84, Col85, Col86, Col87, Col88, Col89, Col90, Col91, Col92, Col93, Col94, Col95, Col96, Col97, Col98, Col99, Col100, Col101, Col102, Col103, Col104, Col105, Col106, Col107, Col108, Col109, Col110, Col111, Col112, Col113, Col114, Col115, Col116, Col117, Col118, Col119, Col120, Col121, Col122, Col123, Col124, Col125, Col126, Col127, Col128, Col129, Col130, Col131, Col132, Col133, Col134, Col135, Col136, Col137, Col138, Col139, Col140, Col141, Col142, Col143, Col144, Col145, Col146, Col147, Col148, Col149, Col150, Col151, Col152, Col153, Col154, Col155, Col156, Col157, Col158, Col159, Col160, Col161, Col162, Col163, Col164, Col165, Col166, Col167, Col168, Col169, Col170, Col171, Col172, Col173, Col174, Col175, Col176, Col177, Col178, Col179, Col180, Col181, Col182, Col183, Col184, Col185, Col186, Col187, Col188, Col189, Col190, Col191, Col192, Col193, Col194, Col195, Col196, Col197, Col198, Col199, Col200, Col201, Col202, Col203, Col204, Col205, Col206, Col207, Col208, Col209, Col210, Col211, Col212, Col213, Col214, Col215, Col216, Col217, Col218, Col219, Col220, Col221, Col222, Col223, Col224, Col225, Col226, Col227, Col228, Col229, Col230, Col231, Col232, Col233, Col234, Col235, Col236, Col237, Col238, Col239, Col240, Col241, Col242, Col243, Col244, Col245, Col246, Col247, Col248, Col249, Col250, Col251, Col252, Col253, Col254, Col255, Col256;
                        
                        while ((line = f.ReadLine()) != null)
                        {
                            String[] strings = line.Split(new char[] { '|' });
                            if (strings.Length == 256)
                            {
                                Col1 = strings[0];
                                Col2 = strings[1];
                                Col3 = strings[2];
                                Col4 = strings[3];
                                Col5 = strings[4];
                                Col6 = strings[5];
                                Col7 = strings[6];
                                Col8 = strings[7];
                                Col9 = strings[8];
                                Col10 = strings[9];
                                Col11 = strings[10];
                                Col12 = strings[11];
                                Col13 = strings[12];
                                Col14 = strings[13];
                                Col15 = strings[14];
                                Col16 = strings[15];
                                Col17 = strings[16];
                                Col18 = strings[17];
                                Col19 = strings[18];
                                Col20 = strings[19];
                                Col21 = strings[20];
                                Col22 = strings[21];
                                Col23 = strings[22];
                                Col24 = strings[23];
                                Col25 = strings[24];
                                Col26 = strings[25];
                                Col27 = strings[26];
                                Col28 = strings[27];
                                Col29 = strings[28];
                                Col30 = strings[29];
                                Col31 = strings[30];
                                Col32 = strings[31];
                                Col33 = strings[32];
                                Col34 = strings[33];
                                Col35 = strings[34];
                                Col36 = strings[35];
                                Col37 = strings[36];
                                Col38 = strings[37];
                                Col39 = strings[38];
                                Col40 = strings[39];
                                Col41 = strings[40];
                                Col42 = strings[41];
                                Col43 = strings[42];
                                Col44 = strings[43];
                                Col45 = strings[44];
                                Col46 = strings[45];
                                Col47 = strings[46];
                                Col48 = strings[47];
                                Col49 = strings[48];
                                Col50 = strings[49];
                                Col51 = strings[50];
                                Col52 = strings[51];
                                Col53 = strings[52];
                                Col54 = strings[53];
                                Col55 = strings[54];
                                Col56 = strings[55];
                                Col57 = strings[56];
                                Col58 = strings[57];
                                Col59 = strings[58];
                                Col60 = strings[59];
                                Col61 = strings[60];
                                Col62 = strings[61];
                                Col63 = strings[62];
                                Col64 = strings[63];
                                Col65 = strings[64];
                                Col66 = strings[65];
                                Col67 = strings[66];
                                Col68 = strings[67];
                                Col69 = strings[68];
                                Col70 = strings[69];
                                Col71 = strings[70];
                                Col72 = strings[71];
                                Col73 = strings[72];
                                Col74 = strings[73];
                                Col75 = strings[74];
                                Col76 = strings[75];
                                Col77 = strings[76];
                                Col78 = strings[77];
                                Col79 = strings[78];
                                Col80 = strings[79];
                                Col81 = strings[80];
                                Col82 = strings[81];
                                Col83 = strings[82];
                                Col84 = strings[83];
                                Col85 = strings[84];
                                Col86 = strings[85];
                                Col87 = strings[86];
                                Col88 = strings[87];
                                Col89 = strings[88];
                                Col90 = strings[89];
                                Col91 = strings[90];
                                Col92 = strings[91];
                                Col93 = strings[92];
                                Col94 = strings[93];
                                Col95 = strings[94];
                                Col96 = strings[95];
                                Col97 = strings[96];
                                Col98 = strings[97];
                                Col99 = strings[98];
                                Col100 = strings[99];
                                Col101 = strings[100];
                                Col102 = strings[101];
                                Col103 = strings[102];
                                Col104 = strings[103];
                                Col105 = strings[104];
                                Col106 = strings[105];
                                Col107 = strings[106];
                                Col108 = strings[107];
                                Col109 = strings[108];
                                Col110 = strings[109];
                                Col111 = strings[110];
                                Col112 = strings[111];
                                Col113 = strings[112];
                                Col114 = strings[113];
                                Col115 = strings[114];
                                Col116 = strings[115];
                                Col117 = strings[116];
                                Col118 = strings[117];
                                Col119 = strings[118];
                                Col120 = strings[119];
                                Col121 = strings[120];
                                Col122 = strings[121];
                                Col123 = strings[122];
                                Col124 = strings[123];
                                Col125 = strings[124];
                                Col126 = strings[125];
                                Col127 = strings[126];
                                Col128 = strings[127];
                                Col129 = strings[128];
                                Col130 = strings[129];
                                Col131 = strings[130];
                                Col132 = strings[131];
                                Col133 = strings[132];
                                Col134 = strings[133];
                                Col135 = strings[134];
                                Col136 = strings[135];
                                Col137 = strings[136];
                                Col138 = strings[137];
                                Col139 = strings[138];
                                Col140 = strings[139];
                                Col141 = strings[140];
                                Col142 = strings[141];
                                Col143 = strings[142];
                                Col144 = strings[143];
                                Col145 = strings[144];
                                Col146 = strings[145];
                                Col147 = strings[146];
                                Col148 = strings[147];
                                Col149 = strings[148];
                                Col150 = strings[149];
                                Col151 = strings[150];
                                Col152 = strings[151];
                                Col153 = strings[152];
                                Col154 = strings[153];
                                Col155 = strings[154];
                                Col156 = strings[155];
                                Col157 = strings[156];
                                Col158 = strings[157];
                                Col159 = strings[158];
                                Col160 = strings[159];
                                Col161 = strings[160];
                                Col162 = strings[161];
                                Col163 = strings[162];
                                Col164 = strings[163];
                                Col165 = strings[164];
                                Col166 = strings[165];
                                Col167 = strings[166];
                                Col168 = strings[167];
                                Col169 = strings[168];
                                Col170 = strings[169];
                                Col171 = strings[170];
                                Col172 = strings[171];
                                Col173 = strings[172];
                                Col174 = strings[173];
                                Col175 = strings[174];
                                Col176 = strings[175];
                                Col177 = strings[176];
                                Col178 = strings[177];
                                Col179 = strings[178];
                                Col180 = strings[179];
                                Col181 = strings[180];
                                Col182 = strings[181];
                                Col183 = strings[182];
                                Col184 = strings[183];
                                Col185 = strings[184];
                                Col186 = strings[185];
                                Col187 = strings[186];
                                Col188 = strings[187];
                                Col189 = strings[188];
                                Col190 = strings[189];
                                Col191 = strings[190];
                                Col192 = strings[191];
                                Col193 = strings[192];
                                Col194 = strings[193];
                                Col195 = strings[194];
                                Col196 = strings[195];
                                Col197 = strings[196];
                                Col198 = strings[197];
                                Col199 = strings[198];
                                Col200 = strings[199];
                                Col201 = strings[200];
                                Col202 = strings[201];
                                Col203 = strings[202];
                                Col204 = strings[203];
                                Col205 = strings[204];
                                Col206 = strings[205];
                                Col207 = strings[206];
                                Col208 = strings[207];
                                Col209 = strings[208];
                                Col210 = strings[209];
                                Col211 = strings[210];
                                Col212 = strings[211];
                                Col213 = strings[212];
                                Col214 = strings[213];
                                Col215 = strings[214];
                                Col216 = strings[215];
                                Col217 = strings[216];
                                Col218 = strings[217];
                                Col219 = strings[218];
                                Col220 = strings[219];
                                Col221 = strings[220];
                                Col222 = strings[221];
                                Col223 = strings[222];
                                Col224 = strings[223];
                                Col225 = strings[224];
                                Col226 = strings[225];
                                Col227 = strings[226];
                                Col228 = strings[227];
                                Col229 = strings[228];
                                Col230 = strings[229];
                                Col231 = strings[230];
                                Col232 = strings[231];
                                Col233 = strings[232];
                                Col234 = strings[233];
                                Col235 = strings[234];
                                Col236 = strings[235];
                                Col237 = strings[236];
                                Col238 = strings[237];
                                Col239 = strings[238];
                                Col240 = strings[239];
                                Col241 = strings[240];
                                Col242 = strings[241];
                                Col243 = strings[242];
                                Col244 = strings[243];
                                Col245 = strings[244];
                                Col246 = strings[245];
                                Col247 = strings[246];
                                Col248 = strings[247];
                                Col249 = strings[248];
                                Col250 = strings[249];
                                Col251 = strings[250];
                                Col252 = strings[251];
                                Col253 = strings[252];
                                Col254 = strings[253];
                                Col255 = strings[254];
                                Col256 = strings[255];                                
                                Global.dt.Rows.Add(Col1,	Col2,	Col3,	Col4,	Col5,	Col6,	Col7,	Col8,	Col9,	Col10,	Col11,	Col12,	Col13,	Col14,	Col15,	Col16,	Col17,	Col18,	Col19,	Col20,	Col21,	Col22,	Col23,	Col24,	Col25,	Col26,	Col27,	Col28,	Col29,	Col30,	Col31,	Col32,	Col33,	Col34,	Col35,	Col36,	Col37,	Col38,	Col39,	Col40,	Col41,	Col42,	Col43,	Col44,	Col45,	Col46,	Col47,	Col48,	Col49,	Col50,	Col51,	Col52,	Col53,	Col54,	Col55,	Col56,	Col57,	Col58,	Col59,	Col60,	Col61,	Col62,	Col63,	Col64,	Col65,	Col66,	Col67,	Col68,	Col69,	Col70,	Col71,	Col72,	Col73,	Col74,	Col75,	Col76,	Col77,	Col78,	Col79,	Col80,	Col81,	Col82,	Col83,	Col84,	Col85,	Col86,	Col87,	Col88,	Col89,	Col90,	Col91,	Col92,	Col93,	Col94,	Col95,	Col96,	Col97,	Col98,	Col99,	Col100,	Col101,	Col102,	Col103,	Col104,	Col105,	Col106,	Col107,	Col108,	Col109,	Col110,	Col111,	Col112,	Col113,	Col114,	Col115,	Col116,	Col117,	Col118,	Col119,	Col120,	Col121,	Col122,	Col123,	Col124,	Col125,	Col126,	Col127,	Col128,	Col129,	Col130,	Col131,	Col132,	Col133,	Col134,	Col135,	Col136,	Col137,	Col138,	Col139,	Col140,	Col141,	Col142,	Col143,	Col144,	Col145,	Col146,	Col147,	Col148,	Col149,	Col150,	Col151,	Col152,	Col153,	Col154,	Col155,	Col156,	Col157,	Col158,	Col159,	Col160,	Col161,	Col162,	Col163,	Col164,	Col165,	Col166,	Col167,	Col168,	Col169,	Col170,	Col171,	Col172,	Col173,	Col174,	Col175,	Col176,	Col177,	Col178,	Col179,	Col180,	Col181,	Col182,	Col183,	Col184,	Col185,	Col186,	Col187,	Col188,	Col189,	Col190,	Col191,	Col192,	Col193,	Col194,	Col195,	Col196,	Col197,	Col198,	Col199,	Col200,	Col201,	Col202,	Col203,	Col204,	Col205,	Col206,	Col207,	Col208,	Col209,	Col210,	Col211,	Col212,	Col213,	Col214,	Col215,	Col216,	Col217,	Col218,	Col219,	Col220,	Col221,	Col222,	Col223,	Col224,	Col225,	Col226,	Col227,	Col228,	Col229,	Col230,	Col231,	Col232,	Col233,	Col234,	Col235,	Col236,	Col237,	Col238,	Col239,	Col240,	Col241,	Col242,	Col243,	Col244,	Col245,	Col246,	Col247,	Col248,	Col249,	Col250,	Col251,	Col252,	Col253,	Col254,	Col255,	Col256);
                            }
                        }
                        f.Close();

                        DataTable ddt = new DataTable();
                        string[] TobeDistinct = { "Col5", "Col24", "Col25", "Col7", "Col8", "Col27" };
                        DataTable dtDistinct = Global.GetDistinctRecords(Global.dt, TobeDistinct);
                      
                        #region comment

                        //sql = "Execute [BOS_CSVREAD_ConCur]  '" + path.Value.ToString() + "','H'";
                        //SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //ors.DoQuery(sql);

                        SAPbouiCOM.Matrix omat = (SAPbouiCOM.Matrix)ofrom.Items.Item("omat").Specific;
                        omat.Columns.Item("V_0").Editable = true;
                        omat.Columns.Item("V_7").Editable = true;
                        omat.Columns.Item("V_6").Editable = true;
                        omat.Columns.Item("V_5").Editable = true;
                        omat.Columns.Item("V_4").Editable = true;

                        omat.Clear();

                        if (dtDistinct.Rows.Count > 0)
                        {
                            ofrom.Freeze(true);
                            SBO_Application.StatusBar.SetText("Please wait documents loding.......", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            //while (ors.EoF == false)
                            foreach (DataRow dr in dtDistinct.Rows)
                            {
                                omat.AddRow(1, -1);

                                SAPbouiCOM.EditText row = (SAPbouiCOM.EditText)omat.Columns.Item("V_-1").Cells.Item(omat.RowCount).Specific;
                                row.Value = omat.RowCount.ToString();

                                SAPbouiCOM.EditText PONUM = (SAPbouiCOM.EditText)omat.Columns.Item("V_0").Cells.Item(omat.RowCount).Specific;
                                PONUM.Value = dr[0].ToString();

                                SAPbouiCOM.EditText DocDate = (SAPbouiCOM.EditText)omat.Columns.Item("V_7").Cells.Item(omat.RowCount).Specific;
                                DocDate.Value = dr[1].ToString().Replace("-","");

                                SAPbouiCOM.EditText DocDueDate = (SAPbouiCOM.EditText)omat.Columns.Item("V_6").Cells.Item(omat.RowCount).Specific;
                                DocDueDate.Value = dr[2].ToString().Replace("-", "");

                                SAPbouiCOM.EditText CardCode = (SAPbouiCOM.EditText)omat.Columns.Item("V_5").Cells.Item(omat.RowCount).Specific;
                                CardCode.Value = "AMEX0";

                                SAPbouiCOM.EditText NumAtCard = (SAPbouiCOM.EditText)omat.Columns.Item("V_4").Cells.Item(omat.RowCount).Specific;
                                NumAtCard.Value = dr[3].ToString() + " "+ dr[4].ToString()+ " "+ dr[5].ToString() + " CONCUR IMPORT";

                            }
                            SBO_Application.StatusBar.SetText("Completed documents loding.......", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                            ofrom.Freeze(false);
                        }
                        ofrom.Items.Item("6").Click();
                        omat.Columns.Item("V_0").Editable = false;
                        omat.Columns.Item("V_7").Editable = false;
                        omat.Columns.Item("V_6").Editable = false;
                        omat.Columns.Item("V_5").Editable = false;
                        omat.Columns.Item("V_4").Editable = false;

                        #endregion
                    }
                }
                #endregion

                #region read details and create po in SBO
                if (pVal.FormUID == "csvpo_frm" && pVal.ItemUID == "crpo_btn" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                {
                    if (pVal.BeforeAction == false)
                    {
                        bool oError = false;
                        SAPbouiCOM.Form ofrom;
                        ofrom = SBO_Application.Forms.GetForm("csvpo_frm", -1);
                        
                        SAPbouiCOM.Matrix omat = (SAPbouiCOM.Matrix)ofrom.Items.Item("omat").Specific;
                        string ofilename = Convert.ToString(((SAPbouiCOM.EditText)ofrom.Items.Item("path_txt").Specific).Value);

                        if (!oCompany.InTransaction)
                            oCompany.StartTransaction();

                        Global.BPERR = Global.BPERR + "Purchase Order creation process started***********************************************************************\r\n";
                        for (int i = 0; i < omat.RowCount; i++)
                        {
                            SAPbouiCOM.EditText PONUM = (SAPbouiCOM.EditText)omat.Columns.Item("V_0").Cells.Item(i + 1).Specific;

                            SAPbouiCOM.EditText DocDate = (SAPbouiCOM.EditText)omat.Columns.Item("V_7").Cells.Item(i + 1).Specific;
                            SAPbouiCOM.EditText DocDueDate = (SAPbouiCOM.EditText)omat.Columns.Item("V_6").Cells.Item(i + 1).Specific;
                            SAPbouiCOM.EditText CardCode = (SAPbouiCOM.EditText)omat.Columns.Item("V_5").Cells.Item(i + 1).Specific;
                            SAPbouiCOM.EditText NumAtCard = (SAPbouiCOM.EditText)omat.Columns.Item("V_4").Cells.Item(i + 1).Specific;

                            SAPbouiCOM.EditText status = (SAPbouiCOM.EditText)omat.Columns.Item("V_1").Cells.Item(i + 1).Specific;
                            SAPbouiCOM.EditText SBOPO = (SAPbouiCOM.EditText)omat.Columns.Item("V_2").Cells.Item(i + 1).Specific;
                            SAPbouiCOM.EditText Message = (SAPbouiCOM.EditText)omat.Columns.Item("V_3").Cells.Item(i + 1).Specific;

                            try
                            {
                                if (Global.getPOCount(oCompany, PONUM.Value.ToString()) == 0)
                                {
                                    SAPbobsCOM.Documents oPO = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

                                    oPO.CardCode = CardCode.Value.ToString();
                                  //  SBO_Application.MessageBox(DocDate.Value.ToString().Replace("/", ""));
                                    //DateTime oDt = DateTime.Parse(DocDate.Value.ToString().Insert(4,"-").Insert(7,"-")) ;
                                    oPO.DocDate = DateTime.Parse(DocDate.Value.ToString().Insert(4, "-").Insert(7, "-"));
                                    oPO.DocDueDate = DateTime.Parse(DocDueDate.Value.ToString().Insert(4, "-").Insert(7, "-"));
                                    oPO.NumAtCard = NumAtCard.Value.ToString();

                                    DataRow[] rows = Global.dt.Select("Col5 = " + PONUM.Value.ToString() + "");

                                    int oRecordIndex = 1;
                                    foreach (DataRow row in rows)
                                    {
                                        string oItemDesc = row[62].ToString() + " " + row[68].ToString() + " (" + row[6].ToString() + " " + row[5].ToString() + ")";
                                        oPO.Lines.ItemCode = Global.getItemCode(oCompany, row[62].ToString());
                                        
                                        if (oItemDesc.Length > 99)
                                            oPO.Lines.ItemDescription = oItemDesc.Substring(0, 99);
                                        else
                                            oPO.Lines.ItemDescription = row[62].ToString() + " " + row[68].ToString() + " (" + row[6].ToString() + " " + row[5].ToString() + ")";


                                        try
                                        {
                                            //-----------------------------------------------------------------------------------------------
                                            // Validations for error log
                                            //-----------------------------------------------------------------------------------------------
                                            string oErrorNew = "";
                                            SAPbobsCOM.Recordset oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            string oStringQuery = "select Count(*) from OACT where AcctCode='" + Global.getAccountCode(oCompany, row[62].ToString()) + "'";
                                            oRecordset.DoQuery(oStringQuery);
                                            if (Convert.ToInt32(oRecordset.Fields.Item(0).Value) <= 0)
                                                oErrorNew = "Expense Type - '" + row[62].ToString() + "' Not found, ";

                                            oRecordset = null;
                                            oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            oStringQuery = "select Count(*) from OITM where ItemCode='" + Global.getItemCode(oCompany, row[62].ToString()) + "'";
                                            oRecordset.DoQuery(oStringQuery);
                                            if (Convert.ToInt32(oRecordset.Fields.Item(0).Value) <= 0)
                                                oErrorNew = oErrorNew + "Item Code not found, ";

                                            oRecordset = null;
                                            oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            oStringQuery = "select Count(*) from OPRC where DimCode = 1 and PrcCode ='" + Global.getDepartment(oCompany, row[5].ToString(), row[6].ToString()) + "'";
                                            oRecordset.DoQuery(oStringQuery);
                                            if (Convert.ToInt32(oRecordset.Fields.Item(0).Value) <= 0)
                                                oErrorNew = oErrorNew + "Department not found, ";

                                            oRecordset = null;
                                            oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            oStringQuery = "select Count(*) from OPRC where DimCode = 2 and PrcCode ='" + Global.getPartner(oCompany, row[5].ToString(), row[6].ToString()) + "'";
                                            oRecordset.DoQuery(oStringQuery);
                                            if (Convert.ToInt32(oRecordset.Fields.Item(0).Value) <= 0)
                                                oErrorNew = oErrorNew + "Partner not found, ";

                                            oRecordset = null;
                                            oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            oStringQuery = "select Count(*) from OPRC where DimCode = 3 and PrcCode ='" + Global.getMngUnit(oCompany, row[5].ToString(), row[6].ToString()) + "'";
                                            oRecordset.DoQuery(oStringQuery);
                                            if (Convert.ToInt32(oRecordset.Fields.Item(0).Value) <= 0)
                                                oErrorNew = oErrorNew + "Management Unit not found, ";

                                            oRecordset = null;
                                            oRecordset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            oStringQuery = "select Count(*) from OPRC where DimCode = 4 and PrcCode ='" + Global.getProduct(oCompany, row[5].ToString(), row[6].ToString()) + "'";
                                            oRecordset.DoQuery(oStringQuery);
                                            if (Convert.ToInt32(oRecordset.Fields.Item(0).Value) <= 0)
                                                oErrorNew = oErrorNew + "Product not found, ";

                                            if(Double.Parse(row[152].ToString())<=0)
                                                oErrorNew = oErrorNew + "Document total should be greater then 0";

                                            if (oErrorNew != "")
                                            {
                                                oError = true;
                                                oErrorNew = oErrorNew.Substring(0, oErrorNew.Length - 2);
                                                Global.ERR = Global.ERR + Environment.NewLine + "[" + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt") + "] " +  oErrorNew;
                                                continue;
                                            }
                                            //-----------------------------------------------------------------------------------------------

                                        }
                                        catch 
                                        {
                                            
                                        }
                                        
                                        

                                        oPO.Lines.Quantity = 1;
                                        oPO.Lines.PriceAfterVAT = Double.Parse(row[152].ToString());
                                        oPO.Lines.AccountCode = Global.getAccountCode(oCompany, row[62].ToString());
                                        oPO.Lines.COGSCostingCode = Global.getDepartment(oCompany, row[5].ToString(), row[6].ToString());
                                        oPO.Lines.COGSCostingCode2 = Global.getPartner(oCompany, row[5].ToString(), row[6].ToString());
                                        oPO.Lines.COGSCostingCode3 = Global.getMngUnit(oCompany, row[5].ToString(), row[6].ToString());
                                        oPO.Lines.COGSCostingCode4 = Global.getProduct(oCompany, row[5].ToString(), row[6].ToString());

                                        if (row[237].ToString().Equals(""))
                                        {
                                            oPO.Lines.VatGroup = "G14";
                                        }
                                        else
                                        {
                                            oPO.Lines.VatGroup = row[237].ToString();
                                        }

                                        if (oRecordIndex < rows.Length)
                                            oPO.Lines.Add();

                                        oRecordIndex++;
                                    }

                                    oPO.UserFields.Fields.Item("U_Concur").Value = "Yes";
                                    oPO.UserFields.Fields.Item("U_FileName").Value = ofilename;

                                    long retCode = oPO.Add();

                                    if (retCode != 0)
                                    {
                                        Global.BPERR = Global.BPERR + "Purchase Order creation failed. Error : " +oCompany.GetLastErrorDescription().ToString() +  " PO Number from file is :" + PONUM.Value.ToString() + "\r\n";
                                        Global.ERR = Global.ERR + Environment.NewLine + "[" + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt") + "] "  + "Purchase Order creation failed. Error : " + oCompany.GetLastErrorDescription().ToString() + " PO Number from file is :" + PONUM.Value.ToString() + "\r\n";
                                        status.Value = "Failed";
                                        oError = true;
                                       
                                    }
                                    else
                                    {
                                        Global.BPERR = Global.BPERR + "Purchase Order creation sucssfully : PO Number from file is :" + PONUM.Value.ToString() + "\r\n";
                                        Global.ERR = Global.ERR + Environment.NewLine + "[" + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt") + "] " + "Purchase Order creation sucssfully : PO Number from file is :" + PONUM.Value.ToString() + "\r\n";
                                        status.Value = "Sucess";
                                        SBOPO.Value = Global.getpodocnum(oCompany);

                                    }

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPO);
                                    oPO = null;
                                    GC.Collect();
                                }
                                else
                                {
                                    oError = true;
                                    Global.BPERR = Global.BPERR + "Purchase Order already exits in the Database :" + PONUM.Value.ToString() + "\r\n";
                                    Global.ERR = Global.ERR + Environment.NewLine + "[" + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt") + "] " + "Purchase Order already exits in the Database :" + PONUM.Value.ToString() + "\r\n";
                                
                                }
                            }
                            catch (Exception ex)
                            {
                                oError = true;
                                Global.BPERR = Global.BPERR +  ""  + ex.Message + "\r\n";
                                Global.ERR = Global.ERR + Environment.NewLine + "[" + DateTime.Now.ToString("dd-MMM-yyyy hh:mm:ss tt") + "] " + ex.Message + "\r\n";
                            }

                            writePOlog(PONUM.Value.ToString());
                            Message.Value = System.Windows.Forms.Application.StartupPath + @"\ErrLogs\ErrLog_" + PONUM.Value.ToString() + ".txt";

                        }

                        if (oError)
                        {
                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                           // return "";
                            
                        }
                        else
                        {
                            if (oCompany.InTransaction)
                                oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        }


                        Global.BPERR = Global.BPERR + "Purchase Order creation process is completed***********************************************************************\r\n";
                        writetolog();
                        SBO_Application.StatusBar.SetText("Purchase Order creation process is completed", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message);
            }
            BubbleEvent = true;
        }


        public void writePOlog(string po)
        {
            if (Global.BPERR != null)
            {
                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\ErrLogs"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\ErrLogs");
                }
                FileStream newfile1 = new FileStream(System.Windows.Forms.Application.StartupPath + @"\ErrLogs\ErrLog_" + po +  ".txt", FileMode.Append, FileAccess.Write);
                TextWriter tw1 = new StreamWriter(newfile1);
                tw1.Write(Global.ERR);
                tw1.Close();

                Global.ERR = "";
                
            }
        }

        public void writetolog()
        {
            if (Global.BPERR != null)
            {
                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Logs"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Logs");
                }
                FileStream newfile = new FileStream(System.Windows.Forms.Application.StartupPath + @"\Logs\Logfile_" + Global.globaltime + ".txt", FileMode.Append, FileAccess.Write);
                TextWriter tw = new StreamWriter(newfile);
                tw.Write(Global.BPERR);
                tw.Close();
                Global.BPERR = "";
            }
        }
        public void LoadFromXML(string filename)
        {
            try
            {
                System.Xml.XmlDocument oXmlDoc = null;
                string sPath = null;
                oXmlDoc = new System.Xml.XmlDocument();
                sPath = System.Windows.Forms.Application.StartupPath + "\\" + filename;
                oXmlDoc.Load(sPath);
                string sXML = oXmlDoc.InnerXml.ToString();
                SBO_Application.LoadBatchActions(ref sXML);
            }
            catch (Exception ex)
            {
            }
        }

        private void AddMenuItems()
        {

            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = SBO_Application.Menus.Item("2304"); // Purchase moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "BO_PO";
            oCreationPackage.String = "Create Purchase Order";
            oCreationPackage.Enabled = true;

            oMenus = oMenuItem.SubMenus;

            try
            {
                oMenus.AddEx(oCreationPackage);
                oMenuItem = SBO_Application.Menus.Item("BO_PO");
                oMenus = oMenuItem.SubMenus;

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "BO_POCSV";
                oCreationPackage.String = "Purchase Order From CSVConcur";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception er)
            {
                //SBO_Application.MessageBox("Menu Already Exists", 1, "Ok", "", "");
            }

        }
        public void createUDF()
        {
            int lRetCode = 0;
            string sErrMsg;

            try
            {

                SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = "OPOR";
                oUserFieldsMD.Name = "csvcur";
                oUserFieldsMD.Description = "CSV concur file path";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
                oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link;

                lRetCode = oUserFieldsMD.Add();

                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out lRetCode, out sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();


                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = "OPOR";
                oUserFieldsMD.Name = "concurpo";
                oUserFieldsMD.Description = "CSV Concur PO Number";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Alpha;
                oUserFieldsMD.EditSize = 10;

                lRetCode = oUserFieldsMD.Add();

                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out lRetCode, out sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();

                oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUserFieldsMD.TableName = "POR1";
                oUserFieldsMD.Name = "csvcur";
                oUserFieldsMD.Description = "CSV concur file path";
                oUserFieldsMD.Type = SAPbobsCOM.BoFieldTypes.db_Memo;
                oUserFieldsMD.SubType = SAPbobsCOM.BoFldSubTypes.st_Link;

                lRetCode = oUserFieldsMD.Add();

                if (lRetCode != 0)
                {
                    oCompany.GetLastError(out lRetCode, out sErrMsg);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
            }
        }
    }
}
