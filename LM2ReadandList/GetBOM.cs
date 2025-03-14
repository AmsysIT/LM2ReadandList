using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace LM2ReadandList
{
    public static class GetBOM
    {

        //SQL 參數
        public static string AMS3_ConnectionString { get; set; }
        private static string SelectCmd;
        private static SqlDataAdapter sqlAdapter;
        private static DataTable Dt;
        private static DataTable Dt2;
        private static List<string> BOMList = new List<string>();
        private static List<BomClass> BOMListQTY = new List<BomClass>();

        public  struct BomClass
        {
            public  string PartNo;
            public  decimal QTY;
        }

        public static List<string> GetBOMList(string MB001)
        {
            BOMList.Clear();

            Dt = new DataTable();
            SelectCmd = "SELECT DISTINCT [MB001],[MB003],[MC002],[MB004],[MB005],[MC004],[MB002],[MC019]" +
                " FROM [BOMMB] A LEFT JOIN [INVMC] B ON [MC001] = [MB003]" +
                " WHERE A.[STOP_DATE] IS NULL AND B.[STOP_DATE] IS NULL AND LEN(B.[MC018]) > 0";
            sqlAdapter = new SqlDataAdapter(SelectCmd, AMS3_ConnectionString);
            sqlAdapter.Fill(Dt);

            SetTreeView(MB001);

            return BOMList;
        }

        private static void SetTreeView(string MB001)
        {
            var q1 = (from p in Dt.AsEnumerable()
                      where p.Field<string>("MB001") == MB001
                      select p.Field<string>("MB003") + "," + p.Field<string>("MC019"));

            if(q1.Any())
            {
                foreach(var item in q1)
                {
                    SetTreeView(item.Split(',')[0]);
                    BOMList.Add(item);
                }
            }
        }


        public static List<BomClass> GetBOMListWithQTY(string MB001)
        {
            BOMListQTY.Clear();

            Dt2 = new DataTable();
            SelectCmd = "SELECT [MB001],[MB003],[MC002],[MB004],[MB005],[MC004],[MB002],[MC019]" +
                " FROM [BOMMB] A LEFT JOIN [INVMC] B ON [MC001] = [MB003]" +
                " WHERE A.[STOP_DATE] IS NULL AND B.[STOP_DATE] IS NULL AND LEN(B.[MC018]) > 0" +
                " ORDER BY LEN([MB003]) DESC,[MB003] DESC";
            sqlAdapter = new SqlDataAdapter(SelectCmd, AMS3_ConnectionString);
            sqlAdapter.Fill(Dt2);

            SetTreeView2(MB001);

            return BOMListQTY;
        }

        private static void SetTreeView2(string MB001)
        {
            var q1 = (from p in Dt2.AsEnumerable()
                      where p.Field<string>("MB001") == MB001
                      select new {PartNo = p.Field<string>("MB003"),Element = p.Field<decimal>("MB005") });

            if(q1.Any())
            {
                foreach(var item in q1)
                {
                    BomClass BC = new BomClass
                    {
                        PartNo = item.PartNo,
                        QTY = item.Element
                    };
                    SetTreeView2(item.PartNo);
                    BOMListQTY.Add(BC);
                }
            }
        }
    }
}