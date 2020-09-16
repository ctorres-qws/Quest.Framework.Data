using Quest.Framework.Persistance;
using Quest.Framework.Persistance.OLEDB;
using Quest.ScanToPrint.Data.Entities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Quest.ScanToPrint.Data
{
    public class OLEDBGlazingPersistenceStrategy : OLEDBPersistenceStrategy<Glazing>
    {
        public OLEDBGlazingPersistenceStrategy(string dbConnectionString) : base(dbConnectionString) { }
        protected override void DefineAddCommand(ref OleDbCommand command, Glazing entity)
        {
            command.CommandText = "INSERT INTO X_GLAZING([BARCODE],[JOB],[FLOOR],[TAG],[DEPT],[EMPLOYEE],[OPENINGS],[FirstComplete],[JOINTS],[DATETIME],[DAY],[MONTH],[YEAR],[TIME],[WEEK],[ONUMBER],[SCANCOUNT],[O1],[O2],[O3],[O4],[O5],[O6],[O7],[O8],[Count]) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

            command.Parameters.AddWithValue("@BARCODE", entity.Barcode);
            command.Parameters.AddWithValue("@JOB", entity.Job);
            command.Parameters.AddWithValue("@FLOOR", entity.Floor);
            command.Parameters.AddWithValue("@TAG", entity.Tag);
            command.Parameters.AddWithValue("@DEPT", entity.Dept);
            command.Parameters.AddWithValue("@EMPLOYEE", entity.Employee);
            command.Parameters.AddWithValue("@OPENINGS", entity.Openings);
            command.Parameters.AddWithValue("@FirstComplete", entity.FirstComplete);
            command.Parameters.AddWithValue("@JOINTS", entity.Joints);
            command.Parameters.AddWithValue("@DATETIME", entity.DateTime);
            command.Parameters.AddWithValue("@DAY", entity.Day);
            command.Parameters.AddWithValue("@MONTH", entity.Month);
            command.Parameters.AddWithValue("@YEAR", entity.Year);
            command.Parameters.AddWithValue("@TIME", entity.Time);
            command.Parameters.AddWithValue("@WEEK", entity.Week);
            command.Parameters.AddWithValue("@ONUMBER", entity.ONumber);
            command.Parameters.AddWithValue("@SCANCOUNT", entity.ScanCount);
            command.Parameters.AddWithValue("@O1", entity.O1);
            command.Parameters.AddWithValue("@O2", entity.O2);
            command.Parameters.AddWithValue("@O3", entity.O3);
            command.Parameters.AddWithValue("@O4", entity.O4);
            command.Parameters.AddWithValue("@O5", entity.O5);
            command.Parameters.AddWithValue("@O6", entity.O6);
            command.Parameters.AddWithValue("@O7", entity.O7);
            command.Parameters.AddWithValue("@O8", entity.O8);
            command.Parameters.AddWithValue("@Count", entity.Count);

        }
        protected override void DefineEditCommand(ref OleDbCommand command, Glazing entity)
        {
            command.CommandText = "UPDATE X_GLAZING SET [JOB] = ?, [FLOOR] = ?, [TAG] = ?, [DEPT] = ?, [EMPLOYEE] = ?, [OPENINGS] = ?, [FirstComplete] = ?, [JOINTS] = ?, [DATETIME] = ?, [DAY] = ?, [MONTH] = ?, [YEAR] = ?, [TIME] = ?, [WEEK] = ?, [ONUMBER] = ?, [SCANCOUNT] = ?, [O1] = ?, [O2] = ?, [O3] = ?, [O4] = ?, [O5] = ?, [O6] = ?, [O7] = ?, [O8] = ?, [Count] = ? WHERE [BARCODE] = ?";
            //command.CommandText = "UPDATE X_GLAZING SET [EMPLOYEE] = ? WHERE [BARCODE] = ?";


            command.Parameters.AddWithValue("@JOB", entity.Job);
            command.Parameters.AddWithValue("@FLOOR", entity.Floor);
            command.Parameters.AddWithValue("@TAG", entity.Tag);
            command.Parameters.AddWithValue("@DEPT", entity.Dept);
            command.Parameters.AddWithValue("@EMPLOYEE", entity.Employee);
            command.Parameters.AddWithValue("@OPENINGS", entity.Openings);
            command.Parameters.AddWithValue("@FirstComplete", entity.FirstComplete);
            command.Parameters.AddWithValue("@JOINTS", entity.Joints);
            command.Parameters.AddWithValue("@DATETIME", entity.DateTime);
            command.Parameters.AddWithValue("@DAY", entity.Day);
            command.Parameters.AddWithValue("@MONTH", entity.Month);
            command.Parameters.AddWithValue("@YEAR", entity.Year);
            command.Parameters.AddWithValue("@TIME", entity.Time);
            command.Parameters.AddWithValue("@WEEK", entity.Week);
            command.Parameters.AddWithValue("@ONUMBER", entity.ONumber);
            command.Parameters.AddWithValue("@SCANCOUNT", entity.ScanCount);
            command.Parameters.AddWithValue("@O1", entity.O1);
            command.Parameters.AddWithValue("@O2", entity.O2);
            command.Parameters.AddWithValue("@O3", entity.O3);
            command.Parameters.AddWithValue("@O4", entity.O4);
            command.Parameters.AddWithValue("@O5", entity.O5);
            command.Parameters.AddWithValue("@O6", entity.O6);
            command.Parameters.AddWithValue("@O7", entity.O7);
            command.Parameters.AddWithValue("@O8", entity.O8);
            command.Parameters.AddWithValue("@Count", entity.Count);

            command.Parameters.AddWithValue("@BARCODE", entity.Barcode);
        }
        protected override void DefineDeleteCommand(ref OleDbCommand command, Glazing entity)
        {
            command.CommandText = "DELETE FROM X_GLAZING WHERE [ID] = ?";

            command.Parameters.AddWithValue("@ID", entity.Job);
        }
        protected override void DefineSelectCommand(ref OleDbCommand command, QueryParameters predicate)
        {
            command.CommandText = "SELECT [BARCODE],[JOB],[FLOOR],[TAG],[DEPT],[EMPLOYEE],[OPENINGS],[FirstComplete],[JOINTS],[DATETIME],[DAY],[MONTH],[YEAR],[TIME],[WEEK],[ONUMBER],[SCANCOUNT],[O1],[O2],[O3],[O4],[O5],[O6],[O7],[O8],[Count] FROM X_GLAZING";
            if (predicate != null)
            {
                foreach (string item in predicate.Keys)
                {
                    command.Parameters.AddWithValue(item, predicate[item]);
                }
            }
        }
        protected override Glazing SelectResultItemMapping(DataRow dataRow)
        {
            //DateTime s = (DateTime)dataRow["DATETIME"];
            return new Glazing()
            {
                Barcode = dataRow["BARCODE"].ToString(),
                Job = dataRow["JOB"].ToString(),
                Floor = dataRow["FLOOR"].ToString(),
                Tag = dataRow["TAG"].ToString(),
                Dept= dataRow["DEPT"].ToString(),
                Employee = dataRow["EMPLOYEE"].ToString(),
                Openings = (int) dataRow["OPENINGS"],
                FirstComplete = dataRow["FirstComplete"].ToString(),
                Joints = (int) dataRow["JOINTS"],

                //DateTime = (DateTime) dataRow["DATETIME"],
                Day = (int)dataRow["DAY"],
                Month = (int)dataRow["MONTH"],
                Year = (int)dataRow["YEAR"],
                //Time = (TimeSpan) dataRow["TIME"],
                Week = (int) dataRow["WEEK"],
                ONumber = (int) dataRow["ONUMBER"],
                ScanCount = (int) dataRow["SCANCOUNT"],

                O1 = dataRow["O1"].ToString(),
                O2 = dataRow["O2"].ToString(),
                O3= dataRow["O3"].ToString(),
                O4= dataRow["O4"].ToString(),
                O5= dataRow["O5"].ToString(),
                O6= dataRow["O6"].ToString(),
                O7= dataRow["O7"].ToString(),
                O8 = dataRow["O8"].ToString(),

                Count = (int) dataRow["Count"]
            };
        }
    }
}
