using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using MongoDB.Bson;
using MongoDB.Driver;

namespace CovertToFirebase
{
    class cProject
    {
        public string code;
        public string name;
        public string country;
        public string location;
        public string startdate;
        public string enddate;
        public string client;
        public double value;
        public string status;
        public string coordinator;
        public string keywords;
        public string description;
        public string department;
        public List<cPerson> seniorstaff = new List<cPerson>();
        public List<cServices> services = new List<cServices>();

        public delegate void ErrorEvent(string message);
        public event ErrorEvent LogError;

        public cProject()
        {

        }

        public bool ProcessFile(Excel._Worksheet worksheet, string dept, ref List<string> ids)
        {
            string worksheetname = "";
            Excel.Range xlRange = worksheet.UsedRange;
            worksheetname = worksheet.Name;
            department = dept;

            var val = "";

                //ID
            try
            {
                val = xlRange.Cells[1, 1].Value2;
                if (val != null)
                {
                    string temp = val.ToString();
                    if (temp.Contains(":"))
                    {
                        string[] ar = temp.Split(':');
                        if (CodeExists(ar[1].Trim()) > 0)
                        {
                            LogError(worksheetname + ": " + "Project code '" + ar[1].Trim() + "' already exists in the database ");
                            return false;
                        }
                        else
                        {
                            code = ar[1].Trim();
                        }
                    }

                }
                else
                    code = "";

                if (code != "" && ids.Contains(code))
                {
                    LogError(worksheetname + ": " + "Duplicate Project ID " + code);
                    return false;
                }

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'ID'.");
            }

                //Name
            try
            {
                val = xlRange.Cells[2, 2].Value2;
                if (val != null)
                {
                    if (PrjNameExists(val.ToString().Trim()) > 0)
                    {
                        LogError(worksheetname + ": " + "Project name '"+ val.ToString().Trim() +"' already exists in the database ");
                        return false;
                    }
                    else
                    {
                        name = val.ToString().Trim();
                    }
                }
                else
                    name = "";
            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Project Name'.");
            }

                //Country
            try
            {
                val = xlRange.Cells[2, 6].Value2;
                if (val != null)
                {
                    country = val.ToString().Trim();
                }
                else
                    country = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Country'.");
            }

                //Project Location Within Country
            try
            {
                val = xlRange.Cells[4, 1].Value2;
                if (val != null)
                {
                    location = val.ToString().Trim();
                }
                else
                    location = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Project Location Within Country'.");
            }

                //StartDate
            try
            {
                val = xlRange.Cells[4, 4].Text;
                if (val != null)
                {
                    startdate = val.ToString().Trim();
                }
                else
                    startdate = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Starting Date'.");
            }

                //EndDate
            try
            {
                val = xlRange.Cells[4, 6].Text;
                if (val != null)
                {
                    enddate = val.ToString().Trim();
                }
                else
                    enddate = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Completion Date'.");
            }

                //Client
            try
            {
                val = xlRange.Cells[5, 2].Value2;
                if (val != null)
                {
                    client = val.ToString().Trim();
                }
                else
                    client = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Name of Client'.");
            }

                //ContractValue
            try
            {
                val = xlRange.Cells[5, 6].Text;
                if (val == "-")
                    val = "0";
                if (val != null)
                {
                    value = double.Parse(val);
                }
                else
                    value = 0;

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Total Contract Value (THB)'.");
            }

                //Status
            try
            {
                val = xlRange.Cells[6, 2].Value2;
                if (val != null)
                {
                    status = val.ToString().Trim();
                }
                else
                    status = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Project Status'.");
            }

                //Coordinator
            try
            {
                val = xlRange.Cells[6, 6].Value2;
                if (val != null)
                {
                    coordinator = val.ToString().Trim();
                }
                else
                    coordinator = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Project Coordinator'.");
            }

                //Keywords
            try
            {
                val = xlRange.Cells[7, 2].Value2;
                if (val != null)
                {
                    keywords = val.ToString().Trim();
                }
                else
                    keywords = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Keywords'.");
            }

            var index = 9;

                //ProjectSeniorStaff
            try
            {
                val = xlRange.Cells[index, 1].Value2;
                while (true)
                {
                    if (val != null)
                    {
                        if (val.ToString().Contains("(")) // if the person has a title
                        {
                            var tmp = val.ToString().Trim().Split('(');
                            cPerson p = new cPerson(tmp[0].Trim(), tmp[1].Trim().Replace(")", "").Trim());
                            seniorstaff.Add(p);
                        }
                        else // if the person has no title
                        {
                            cPerson p = new cPerson(val.ToString().Trim(), "");
                            seniorstaff.Add(p);
                        }
                            index++;
                            val = xlRange.Cells[index, 1].Value2;
                            if (val.ToString().Trim() == "Detailed Narrative Description of Project:")
                                break;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Senior Staff Involved and Functions Performed'.");
            }

                //ProjectDescription
            try
            {
                index += 1;
                description = CleanString(xlRange.Cells[index, 1].Value2.ToString());
            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Detailed Narrative Description of Project'.");
            }

                //ProjectServices
            try
            {
                index += 2;
                var ser = new cServices();
                while (true)
                {
                    val = xlRange.Cells[index, 1].Value2;
                    if (val != null)
                    {
                        string line = val.ToString();

                        if (Char.IsUpper(val, 0))
                        {
                            if (ser.Service.Count > 0)
                            {
                                services.Add(ser);
                                ser = new cServices();
                            }

                            ser.Header = line.Split(new char[] { ')' }, 2)[1].Trim();
                        }
                        else if (Char.IsNumber(val, 0))
                        {
                            ser.Service.Add(line.Split(new char[] { ')' }, 2)[1].Trim());
                        }
                        else
                        {
                            LogError(worksheetname + ": " + "Invalid format at field 'Description of Services Provided'.");
                        }
                        index++;
                    }
                    else
                        break;
                }
                services.Add(ser);

            }
            catch (Exception ex)
            {
                LogError(worksheetname + ": " + "Invalid data at field 'Description of Services Provided'.");
            }

            return true;
        }

        public string CleanString(string value)
        {
            var ret = value;
            while (ret.Contains("  "))
            {
                ret = ret.Replace("  ", " ");
            }
            ret = ret.Trim();
            ret = ret.Replace("\n", " ");
            return ret;
        }

        public long CodeExists(string code)
        {
            var connectionString = "mongodb://192.168.42.85:27017";
            var client = new MongoClient(connectionString);
            var db = client.GetDatabase("local");
            var collection = db.GetCollection<cProject>("ztest");

            var count = collection.Count(new BsonDocument("code", code));

            Console.WriteLine("code: " + code + "count: " + count);

            return count;
        }

        public long PrjNameExists(string name)
        {
            var connectionString = "mongodb://192.168.42.85:27017";
            var client = new MongoClient(connectionString);
            var db = client.GetDatabase("local");
            var collection = db.GetCollection<cProject>("ztest");

            var count = collection.Count(new BsonDocument("name", name));

            Console.WriteLine("name: " + name + "count: " + count);

            return count;
        }
    }

}
