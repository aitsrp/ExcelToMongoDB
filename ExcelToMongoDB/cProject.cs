﻿using System;
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

        public bool ProcessFile(Excel._Worksheet worksheet, string dept)
        {
            string worksheetname = "";
            Excel.Range xlRange = worksheet.UsedRange;
            worksheetname = worksheet.Name;
            department = dept;
            bool unique = true;

            var val = "";

            //ID
            try
            {
                val = xlRange.Cells[1, 1].Value2;
                if (val != null)
                {
                    string temp = val.ToString();
                    string fin = "";

                    if (temp.Contains(":"))
                    {
                        string[] ar = temp.Split(':');
                        fin = ar[1].Trim();
                    }
                    else if (temp.Contains("Project ID "))
                    {
                        fin = temp.Replace("Project ID ", "").Trim();
                    }
                    else
                    {
                        fin = temp.Trim();
                    }
                    code = fin;


                }
                else
                    code = "";

            }
            catch (Exception ex)
            {
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'ID'.");
            }

            //Name
            try
            {
                val = xlRange.Cells[2, 2].Value2;
                if (val != null)
                {
                    name = val.ToString().Trim();
                }
                else
                    name = "";
            }
            catch (Exception ex)
            {
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Project Name'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Country'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Project Location Within Country'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Starting Date'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Completion Date'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Name of Client'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Total Contract Value (THB)'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Project Status'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Project Coordinator'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Keywords'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Senior Staff Involved and Functions Performed'.");
            }

            //ProjectDescription
            try
            {
                index += 1;
                description = CleanString(xlRange.Cells[index, 1].Value2.ToString());
            }
            catch (Exception ex)
            {
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Detailed Narrative Description of Project'.");
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
                            LogError(worksheetname.Trim() + ": " + "Invalid format at field 'Description of Services Provided'.");
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
                LogError(worksheetname.Trim() + ": " + "Invalid data at field 'Description of Services Provided'.");
            }

            if (ClientExists(client) == 0) // if client name does not exist on the official list
            {
                LogError(worksheetname.Trim() + ": " + "client name: " + client + " does not exist on the list.");
                unique = false;
            }
            else if (CountrytExists(country) == 0) // if country does not exist on the official list
            {
                LogError(worksheetname.Trim() + ": " + "country: " + country + " does not exist on the list.");
                unique = false;
            }
            else
            {
                if ((CodeExists(code) > 0) && (PrjNameExists(name) > 0)) // if project code or project name already exists in the database
                {
                    LogError(worksheetname.Trim() + ": Updated entry");
                    unique = false;
                }
            }

            return unique;
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
            var dbclient = new MongoClient(connectionString);
            var db = dbclient.GetDatabase("local");
            var collection = db.GetCollection<cProject>("ztest");

            var count = collection.Count(new BsonDocument("code", code));

            return count;
        }

        public long PrjNameExists(string name)
        {
            var connectionString = "mongodb://192.168.42.85:27017";
            var dbclient = new MongoClient(connectionString);
            var db = dbclient.GetDatabase("local");
            var collection = db.GetCollection<cProject>("ztest");

            var count = collection.Count(new BsonDocument("name", name));

            return count;
        }

        public long ClientExists(string client)
        {
            var connectionString = "mongodb://192.168.42.85:27017";
            var dbclient = new MongoClient(connectionString);
            var db = dbclient.GetDatabase("local");
            var collection = db.GetCollection<cProject>("clients");

            var count = collection.Count(new BsonDocument("name", client));
            Console.WriteLine("count: "+count);

            return count;
        }

        public long CountrytExists(string country)
        {
            var connectionString = "mongodb://192.168.42.85:27017";
            var dbclient = new MongoClient(connectionString);
            var db = dbclient.GetDatabase("local");
            var collection = db.GetCollection<cProject>("country");

            var count = collection.Count(new BsonDocument("Name", country));
            Console.WriteLine("count: " + count);

            return count;
        }

        public string ClientRename(string client)
        {
            var newName = "";
            var list1 = new List<string>();
            string[] array1 = new string[] {};
            list1.AddRange(array1);

            return newName;
        }
    }

}
