using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.Script.Serialization;
using MongoDB.Bson;
using MongoDB.Driver;

namespace CovertToFirebase
{
    public partial class Form1 : Form
    {

        List<string> countriesInDB = new List<string>();
        List<string> clientsInDB = new List<string>();

        public Form1()
        {
            InitializeComponent();
            // Uncomment these to see existing items in the DB through the console:
            // listItemsFromDB("clients");
            // listItemsFromDB("country");
        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            DragDropEffects effect = DragDropEffects.None;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var path = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                FileInfo inf = new FileInfo(path);
                string ext = inf.Extension.ToLower();
                if (ext == ".xls" || ext == ".xlsx")
                    effect = DragDropEffects.Copy;
            }

            e.Effect = effect;
        }

        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var filelist = ((string[])e.Data.GetData(DataFormats.FileDrop));
                textBox1.Text = "";
                for (var i = 0; i < filelist.Length; i++)
                {
                    string path = filelist[i];
                    FileInfo inf = new FileInfo(path);
                    textBox1.Text += ExcelToJSON(path);
                    if (i == filelist.Length - 1)
                    {
                        Proj_LogError("Process done");
                    }
                }
                if (textBox1.Text != "")
                    Clipboard.SetText(textBox1.Text);
            }
        }

        public string ExcelToJSON(string file)
        {

            StringBuilder ret = new StringBuilder();
            if (file != "")
            {
                //Read File
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file);

                for (var i = 1; i <= xlWorkbook.Sheets.Count; i++)
                {
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[i];
                    cProject proj = new cProject();
                    proj.LogError += Proj_LogError;
                    if (proj.ProcessFile(xlWorksheet, textBox2.Text))
                    {
                        ret.AppendLine(new JavaScriptSerializer().Serialize(proj));
                        JSONToMongoDB(new JavaScriptSerializer().Serialize(proj).ToString(), "insert", "", "");
                    }
                    else
                    {
                        string field = "";
                        string val = "";
                        if (proj.code != null)
                        {
                            field = "code";
                            val = proj.code;
                            JSONToMongoDB(new JavaScriptSerializer().Serialize(proj).ToString(), "replace", field, val);
                        }
                        else if (proj.name != null)
                        {
                            field = "name";
                            val = proj.name;
                            JSONToMongoDB(new JavaScriptSerializer().Serialize(proj).ToString(), "replace", field, val);
                        }
                        ret.AppendLine(new JavaScriptSerializer().Serialize(proj));
                    }
                    if (i == xlWorkbook.Sheets.Count)
                    {
                        Proj_LogError("Workbook '" + xlWorkbook.Name + "' done: " + xlWorkbook.Sheets.Count + " worksheet(s)");
                    }
                }
            }

            return ret.ToString(); ;
        }

        private void Proj_LogError(string message)
        {
            StringBuilder bld = new StringBuilder();
            bld.AppendLine(message);
            bld.AppendLine(txtLog.Text);
            txtLog.Text = bld.ToString();
        }

        public void JSONToMongoDB(string file, string act, string field, string val)
        {
            var connectionString = "mongodb://192.168.42.85:27017";
            var client = new MongoClient(connectionString);
            var db = client.GetDatabase("local");
            var collection = db.GetCollection<BsonDocument>("ztest");

            MongoDB.Bson.BsonDocument document = MongoDB.Bson.Serialization.BsonSerializer.Deserialize<BsonDocument>(file);

            switch (act)
            {
                case "insert":
                    collection.InsertOne(document);
                    break;
                case "replace":
                    var filter = new BsonDocument(field, val);
                    collection.FindOneAndReplace(filter, document);
                    break;
                default:
                    Console.WriteLine("You just broke the code.");
                    break;
            }
        }

        public void listItemsFromDB(string table)
        {
            var connectionString = "mongodb://192.168.42.85:27017";
            var client = new MongoClient(connectionString);
            var db = client.GetDatabase("local");
            var filter = new BsonDocument();
            switch (table)
            {
                case "clients":
                    clientsInDB.Clear();
                    var clientCollection = db.GetCollection<BsonDocument>(table);
                    using (var cursor = clientCollection.Find(filter).ToCursor())
                    {
                        while (cursor.MoveNext())
                        {
                            foreach (var doc in cursor.Current)
                            {
                                clientsInDB.Add(doc["name"].ToString());
                            }
                            //for reading
                            foreach (string c in clientsInDB)
                            {
                                Console.WriteLine("clientsInDB: " + c);
                            }
                        }
                        Console.WriteLine("Number of clients: " + clientsInDB.Count);//count the number of items in list
                    }
                    break;
                case "country":
                    countriesInDB.Clear();
                    var countryCollection = db.GetCollection<BsonDocument>(table);
                    using (var cursor = countryCollection.Find(filter).ToCursor())
                    {
                        while (cursor.MoveNext())
                        {
                            foreach (var doc in cursor.Current)
                            {
                                countriesInDB.Add(doc["Name"].ToString());
                            }
                            //for reading
                            foreach (string c in countriesInDB)
                            {
                                Console.WriteLine("countriesInDB: " + c);
                            }
                        }
                        Console.WriteLine("Number of countries: " + countriesInDB.Count);//count the number of items in list
                    }
                    break;
            }
        }
    }
}
