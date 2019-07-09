using System;
using System.Collections.Generic;
using System.IO;
using System.Data.SqlClient;

namespace Data_Export_v2
{
    class Program
    {
        static string docCSV = @"export\Document List.csv";
        static string imgCSV = @"export\Image List.csv";
        static string catCSV = @"export\Product Attributes.csv"; 

        private static void makeFileList()
        {
            Directory.CreateDirectory("Export"); //create folder "Export", include System.IO

            string imgSQL = "SELECT ProductID, FileName FROM MSCS_IDEC_PRD.dbo.ProductImages WHERE ProductID IN (";
            string docSQL = "SELECT ProductID, FileLocation, DisplayName FROM MSCS_IDEC_PRD.dbo.ProductDocuments WHERE ProductID IN (";
            string catSQL = "SELECT * FROM MSCS_IDEC_PRD.dbo.IDEC_USA_CatalogProducts WHERE ProductID IN (";

            const string SQLUser = "activity_read"; //default login account, the account is configured on the .204 database 
            const string SQLPassword = "yaQ6staf";  //default password 
            const string SQLIP = "192.168.110.204"; //always connected to the IP 192.168.110.204
            const string dbName = "MSCS_IDEC_PRD";  //always connected to the database MSCS_IDEC_PRD
            
            //Establish database connection using SqlConnection()
            //need to add the package reference manually in the .csproj file it is causing issue when compiling
            SqlConnection sql = new SqlConnection(
                "user id=" + SQLUser + ";" +
                "password=" + SQLPassword + ";" +
                "server=" + SQLIP + ";" +
                "connection timeout=30"
            );

            string pids = getProductIDs(sql);
            imgSQL += pids + ") ORDER BY ProductID ASC";
            docSQL += pids + ") ORDER BY ProductID ASC";
            catSQL += pids + ") ORDER BY ProductID ASC";
              
            //send SQL command to the database SqlCommand(SQLcmd, SQLConnection)
            SqlCommand cmdDoc = new SqlCommand(docSQL, sql);
            SqlCommand cmdImg = new SqlCommand(imgSQL, sql);
            SqlCommand cmdCat = new SqlCommand(catSQL, sql);

            File.Create(docCSV).Close(); //create or overwrite \export\Document List.csv
            File.Create(imgCSV).Close(); //create or overwrite \export\Image List.csv
            File.Create("filelist.csv").Close(); // create or overwrite filelist.csv

            sql.Open(); //open database connection
            
            //SqlDataReader - A way of reading forwared-only stream of rows from SQL database
            SqlDataReader rdr = cmdDoc.ExecuteReader(); 

            //StreamWriter - Writing characters to a stream 
            StreamWriter sw = new StreamWriter("filelist.csv"); //writing data to filelist.csv
            StreamWriter swDoc = new StreamWriter(docCSV); //writing data to Document List.csv
            StreamWriter swImg = new StreamWriter(imgCSV); //writing data to Image List.csv
            StreamWriter swCat = new StreamWriter(catCSV); //writing data to Product Attributes.csv

            swDoc.WriteLine("ProductID,Filepath,Display Name");
            swImg.WriteLine("ProductID, FilePath");
            swCat.WriteLine("Column Names");

            while(rdr.Read()) //SQL data reader rdr
            {
                sw.WriteLine(rdr["FileLocation"].ToString()); //??
                swDoc.WriteLine(rdr["ProductId"].ToString() + ',' + rdr["FileLocation"].ToString() + ',' + rdr["DisplayName"].ToString()); //?
            }
            sw.Close();
            swImg.Close();
            rdr.Close();

            rdr = cmdCat.ExecuteReader();
            while(rdr.Read())
            {
                // write columns for Catalog Products here.
            }
            swCat.Close();
            rdr.Close();

            sql.Close();
        } 

        public static string getProductIDs(SqlConnection conn)
        {
            string output = "";
            if(!File.Exists("ProductIDs.csv"))  //if ProductIDs.csv doesn't exists?? File.Exists() return true if the file exists
                generateProductIDs(conn);
            
            string[] PIDs = File.ReadAllLines("ProductIDs.csv");
            foreach(string s in PIDs)
                output+= '\'' + s + "',";

            return output.Substring(0, output.Length-1);
            
        }

        private static void generateProductIDs(SqlConnection conn)
        {
            SqlCommand cmd = new SqlCommand(
                "SELECT ProductID FROM [MSCS_IDEC_PRD].[dbo].[IDEC_USA_CatalogProducts]" +
                "JOIN MSCS_IDEC_PRD.dbo.NAV_Item ON NavisionItemNumber = MSCS_IDEC_PRD.dbo.NAV_Item.No " +
                "WHERE ShowOnCommerceSite = 1 AND ShowOnPrimarySite = 1 AND Blocked = 0 AND Sellable = 1 " +
                "AND Item_Category_Code in ('1', '2A', '2B')",
                conn);
            conn.Open(); //open the database connection
            SqlDataReader rdr = cmd.ExecuteReader(); //Sql Data Reader class rdr

            List<string> IDs = new List<string>(); //System.Collections.Generic;

            while(rdr.Read())
                IDs.Add(rdr["ProductID"].ToString());
            conn.Close(); //close the database connection

            File.Create("ProductIDs.csv").Close();
            StreamWriter sw = new StreamWriter("ProductIDs.csv");
            foreach(string s in IDs)
                sw.WriteLine(s);
            
            sw.Close();
        }

        static void Main(string[] args)
        {
            //testing code
            //Console.WriteLine("Hello World!");
            /*
            String path = @"C:\ProductIDs.csv";         
            if(File.Exists(path))
                 Console.WriteLine("File ProductIDs exists...");
            else
                 Console.WriteLine("Can't find the target!");

            Console.WriteLine("The return value is:{0}", File.Exists(path));
            Console.ReadLine();
            */
            Console.WriteLine("Please make sure the .csv you wish to use is in the same directory as this .exe file." +
                "\n\nThe title of the file should be 'ProductIDs.csv' and should contain nothig other than product IDs." + 
                "\n\nIf no product IDs have been specified, this program will generate one for sellable products (see documentation)." +
                "\n\nPress any key to continue.");
            Console.ReadKey();
            makeFileList();

            Console.WriteLine("\n\n");

            File.Create("log.txt").Close();
            StreamWriter sw = new StreamWriter("log.txt");

            string imgDirectory = @"E:\Websites\idec\productimages\Originals\";
            string docDirectory = @"E:\Websites\idec\Productdocuments\";
            
            string[] paths = File.ReadAllLines("filelist.csv");  //This line needs to be checked if it exceed the array index
            int errors = 0;

            //Remove duplicate lines.
            for(int outdex = 0; outdex < paths.Length; outdex++)
                for(int index = outdex+1 ; index < paths.Length; index++)
                    if(paths[outdex].Length > 0 && paths[outdex].Equals(paths[index]))
                       paths[outdex] = "";

            for(int i=0 ; i < paths.Length; i++)
            {
                if(paths[i].Length > 0)
                {
                    paths[i] = paths[i].Replace('/', '\\');
                    try
                    {
                        string destination = "export\\images\\" + paths[i].Remove(paths[i].LastIndexOf('\\'));
                        if(!System.IO.Directory.Exists(destination))
                           System.IO.Directory.CreateDirectory(destination);
                        //if(!File.Exists(imgDirectory + paths[i]))
                        File.Copy(imgDirectory + paths[i], @"export\images" + paths[i]);
                    }
                    catch(Exception e)
                    {
                        try
                        {
                            string destination = "export\\documents\\" + paths[i].Remove(paths[i].LastIndexOf('\\'));
                            if(!System.IO.Directory.Exists(destination) && destination.Contains(@"\english\"));
                               System.IO.Directory.CreateDirectory(destination);
                            //if (!File.Exists(docDirectory + paths[i]))
                            File.Copy(docDirectory + paths[i], @"export\documents" + paths[i]);
                        }
                        catch(Exception ex)
                        {
                            //Console.WriteLine(ex.Message); errors++;
                            sw.WriteLine(ex.Message + '\n');
                            errors++;
                        }
                    }
                    if (i%10 == 0 && i>0)
                        Console.WriteLine(i.ToString() + " of" + paths.Length.ToString() + "copied.");
                }
            }
            sw.Close();
            deleteEmptyDirs();
            Console.WriteLine("Copying completed. There were " +  errors.ToString() + " errors.\n\nPlease see log.txt for more information. \n\n Press any key to exit.");
            Console.ReadKey();
     
        }

        private static void deleteEmptyDirs(string startLocation = "..\\")
        {
            foreach(var directory in Directory.GetDirectories(startLocation))
            {
                deleteEmptyDirs(directory);
                if(Directory.GetFiles(directory).Length == 0 && Directory.GetDirectories(directory).Length == 0)
                   Directory.Delete(directory, false);
            }
        }

    }
}
