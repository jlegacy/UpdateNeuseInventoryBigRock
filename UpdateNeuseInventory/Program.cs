using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using LINQtoCSV;
using Newtonsoft.Json;
using SharpCompress.Archive;
using SharpCompress.Common;
using UpdateNeuseInventory.Properties;

namespace UpdateNeuseInventory
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            bool haveFile = false;

            // Data Source=sql-shark.cybersharks.net;Initial Catalog=nssnc;User ID=jlegacy;Password=***********
            // Data Source=JOSEPH\SQLEXPRESS;Initial Catalog=nssnc;Integrated Security=True

            string url =
                "https://retriever.bigrocksports.com/imagelib/getimages.php?User=h03110ds&Pwd=20152015&Action=getqty";
            WebRequest request = WebRequest.Create(url);
            request.ContentType = "application/json; charset=utf-8";
            string text;
            var response = (HttpWebResponse) request.GetResponse();

            using (var sr = new StreamReader(response.GetResponseStream()))
            {
                text = sr.ReadToEnd();
            }

            var deserializedResult = JsonConvert.DeserializeObject<Result>(text);

            if (deserializedResult.status.CompareTo("COMPLETE (200)") == 0)
            {
                haveFile = WriteFile(deserializedResult.description);

                if (haveFile)
                {
                    ReadCsv();
                }
            }

            Environment.Exit(0);
        }

        public static void ReadCsv()
        {
            double count = 0;
            double count2 = 0;
            String LeadingZeroItem = null;

            var inputFileDescription = new CsvFileDescription
            {
                SeparatorChar = ',',
                FirstLineHasColumnNames = true
            };

            var cc = new CsvContext();

            IEnumerable<SectionCsv> sections =
                cc.Read<SectionCsv>(Settings.Default.tempFileName, inputFileDescription);

            IEnumerable<SectionCsv> sectionsByName =
                from p in sections
                select p;

            foreach (SectionCsv item in sectionsByName)
            {
                count2++;

                Console.SetCursorPosition(0, 4); //move cursor
                Console.Write("CSV Item: " + count2 + " " + item.item);

                try
                {
                    double notUsed = Convert.ToDouble(item.upc);
                }
                catch (Exception)
                {
                    continue;
                }

                LeadingZeroItem = Convert.ToString(item.upc);

                double temp = Convert.ToDouble(item.upc);
                String NonLeadingZeroItem = Convert.ToString(temp);

                var cs = new ProductsDataContext();
                IQueryable<product> q =
                    from a in cs.GetTable<product>()
                    where a.pID.Equals(LeadingZeroItem) || a.pID.Equals(NonLeadingZeroItem)
                    select a;

                foreach (product a in q)
                {
                    buildData(a, item);
                    try
                    {
                        cs.SubmitChanges();
                        count = count + 1;
                        Console.SetCursorPosition(0, 0); //move cursor
                        Console.Write("Update Record: " + count);
                    }
                    catch (Exception e)
                    {
                        Console.Write(e.ToString());
                    }
                }
            }
        }

        private static void buildData(product myProduct, SectionCsv item)
        {
            myProduct.pInStock = item.qty;
        }

        public static Boolean WriteFile(String filename)
        {
            String toFileName = Settings.Default.zipDownLoadFile;

            var myWebClient = new WebClient();

            myWebClient.DownloadFile(filename, toFileName);

            IArchive archive = ArchiveFactory.Open(toFileName);
            foreach (IArchiveEntry entry in archive.Entries)
            {
                if (!entry.IsDirectory)
                {
                    entry.WriteToDirectory(Settings.Default.tempDir,
                        ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
                    return true;
                }
            }

            return false;
        }

        public class Result
        {
            public string description;
            public string status;
        }

        internal class SectionCsv
        {
            [CsvColumn(Name = "item #", FieldIndex = 1)]
            public string item { get; set; }

            [CsvColumn(Name = "upc #", FieldIndex = 2)]
            public string upc { get; set; }

            [CsvColumn(Name = "qty", FieldIndex = 3)]
            public int qty { get; set; }

            [CsvColumn(Name = "loc", FieldIndex = 4)]
            public string loc { get; set; }

            [CsvColumn(Name = "mastpk", FieldIndex = 5)]
            public string mastpk { get; set; }

            [CsvColumn(Name = "break", FieldIndex = 6)]
            public string breakPack { get; set; }
        }
    }
}