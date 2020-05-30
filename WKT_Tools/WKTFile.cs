using System;
using System.Data.SqlTypes;
using System.IO;
using System.Text;
using Microsoft.SqlServer.Types;

namespace WKT_Tools
{
    public class WktFile
    {
        public string FileNameWithPath { get; }
        public string FileName { get; }
        public long FileSize { get; }
        public string WktText { get; }
        public int WktSize { get; }
        public string GeometryType { get; }
        public string IsValid { get; }

        public WktFile(string fileName)
        {
            FileNameWithPath = fileName;
            FileName = Path.GetFileName(fileName);
            FileSize = new FileInfo(fileName).Length;
            
            WktText = File.ReadAllText(fileName, Encoding.UTF8);
            
            WktSize = WktText.Length;

            File.WriteAllText(fileName, WktText, new UTF8Encoding(false));

            try
            {
                SqlGeometry geom = SqlGeometry.STGeomFromText(new SqlChars(WktText), 0);

                IsValid = geom.IsValidDetailed();

                if (geom.STIsValid())
                {
                    GeometryType = geom.STGeometryType().ToString();
                }
                else
                {
                    GeometryType = "Unknown";
                    Console.WriteLine(IsValid);
                }
            }
            catch (Exception e)
            {
                IsValid = e.Message;
                GeometryType = "Unknown";

                Console.WriteLine(e.Message);
            }

            if (WktText.Length > 32767)
            {
                WktText = WktText.Substring(0, 32767);
            }
            
        }
    }
}
