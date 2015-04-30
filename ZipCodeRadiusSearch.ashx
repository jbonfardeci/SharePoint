<%@ Assembly Name="System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=B77A5C561934E089" %>
<%@ Assembly Name="System.Data.DataSetExtensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=B77A5C561934E089" %>
<%@ Assembly Name="System.Web.Extensions, Version=3.5.0.0, Culture=neutral, PublicKeyToken=31BF3856AD364E35" %>
<%@ Assembly Name="System.Xml.Linq, Version=3.5.0.0, Culture=neutral, PublicKeyToken=B77A5C561934E089"  %>
<%@ WebHandler Language="C#" Class="Webster.ZipCodeRadiusSearch" %>

using System;
using Microsoft.SharePoint;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Web;
using System.Linq;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace MySharePointNS
{
    public class ZipCodeRadiusSearch : IHttpHandler
    {
		//default list name
        private const string ZIPCODE_LISTNAME = "Zip Codes";

        public void ProcessRequest(HttpContext ctx)
        {
            //turn off caching           
            ctx.Response.CacheControl = "no-cache";
            ctx.Response.AddHeader("Pragma", "no-cache");
            ctx.Response.Expires = -1;
            ctx.Response.ContentType = "application/json";
            ctx.Response.ContentEncoding = UTF8Encoding.UTF8;
            
            HttpRequest r = ctx.Request;

            string zipCode = !string.IsNullOrEmpty(r["zip"]) 
                ? (string)ctx.Request["zip"] 
                : null;

            int radius = !string.IsNullOrEmpty(r["radius"]) 
                ? Convert.ToInt32( ctx.Request["radius"] )
                : 100;
                
            string listName = !string.IsNullOrEmpty(r["listName"]) 
                ? (string)ctx.Request["listName"]
                : ZIPCODE_LISTNAME;
            
            string json = null;

            if (zipCode != null)
                json = new JavaScriptSerializer().Serialize( GetZipCodeResults(listName, zipCode, radius) );

            ctx.Response.Write(json);
        }

		//JSON container for results
        public class Container
        {
            private List<KeyValuePair<string, double>> zipCodes;
            private string error;
            public List<KeyValuePair<string, double>> ZipCodes { get { return zipCodes; } set { zipCodes = value; } }
            public string Error { get { return error; } set { error = value; } }
        }

        public static Container GetZipCodeResults(listName, String zipCode, int radius)
        {
            Container container = new Container();
            container.ZipCodes = new List<KeyValuePair<string, double>>();

            try
            {
                using (SPWeb webContext = SPContext.Current.Web)
                {
                    SPSecurity.RunWithElevatedPrivileges(delegate()
                    {
                        using (SPSite site = new SPSite(webContext.Site.ID))
                        {
                            using (SPWeb web = site.OpenWeb(webContext.ID))
                            {

                                SPQuery q = new SPQuery();
                                q.Query = "<Where><Eq><FieldRef Name='Title'/><Value Type='Text'>" + zipCode + "</Value></Eq></Where>";
                                q.ViewFields = "<FieldRef Name=\"Title\" /><FieldRef Name=\"Latitude\" /><FieldRef Name=\"Longitude\" />";
                                q.RowLimit = 1;
                                SPList zipCodesList = web.Lists[listName];
                                SPListItemCollection startingZipCode = zipCodesList.GetItems(q);

                                if (startingZipCode.Count > 0)
                                {
                                    SPListItem startZip = startingZipCode[0];
                                    double startLatitude = (double)startZip["Latitude"];
                                    double startLongitude = (double)startZip["Longitude"];

                                    ZipCodeAssistant assistant = new ZipCodeAssistant(startLatitude, startLongitude, radius);

                                    SPQuery q2 = new SPQuery();
                                    q2.Query = String.Format(
                                        "<Where>" +
                                            "<And>" +
                                                "<And>" +
                                                    "<And>" +
                                                        "<Geq><FieldRef Name='Latitude'/><Value Type='Number'>{0}</Value></Geq>" +
                                                        "<Leq><FieldRef Name='Latitude'/><Value Type='Number'>{1}</Value></Leq>" +
                                                    "</And>" +
                                                    "<Geq><FieldRef Name='Longitude'/><Value Type='Number'>{2}</Value></Geq>" +
                                                "</And>" +
                                                "<Leq><FieldRef Name='Longitude'/><Value Type='Number'>{3}</Value></Leq>" +
                                            "</And>" +
                                        "</Where>" + 
                                        "<OrderBy>" + 
                                            "<FieldRef Name=\"Title\" Ascending=\"True\" />" + 
                                        "</OrderBy>", assistant.MinLatitude, assistant.MaxLatitude, assistant.MinLongitude, assistant.MaxLongitude);
                                    
                                    q2.ViewFields = "<FieldRef Name=\"Title\" /><FieldRef Name=\"Latitude\" /><FieldRef Name=\"Longitude\" />";
                                    q2.RowLimit = 100;

                                    SPListItemCollection zipCodeResults = zipCodesList.GetItems(q2);

                                    string tmpZip = null, thisZip = null;
                                    foreach (SPListItem zip in zipCodeResults)
                                    {
                                        thisZip = (string)zip["Title"];
                                        if (tmpZip != thisZip)
                                        {
                                            double d = Math.Round(assistant.GetDistance(startLatitude, startLongitude, (double)zip["Latitude"], (double)startZip["Longitude"]), 2);
                                            container.ZipCodes.Add(new KeyValuePair<string, double>(thisZip, d));
                                        }
                                        tmpZip = thisZip;
                                    }

                                    container.ZipCodes.Sort(CompareDistance);
                                    
                                }
                            }
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                container.Error = ex.ToString();
            }

            return container;
        }

        static int CompareDistance(KeyValuePair<string, double> a, KeyValuePair<string, double> b)
        {
            return a.Value.CompareTo(b.Value);
        }

        /// <summary>
        /// 
        /// </summary>
        public class ZipCodeAssistant
        {
            public const double EQUATOR_LAT_MILE = 69.172;
            private const int EARTH_RADIUS_MILES = 3963;
            private double maxLat, minLat, maxLong, minLong;

            public double MaxLatitude
            {
                get { return maxLat; }
            }

            public double MinLatitude
            {
                get { return minLat; }
            }

            public double MaxLongitude
            {
                get { return maxLong; }
            }

            public double MinLongitude
            {
                get { return minLong; }
            }

            public ZipCodeAssistant(double Latitude, double Longitude, int Miles)
            {
                maxLat = Latitude + Miles / EQUATOR_LAT_MILE;
                minLat = Latitude - (maxLat - Latitude);
                maxLong = Longitude + Miles / (Math.Cos(minLat * Math.PI / 180) * EQUATOR_LAT_MILE);
                minLong = Longitude - (maxLong - Longitude);
            }

            /// <summary>
            /// Haversine Distance Formula
            /// </summary>
            /// <param name="dblLat1"></param>
            /// <param name="dblLong1"></param>
            /// <param name="dblLat2"></param>
            /// <param name="dblLong2"></param>
            /// <returns></returns>
            public double GetDistance(double dblLat1, double dblLong1, double dblLat2, double dblLong2)
            {
                //convert degrees to radians
                dblLat1 = dblLat1 * Math.PI / 180;
                dblLong1 = dblLong1 * Math.PI / 180;
                dblLat2 = dblLat2 * Math.PI / 180;
                dblLong2 = dblLong2 * Math.PI / 180;

                double dist = 0;

                if (dblLat1 != dblLat2 || dblLong1 != dblLong2)
                {
                    //the two points are not the same
                    dist =
                        Math.Sin(dblLat1) * Math.Sin(dblLat2)
                        + Math.Cos(dblLat1) * Math.Cos(dblLat2)
                        * Math.Cos(dblLong2 - dblLong1);

                    dist =
                        EARTH_RADIUS_MILES
                        * (-1 * Math.Atan(dist / Math.Sqrt(1 - dist * dist)) + Math.PI / 2);
                }
                return dist;
            }
        }         
        
        public bool IsReusable
        {
            get
            {
                return false;
            }
        }

    }
}
