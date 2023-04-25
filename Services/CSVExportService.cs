using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;

namespace Core.Services
{
    public class CSVExportService : ICSVExportService
    {
        private readonly IUmbracoContextFactory _context;
        private readonly IWillService _willService;
        private readonly IReferralCodeReservationService _referralCodeReservationService;
        private readonly IMemberService _memberService;
        private readonly ILogger _logger;

        public CSVExportService(ILogger logger, IWillService willService, IReferralCodeReservationService referralCodeReservationService, IUmbracoContextFactory context, IMemberService memberService)
        {
            _logger = logger;
            _willService = willService;
            _referralCodeReservationService = referralCodeReservationService;
            _context = context;
            _memberService = memberService;
        }

        /// <summary>
        /// Parses a list of objects and returns a csv file with all the correct header information and property values. 
        /// </summary>
        /// <param name="dList">The list to format</param>
        /// <param name="FileName">The name of the csv file (a timestamp is added)</param>
        public void GenericExport(object[] dList, string FileName, bool SpaceCapitals = false)
        {
            string data = GenerateCSVData(dList, SpaceCapitals);

            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.ContentType = "text/csv";
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" + FileName + DateTime.UtcNow.ToString("dd-MM-yyyy") + ".csv");
            HttpContext.Current.Response.Write(data);
            HttpContext.Current.Response.End();

        }

        /// <summary>
        /// Generates a CSV data string for the given params
        /// </summary>
        /// <param name="dList"></param>
        /// <param name="SpaceCapitals"></param>
        /// <returns></returns>
        public string GenerateCSVData(object[] dList, bool SpaceCapitals = false)
        {
            if (!dList.Any())
            { return ""; }

            var input = dList.First();

            Type dType = input.GetType();

            string HeaderText = "";

            //loop through the properties and format the header of the csv file
            //Any item will do, they all have the same properties.
            int limit = dType.GetProperties().Count();
            int count = 0;
            foreach (var prop in dType.GetProperties())
            {
                if (SpaceCapitals)
                { HeaderText += AddSpacesToSentence(prop.Name, true); }
                else
                { HeaderText += prop.Name; }

                //If the first letters of a csv is 'ID', excel assumes you're trying to open an SYLK file.
                //To avoid formatting errors, enclose an ID in quotes (which formats normally)
                if (HeaderText.ToLower() == "id")
                { HeaderText = "\"" + HeaderText + "\""; }

                count++;

                if (count != limit)
                { HeaderText += ", "; }

            }

            var sbOutput = new StringBuilder();

            sbOutput.Append(HeaderText + "\r\n");

            //Go through each item in list
            foreach (object item in dList)
            {
                Type itmType = item.GetType();

                //Go through each property of each item
                foreach (PropertyInfo pi in itmType.GetProperties())
                {
                    object propValue = pi.GetValue(item, null);

                    var piType = pi.PropertyType;
                    if (propValue == null)
                    {
                        sbOutput.Append(FormatLine("NULL"));
                    }
                    else
                    {
                        //Format certain values certain ways.
                        if (piType.IsEquivalentTo(typeof(DateTime)))
                        {
                            sbOutput.Append(FormatLine(((DateTime)propValue).ToString("dd.MM.yy")));
                        }
                        else if (piType.IsEquivalentTo(typeof(decimal)))
                        {
                            sbOutput.Append(FormatLine(((decimal)propValue).ToString("N2")));
                        }
                        else
                        {
                            sbOutput.Append(FormatLine(propValue.ToString()));
                        }
                    }
                }

                sbOutput.Append("\r\n");
            }

            return sbOutput.ToString();
        }

        public static string FormatLine(string intput)
        {
            return "\"" + intput.Replace(",", "").Replace(System.Environment.NewLine, "").Replace("\t", "").Replace("\n", "").Replace("\r", "") + "\"" + ",";
        }

        //credit http://stackoverflow.com/a/272929
        private static string AddSpacesToSentence(string text, bool preserveAcronyms)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;
            StringBuilder newText = new StringBuilder(text.Length * 2);
            newText.Append(text[0]);
            for (int i = 1; i < text.Length; i++)
            {
                if (char.IsUpper(text[i]))
                    if ((text[i - 1] != ' ' && !char.IsUpper(text[i - 1])) ||
                        (preserveAcronyms && char.IsUpper(text[i - 1]) &&
                         i < text.Length - 1 && !char.IsUpper(text[i + 1])))
                        newText.Append(' ');
                newText.Append(text[i]);
            }
            return newText.ToString();
        }

    }

}
