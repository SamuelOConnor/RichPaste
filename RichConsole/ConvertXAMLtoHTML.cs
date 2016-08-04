using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RichConsole
{
    public class ConvertXAMLtoHTML
    {
        public static string convert(string XAML)
        {
            string result = XAML.Substring(XAML.IndexOf("<Paragraph"));
            result = result.Replace("<Paragraph", "<p><span").Replace("</Paragraph", "</p");
            result = result.Replace("<Run>", "").Replace("</Run>", "");
            result = result.Replace("FontFamily=\"", "style=\"font-family:");
            result = result.Replace("\" FontSize=\"", "; font-size:");

            int StartMargin = result.IndexOf("\" Margin");
            string findEndOfMargin = result.Substring(StartMargin);
            int EndMargin = findEndOfMargin.IndexOf("\">")+1;
            
            //Remove Margin
            result = result.Substring(0, StartMargin) + ";\" " + result.Substring(StartMargin + EndMargin);

            result = result.Replace("<Span Foreground=\"#FF", "<span style=\"color:#");
            result = result.Replace("\">", ";\">");

            result = result.Replace("</Span>", "</span>");

            result = result.Replace("</Section>", "").Replace("</FlowDocument>","");


            return result;
        }
    }
}
