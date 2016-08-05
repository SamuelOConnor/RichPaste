using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Documents;
using System.Windows.Forms;
using Microsoft.Office.Core;
using System.Threading;
using Microsoft.Office.Interop.OneNote;
using Application = Microsoft.Office.Interop.OneNote.Application;

namespace RichConsole
{
    [Serializable()]
    public class Program
    {
        [STAThread]
        private static string ConvertRtfToXaml(string rtfText)
        {
            string result = null;

            Exception threadEx = null;

            //Open new Thread
            Thread staThread = new Thread(
                delegate ()
                {
                try
                {
                    var richTextBox = new System.Windows.Controls.RichTextBox();

                    if (string.IsNullOrEmpty(rtfText))
                    result = ""; 
                   

                    TextRange textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);

                    //Create a MemoryStream of the Rtf content 

                    using (var rtfMemoryStream = new MemoryStream())
                    {
                        using (var rtfStreamWriter = new StreamWriter(rtfMemoryStream))
                        {
                            rtfStreamWriter.Write(rtfText);
                            rtfStreamWriter.Flush();
                            rtfMemoryStream.Seek(0, SeekOrigin.Begin);

                            //Load the MemoryStream into TextRange ranging from start to end of RichTextBox. 
                            textRange.Load(rtfMemoryStream, DataFormats.Rtf);
                        }
                    }


                    using (var rtfMemoryStream = new MemoryStream())
                    {

                        textRange = new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd);
                        textRange.Save(rtfMemoryStream, "Xaml");
                        rtfMemoryStream.Seek(0, SeekOrigin.Begin);
                        using (var rtfStreamReader = new StreamReader(rtfMemoryStream))
                        {
                                result = rtfStreamReader.ReadToEnd();
                        }

                    }
                    }

                    catch (Exception ex)
                    {
                        threadEx = ex;
                    }
                });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            return result;
        }
        

        public void OnHierarchyChange(string bstrActivePageID)
        {
        }

        public void OnNavigate()
        {
        }

 /*   < This [STAThread] is essential to access the clipboard    >            
  *     (If you are only using Class Libaries then this isn't enough and you   
  *      have to specifically open a new STAThread and have the Clipboard.Get          
  *      methods in there.)                                                          */
        [STAThread]
        public static void Main()
        {
            
            //Create New Instance of OneNote Application
            var oneNote = new Application();


            //Get current active page (see commented for current NoteBook/Section)

            /* string thisNoteBook = oneNote.Windows.CurrentWindow.CurrentNotebookId;
               string thisSection  = oneNote.Windows.CurrentWindow.CurrentSectionId;    */
            string thisPage = oneNote.Windows.CurrentWindow.CurrentPageId;


            //Get the content of the page
            string xmlPage;
            oneNote.GetPageContent(thisPage, out xmlPage);


            //Declarations for paste
            string returnStringText = null;
            string Pasteresult = null;


            //Set Namespace from our page
            string ns = xmlPage.Substring(xmlPage.IndexOf("xmlns:") + 6, xmlPage.IndexOf("=\"http://") - (xmlPage.IndexOf("xmlns:") + 6)) + @":";


            //Find the start and end of the page so that we can insert 
            //a new section in between for our pasted content
            int start = xmlPage.LastIndexOf("</"+ns+"Page>");
            string startstr = xmlPage.Substring(0, start);
            string endstr = xmlPage.Substring(start, xmlPage.Length - start);


            //New section for our content 
            //(this is essential when the page in OneNote is currently empty)
            string outline =
             "<" + ns + "Outline >" +
               "<" + ns + "Position x=\"35.0\" y=\"60.0\"/>" + //This puts it just under title
               "<" + ns + "Size width=\"750.75\" height=\"13.50\" isSetByUser=\"true\"/>" + //Makes the box nice and big
                 "<" + ns + "OEChildren>" +
                   "<" + ns + "OE>" +
                     "<" + ns + "T>"+
                         "<![CDATA[]]>"+
                     "</" + ns + "T>" +
                 "</" + ns + "OE>" +
               "</" + ns + "OEChildren>" +
             "</" + ns + "Outline>";


            //Form the new page and replace our original page
            string newpage = startstr + outline + endstr;
            xmlPage = newpage;


/*         (Below I've specifically opened a new STAThread to show you the code, 
 *          if you are using a console project then you only need [STAThread])                                                                    */

            //Get Text in Clipboard using full STAThread

            //Declarations to get out of the thread
            string PastedText = null;
            bool containsRTF = false;
            bool containsHTML = false;
            bool containsTEXT = false;
            Exception threadEx = null;

            //Open new Thread
            Thread staThread = new Thread(
                delegate ()
                {
                    try
                    {
                        //Query our clipboard
                        containsRTF = Clipboard.ContainsText(TextDataFormat.Rtf);
                        containsHTML = Clipboard.ContainsText(TextDataFormat.Html);
                        containsTEXT = Clipboard.ContainsText(TextDataFormat.Text);

                        //Get relevant data
                        if (containsRTF == true && containsHTML == false)
                        {
                            PastedText = Clipboard.GetText(TextDataFormat.Rtf);
                        }
                        else if (containsHTML == true)
                        {
                            PastedText = Clipboard.GetText(TextDataFormat.Html);
                        }
                        else if (containsTEXT == true)
                        {
                            PastedText = Clipboard.GetText(TextDataFormat.Text);
                        }

                    }

                    catch (Exception ex)
                    {
                        threadEx = ex;
                    }
                });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();


            //Does the clipboard contain Rich Text? (we put an extra && just to be sure it isn't html as they can sometimes get confused)
            if (containsRTF == true && containsHTML == false)
            {
                string XAML = ConvertRtfToXaml(PastedText);

                //Get HTML from Rich Text
                Pasteresult = ConvertXAMLtoHTML.convert(XAML);
            }

            //Does the clipboard contain HTML?
            else if (containsHTML == true)
            {

                //Get HTML from Clipboard
                returnStringText = PastedText;

                //Get rid of everything before <body> or the first <span>
                int pFrom = returnStringText.IndexOf("<BODY");
                if (pFrom == -1) pFrom = returnStringText.IndexOf("<BODY");
                if (pFrom == -1) pFrom = returnStringText.IndexOf("<SPAN");
                if (pFrom == -1) pFrom = returnStringText.IndexOf("<span");
                if (pFrom == -1) pFrom = returnStringText.IndexOf("<body");

                //Get rid of everything after </body> or the last </span>
                int pTo = returnStringText.LastIndexOf("</BODY>");
                if (pTo == -1) pTo = returnStringText.LastIndexOf("</body>");
                if (pTo == -1) pTo = returnStringText.LastIndexOf("</SPAN>");
                if (pTo == -1) pTo = returnStringText.LastIndexOf("</span>");

                Pasteresult = returnStringText.Substring(pFrom, pTo - pFrom + 7);
            }

            //Does the clipboard contain Plain Text?
            else if (containsTEXT == true)
            {
                //Get the Text
                Pasteresult = PastedText;

                //Replace any "<" or ">" to ensure that it isn't read as HTML
                Pasteresult = Pasteresult.Replace("<", "&lt;").Replace(">", "&gt;");

                //Skip the block of HTML tag replacements
                goto Skip;
            }

            //Remove all un-supported HTML tags (leaving pretty much just <span>'s)
            //Also make sure the HTML tags that remain are all lower case or onenote will complain
            Pasteresult = Pasteresult.Replace("<body", "<span").Replace("</body", "</span").Replace("<BODY", "<span").Replace("</BODY", "</span");
            Pasteresult = Pasteresult.Replace("<div", "<span").Replace("</div", "</span").Replace("</LI", "</li");
            Pasteresult = Pasteresult.Replace("<DIV", "<span").Replace("</DIV", "</span").Replace("<LI", "<li");
            Pasteresult = Pasteresult.Replace("<SPAN", "<span").Replace("</SPAN", "</span").Replace("</DIV", "</span");
            Pasteresult = Pasteresult.Replace("STYLE=", "style=").Replace("<UL", "<ul").Replace("</UL", "</ul");
            Pasteresult = Pasteresult.Replace("<!--EndFragment-->", "").Replace("<!--StartFragment-->", "");
            Pasteresult = Pasteresult.Replace("class=MsoNormal", "").Replace("mso-fareast-language:EN-US", "");
            Pasteresult = Pasteresult.Replace("<o:p>", "").Replace("</o:p>", "");
            Pasteresult = Pasteresult.Replace(System.Environment.NewLine, " ");

            //Do we have any paragraphs? If we do then we need to replace them with new XML CDATA sections
            int containsPs = Regex.Matches(Pasteresult, "<[p,P]").Count;

            //Get the ID of the last onenote block, then increment by 1
            int LastIDNumberPosition = xmlPage.LastIndexOf("{B0}");
            int LastIDNumberint = Convert.ToInt32(xmlPage.Substring(LastIDNumberPosition - 3, 2)) + 1;

            //For each <p> 
            while (containsPs > 0)
            {
                //This is the text we will replace the <p>s with, it adds a new block 
                string replace = " ]]></"+ns+"T></"+ns+"OE><"+ns+"OE><"+ns+"T><![CDATA[<span";  //Start current block (this starts a new line)

                //find the position of the <p>
                int pos = Pasteresult.IndexOf("<P", StringComparison.CurrentCultureIgnoreCase);

                //If there isn't any left then break
                if (pos < 0) { break; }
                else
                {
                    //Else replace the <p> with the new block
                    Pasteresult = Pasteresult.Substring(0, pos) + replace + Pasteresult.Substring(pos + "<P".Length);
                }

                //Increment numbers
                LastIDNumberint += 1;
                containsPs -= 1;
            }

            //Remove the orphaned </p> tags
            Pasteresult = Pasteresult.Replace("</P>", "").Replace("</p>", "");

        //Plain text will now re-join here
        Skip:


            //Get the position of the end of the last code block
            int PageContent = xmlPage.LastIndexOf("]]></"+ns+"T>");
            string startofPage = xmlPage.Substring(0, PageContent);
            string endofPage = xmlPage.Substring(PageContent, xmlPage.Length - PageContent);

            //Insert out pasted text into the current page
            string NewPage = startofPage + Pasteresult + endofPage;


            try
            {
                //Update the onenote page
                oneNote.UpdatePageContent(NewPage);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + Environment.NewLine + Environment.NewLine + "Try using Ctrl + P", "Error Message");
            }


            oneNote = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();

        }


    }
}
