using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenPop.Pop3;
using OpenPop.Mime;
using System.IO;
using System.Net.Mail;

using Microsoft.Office.Interop.Excel;

namespace FGAG
{
    public class FTClient
    { 
        Microsoft.Office.Interop.Excel.Application _excel = new Application();
        public  void GetAttachments(string hostname, int port, bool useSsl, string username, string password, string savelocation)
        {
        

            // The client disconnects from the server when being disposed
            using (Pop3Client client = new Pop3Client())
            {
                // Connect to the server
                client.Connect(hostname, port, useSsl);
               
                // Authenticate ourselves towards the server
                client.Authenticate(username, password);

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();
                _excel.Visible = true;
                _excel.Workbooks.Open(@"C:\Users\TEMP.UPOFFICE.007\Documents\book1.xlsm", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Workbook w = _excel.ActiveWorkbook;
                _excel.Run("boss");
             

                // We want to download all messages
                List<Message> allMessages = new List<Message>(messageCount);

                // Messages are numbered in the interval: [1, messageCount]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number
                for (int i = messageCount; i > 0; i--)
                {
                    var mess = client.GetMessage(i);
                
                    if (mess.Headers.Subject.ToLower().Contains("current"))
                    {
                        var atts = mess.FindAllAttachments();
                        foreach (MessagePart attachment in atts)
                        {
                            string filePath = Path.Combine(@"C:\Attachment", attachment.FileName);
                         
                                FileStream Stream = new FileStream(filePath, FileMode.Create);
                                BinaryWriter BinaryStream = new BinaryWriter(Stream);
                                BinaryStream.Write(attachment.Body);
                                BinaryStream.Close();
                            
                        }
                        client.DeleteMessage(i);
                    }

           
           
                }

                // Now return the fetched messages
            
            }
        }

        
        }
    }

