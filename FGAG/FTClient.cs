using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenPop.Pop3;
using OpenPop.Mime;
using System.IO;
using System.Net.Mail;

namespace FGAG
{
    public class FTClient
    {

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

