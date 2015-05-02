using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections;
using System.Linq;
using AE.Net.Mail;

using AE.Net.Mail.Imap;
namespace RainMail_CLI
{
    class Program
    {
        
        public static string result = "";
        public static string usage = "USAGE:\n\n" +
                        "RainMail-CLI.exe <parameter>\nhost=<Required:  IMAP-Server Adress>\nuser=<Required:    IMAP-Login Username>\npassword=<Optional: IMAP-Login Password if set>\nport=<Optional: IMAP-Server port>\nssl=1 (If server needs ssl connection)";
        public static bool unreadCountOnly = false;
        public static bool useSSL = false;

        public static int unreadMailCount = 0;
        public static string unread_mail_1_sender = "";
        public static string unread_mail_1_subject = "";
        public static string unread_mail_1_date = "";

        public static string unread_mail_2_sender = "";
        public static string unread_mail_2_subject = "";
        public static string unread_mail_2_date = "";

        public static string unread_mail_3_sender = "";
        public static string unread_mail_3_subject = "";
        public static string unread_mail_3_date = "";

        public static string unread_mail_4_sender = "";
        public static string unread_mail_4_subject = "";
        public static string unread_mail_4_date = "";

        public static string unread_mail_5_sender = "";
        public static string unread_mail_5_subject = "";
        public static string unread_mail_5_date = "";

        static void Main(string[] args)
        {
            File.Delete("imap_mail_results.ini");
            if (args.Length < 1){ExitWithError();}
            else 
            {
                var arguments = new Dictionary<string, string>();
                foreach (string argument in args)
                {
                    string[] splitted = argument.Split('=');
                    if (splitted.Length == 2){arguments[splitted[0]] = splitted[1];}
                }

                string imapHost = arguments.ContainsKey("host") ? arguments["host"] : "";
                if (imapHost == "") { ExitWithError(); }

                string imapUser = arguments.ContainsKey("user") ? arguments["user"] : "";
                if (imapUser == "") { ExitWithError(); }

                string imapPassword = arguments.ContainsKey("password") ? arguments["password"] : "";
                int imapPort = arguments.ContainsKey("port") ? int.Parse(arguments["port"]) : 143;
                int imapSSL = arguments.ContainsKey("ssl") ? int.Parse(arguments["ssl"]) : 0;
                useSSL = imapSSL==1 ? true : false;
                unreadCountOnly = arguments.ContainsKey("count") ? true : false;
                
                Console.WriteLine("These IMAP-Connection Parameters will be used:\n");
                Console.WriteLine("Server-Host:\t"+imapHost);
                Console.WriteLine("Server-Port:\t" + imapPort);
                Console.WriteLine("Server-User:\t" + imapUser);
                Console.WriteLine("Server-Password:\t" + imapPassword);
                Console.WriteLine("SSL:\t\t"+useSSL.ToString());

                using (ImapClient ic = new ImapClient(imapHost, imapUser, imapPassword, AuthMethods.Login, imapPort, useSSL))
                {
                    ic.SelectMailbox("INBOX");
                    Console.Clear();
                    Console.WriteLine(ic.IsConnected.ToString());
                    //Console.WriteLine(ic.ListMailboxes("", ""));
                    // Note that you must specify that headersonly = false
                    // when using GetMesssages().
                    
                    MailMessage[] mm = ic.GetMessages(0, int.Parse((ic.GetMessageCount()-1).ToString()), true).Where(m => !m.Flags.HasFlag(Flags.Seen)).ToArray();
                    unreadMailCount = mm.Length;
                    int tmpCounter = 0;
                    foreach (MailMessage m in mm)
                    {
                        tmpCounter++;
                        switch (tmpCounter)
                        {
                            case 1:
                                unread_mail_1_sender = m.From.DisplayName;
                                unread_mail_1_subject = m.Subject;
                                unread_mail_1_date = m.Date.ToShortDateString() + " " + m.Date.ToShortTimeString().ToString();
                                break;
                            case 2:
                                unread_mail_2_sender = m.From.DisplayName;
                                unread_mail_2_subject = m.Subject;
                                unread_mail_2_date = m.Date.ToShortDateString() + " " + m.Date.ToShortTimeString().ToString();
                                break;
                            case 3:
                                unread_mail_3_sender = m.From.DisplayName;
                                unread_mail_3_subject = m.Subject;
                                unread_mail_3_date = m.Date.ToShortDateString() + " " + m.Date.ToShortTimeString().ToString();
                                break;
                            case 4:
                                unread_mail_4_sender = m.From.DisplayName;
                                unread_mail_4_subject = m.Subject;
                                unread_mail_4_date = m.Date.ToShortDateString() + " " + m.Date.ToShortTimeString().ToString();
                                break;
                            case 5:
                                unread_mail_5_sender = m.From.DisplayName;
                                unread_mail_5_subject = m.Subject;
                                unread_mail_5_date = m.Date.ToShortDateString() + " " + m.Date.ToShortTimeString().ToString();
                                break;
                        }
                        
                    }
                    WriteResultFile();
                    ic.Logout();
                }


            }            

        }

        static void ExitWithError()
        {
            result = "No arguments given!\n" + usage;
            Console.WriteLine(result);
            Environment.Exit(1);
        }


        static void WriteResultFile()
        {
            string content = "";
            content +=  "[Variables]"+Environment.NewLine;
            content += "imap_mail_unreadcount="+unreadMailCount+Environment.NewLine;
            
            content += "imap_mail_1_from=" + unread_mail_1_sender + Environment.NewLine;
            content += "imap_mail_1_subject=" + unread_mail_1_subject + Environment.NewLine;
            content += "imap_mail_1_date=" + unread_mail_1_date + Environment.NewLine;

            content += "imap_mail_2_from=" + unread_mail_2_sender + Environment.NewLine;
            content += "imap_mail_2_subject=" + unread_mail_2_subject + Environment.NewLine;
            content += "imap_mail_2_date=" + unread_mail_2_date + Environment.NewLine;

            content += "imap_mail_3_from=" + unread_mail_3_sender + Environment.NewLine;
            content += "imap_mail_3_subject=" + unread_mail_3_subject + Environment.NewLine;
            content += "imap_mail_3_date=" + unread_mail_3_date + Environment.NewLine;

            content += "imap_mail_4_from=" + unread_mail_4_sender + Environment.NewLine;
            content += "imap_mail_4_subject=" + unread_mail_4_subject + Environment.NewLine;
            content += "imap_mail_4_date=" + unread_mail_4_date + Environment.NewLine;

            content += "imap_mail_5_from=" + unread_mail_5_sender + Environment.NewLine;
            content += "imap_mail_5_subject=" + unread_mail_5_subject + Environment.NewLine;
            content += "imap_mail_5_date=" + unread_mail_5_date + Environment.NewLine;
            
            File.WriteAllText("imap_mail_results.ini", content);
        }

 
    } 

    }

