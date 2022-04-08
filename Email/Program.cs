using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Email
{
    class Program
    {
        enum EventCode : int
        {
            Success = 0,
            CommangArgsMissing = 50,
            DirectoryNotFound = 100,
            ErrorSendingEmail = 200,
            BadEmailAddress = 250,
            PdfFileNotFound = 300,
            UnknownError = 900,
        }


        /// <summary>
        /// Args[0] = SMTP Server
        /// Args[1] = SMTP Port
        /// Args[2] = SMTP Requires Auth (1 for true , 0 for false)
        /// Args[3] = SMTP Requires SSL (1 for true , 0 for false)
        /// Args[4] = From Email
        /// Args[5] = From Display Name
        /// Args[6] = PDF and mail file Location
        /// Args[7] = SMTP UserName
        /// Args[8] = SMTP Password
        /// 
        /// Email Text File - ToAddresses | Subject | Body |  absolute path to PDF file
        /// 
        /// </summary>
        /// <param name="args"></param>

        static void Main(string[] args)
        {
            if (args.Length < 3)
            {
                LogError("Required command arguments missing", EventLogEntryType.Error, EventCode.CommangArgsMissing);
                Environment.Exit((int)EventCode.CommangArgsMissing);
            }
            else
            {
                int nAuthReq;
                int.TryParse((string)args.GetValue(2).ToString().Replace(" ", "").Trim(), out nAuthReq);
                bool bAuthRequired = nAuthReq == 1;

                    if (args.Length < 7 && !bAuthRequired)
                {
                    LogError("One or more command arguments missing\n\n Auth Required: " + bAuthRequired.ToString() + "\n\n args count: " + args.Length.ToString(), EventLogEntryType.Error, EventCode.CommangArgsMissing);
                    Environment.Exit((int)EventCode.CommangArgsMissing);
                }
                else if (args.Length < 9 && bAuthRequired)
                {
                    LogError("One or more command arguments missing\n\n Auth Required: " + bAuthRequired.ToString() + "\n\n args count: " + args.Length.ToString(), EventLogEntryType.Error, EventCode.CommangArgsMissing);
                    Environment.Exit((int)EventCode.CommangArgsMissing);
                }
            }
            string sPDFFolderLocation = (string)args.GetValue(6).ToString().Replace("~", " ").Trim();
            string sDisplayName = (string)args.GetValue(5).ToString().Replace("~", " ").Trim();
            string sFromEmail = (string)args.GetValue(4).ToString().Replace("~", " ").Trim();
            StringBuilder sbLog = new StringBuilder();


            if (!System.IO.Directory.Exists(sPDFFolderLocation))
            {
                Environment.Exit((int)EventCode.DirectoryNotFound);
            }
            bool bErrors = false;
            ParseEmailsAndSend(sPDFFolderLocation, sFromEmail, sDisplayName, ref bErrors, args, ref sbLog);
            int nExitCode = 0;
            if (bErrors)
            {
                nExitCode = (int)EventCode.ErrorSendingEmail;
                LogError("ParseEmailsAndSend returned :" + bErrors.ToString(), EventLogEntryType.Error, EventCode.UnknownError);

            }
            else
            {
                nExitCode = (int)EventCode.Success;
            }
            //sbLog.AppendLine("</root>");
            System.IO.File.WriteAllText(sPDFFolderLocation + "\\Log.txt", sbLog.ToString());
            Environment.Exit(nExitCode);
        }

        private static SmtpClient SMTPServer(string[] args)
        {
            SmtpClient SmtpServer = new SmtpClient();
            try
            {
                string sServer = "";
                string sUserName = "";
                string sPassword = "";
                int nPort = 0;
                int nAuthReq = 0;
                int nSSLReq = 0;

                int.TryParse((string)args.GetValue(2).ToString().Replace(" ", "").Trim(), out nAuthReq);
                int.TryParse((string)args.GetValue(3).ToString().Replace(" ", "").Trim(), out nSSLReq);

                bool bAuthRequired = nAuthReq == 1;
                bool bSSLRequired = nSSLReq == 1;

                NetworkCredential basicCredential;
                if (bAuthRequired)
                {
                    sServer = (string)args.GetValue(0).ToString().Replace("~", " ").Trim();
                    int.TryParse((string)args.GetValue(1).ToString().Replace(" ", "").Trim(), out nPort);
                    sUserName = (string)args.GetValue(7).ToString().Replace("~", " ").Trim();
                    sPassword = (string)args.GetValue(8).ToString().Replace("~", " ").Trim();
                    basicCredential = new NetworkCredential(sUserName, sPassword);
                }
                else
                {
                    sServer = (string)args.GetValue(0).ToString().Replace("~", " ").Trim();
                    int.TryParse((string)args.GetValue(1).ToString().Replace(" ", "").Trim(), out nPort);
                    basicCredential = new NetworkCredential();
                }

                SmtpServer.Host = sServer;
                SmtpServer.UseDefaultCredentials = !bAuthRequired;
                SmtpServer.Credentials = basicCredential;
                SmtpServer.Port = nPort;
                SmtpServer.EnableSsl = bSSLRequired;
            }
            catch (Exception ex)
            {
                LogError(ex.ToString(), EventLogEntryType.Error, EventCode.UnknownError);
            }

            return SmtpServer;
        }

        private static void ParseEmailsAndSend(string sPDFFolderLocation, string sFromEmail, string sDisplayName, ref bool bErrors, string[] args, ref StringBuilder sbLog)
        {
            try
            {
                string sEmails = "" + System.IO.File.ReadAllText(sPDFFolderLocation + "\\Emails.txt");
                sEmails = sEmails.Replace("\"", "");
                sEmails = sEmails.Trim();

                string[] saEmails = sEmails.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                foreach (string sEmail in saEmails)
                {
                    if (sEmail == "")
                    {
                        continue;
                    }

                    string[] saEmail = sEmail.Split(new Char[] { '|' });
                    string sEmailAddresses = saEmail[0].Trim();
                    string sSubject = saEmail[1].Trim();
                    string sBody = saEmail[2].Trim();
                    string sPDFFilePath = saEmail[3];

                    if (File.Exists(sPDFFilePath))
                    {
                        if (!SendMail(sEmailAddresses, sPDFFilePath, sFromEmail, sDisplayName, sSubject, sBody, args, ref sbLog))
                        {
                            bErrors = true;
                            continue;
                        }
                    }
                    else
                    {
                        sbLog.Append("The file path \"" + sPDFFilePath + "\" was not found.\n");
                        LogError(sbLog.ToString(), EventLogEntryType.Error, EventCode.UnknownError);
                        bErrors = true;
                    }

                }
            }
            catch (Exception ex)
            {
                LogError(ex.ToString(), EventLogEntryType.Error, EventCode.UnknownError);
            }

        }

        private static bool SendMail(string sEmails, string sPDFPath, string sFromEmail, string sFromDisplayName, string sSubject, string sBody, string[] Args, ref StringBuilder sbLog)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient oSmtpServer = SMTPServer(Args);


                mail.From = new MailAddress(sFromEmail, sFromDisplayName);
                bool bBadEmails = false;
                SplitMultipleEmails(sEmails, ref mail, ref bBadEmails);
                if (bBadEmails)
                {
                    return false;
                }
                mail.Subject = sSubject;
                mail.Body = sBody;
                mail.IsBodyHtml = true;

                if (!System.IO.File.Exists(sPDFPath))
                {
                    LogError("PDF file not found @ " + sPDFPath, EventLogEntryType.Error, EventCode.PdfFileNotFound);
                    sbLog.AppendLine(sEmails);
                    sbLog.AppendLine("PDF file not found @ " + sPDFPath);
                    return false;
                }
                else
                {
                    string sFileName = Path.GetFileName(sPDFPath);
                    Attachment Attach = new Attachment(sPDFPath);
                    Attach.Name = sFileName;
                    mail.Attachments.Add(Attach);
                }

                ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                oSmtpServer.Send(mail);
                sbLog.AppendLine(sEmails);
                sbLog.AppendLine("Success");
                return true;
            }
            catch (SmtpFailedRecipientException ex)
            {
                sbLog.AppendLine(sEmails);
                sbLog.AppendLine(ex.Message);
                LogError(ex.Message, EventLogEntryType.Error, EventCode.ErrorSendingEmail);
            }
            catch (SmtpException ex)
            {
                sbLog.AppendLine(sEmails);
                sbLog.AppendLine(ex.Message);
                LogError(ex.Message, EventLogEntryType.Error, EventCode.ErrorSendingEmail);
            }
            catch (Exception ex)
            {
                sbLog.AppendLine(sEmails);
                sbLog.AppendLine(ex.Message);
                LogError(ex.Message, EventLogEntryType.Error, EventCode.ErrorSendingEmail);
            }
            return false;
        }

        private static bool IsValidEmail(string sEmailAddress)
        {
            try
            {
                MailMessage oTestMail = new MailMessage();
                bool bSuccess = false;
                try
                {
                    oTestMail.To.Add(sEmailAddress);
                    bSuccess = true;
                }
                catch (Exception)
                {
                    bSuccess = false;
                }
                oTestMail.Dispose();
                return bSuccess;
            }
            catch (Exception ex)
            {
                LogError(ex.ToString(), EventLogEntryType.Error, EventCode.UnknownError);
                return false;
            }

        }

        private static void SplitMultipleEmails(string sEmails, ref MailMessage oMail, ref bool bBadEmails)
        {
            string[] arEmails = null;
            string sValidatedEmails = sEmails.Replace(" ", "").Replace(Environment.NewLine, "").Replace(";", ",").Trim();
            arEmails = sValidatedEmails.Split(new Char[] { ',' });
            for (int i = 0; i <= arEmails.Length - 1; i++)
            {
                try
                {
                    string sEmail = arEmails[i].Trim();
                    if (sEmail.Length > 0)
                    {
                        if (IsValidEmail(sEmail))
                        {
                            oMail.To.Add(sEmail);
                        }
                        else
                        {
                            bBadEmails = true;
                            LogError("Bad Email address: " + sEmail, EventLogEntryType.Error, EventCode.BadEmailAddress);
                        }
                        
                    }

                }
                catch (Exception ex)
                {
                    LogError(ex.ToString(), EventLogEntryType.Error, EventCode.UnknownError);
                }
            }
        }

        private static void LogError(string sMessage, EventLogEntryType oType, EventCode oCode)
        {
            string sSource = "Email PDF";
            string sLog = "TerraVista";
            if (!EventLog.SourceExists(sSource))
            {
                EventLog.CreateEventSource(sSource, sLog);
            }
            EventLog.WriteEntry(sSource, sMessage, oType, (int)oCode);
        }
    }
}
