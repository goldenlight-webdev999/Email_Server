using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace TestEmail
{
    class Program
    {
        enum EventCode : int
        {
            CommangArgsMissing = 50,
            ErrorSendingEmail = 200,
            BadEmailAddress = 250,
            UnknownError = 900,
        }


        /// <summary>
        /// Args[0] = SMTP Server
        /// Args[1] = SMTP Port
        /// Args[2] = SMTP Requires Auth (1 for true , 0 for false)
        /// Args[3] = SMTP Requires SSL (1 for true , 0 for false)
        /// Args[4] = From Email
        /// Args[5] = From Display Name
        /// Args[6] = SMTP UserName
        /// Args[7] = SMTP Password
        /// </summary>
        /// <param name="args"></param>

        static void Main(string[] args)
        {
            int nExitCode = 0;

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

                if (args.Length < 6 && !bAuthRequired)
                {
                    LogError("One or more command arguments missing\n\n Auth Required: " + bAuthRequired.ToString() + "\n\n args count: " + args.Length.ToString(), EventLogEntryType.Error, EventCode.CommangArgsMissing);
                    Environment.Exit((int)EventCode.CommangArgsMissing);
                }
                else if (args.Length < 8 && bAuthRequired)
                {
                    LogError("One or more command arguments missing\n\n Auth Required: " + bAuthRequired.ToString() + "\n\n args count: " + args.Length.ToString(), EventLogEntryType.Error, EventCode.CommangArgsMissing);
                    Environment.Exit((int)EventCode.CommangArgsMissing);
                }
            }


            try
            {
                string sFromEmail = (string)args.GetValue(4).ToString().Replace("~", " ").Trim();
                SendMail(sFromEmail, sFromEmail, "TerraTRASH", "TerraTRASH Test Message", "This is an e-mail message sent automatically by TerraTRASH while testing the settings for your account", args);
            }
            catch (Exception ex)
            {
                LogError("Main: " + ex.Message, EventLogEntryType.Error, EventCode.UnknownError);
                nExitCode = (int)EventCode.UnknownError;
            }


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
                    sUserName = (string)args.GetValue(6).ToString().Replace("~", " ").Trim();
                    sPassword = (string)args.GetValue(7).ToString().Replace("~", " ").Trim();
                    basicCredential = new NetworkCredential(sUserName, sPassword);
                }
                else
                {
                    sServer = (string)args.GetValue(0).ToString().Replace("~", " ").Trim();
                    int.TryParse((string)args.GetValue(1).ToString().Replace(" ", "").Trim(), out nPort);
                    basicCredential = new NetworkCredential();
                }

                SmtpServer.Host = sServer;
                SmtpServer.UseDefaultCredentials = false;
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

        private static void SendMail(string sToEmail, string sFromEmail, string sFromDisplayName, string sSubject, string sBody, string[] Args)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient oSmtpServer = SMTPServer(Args);

                mail.From = new MailAddress(sFromEmail, sFromDisplayName);
                IsValidEmail(sToEmail);
                mail.To.Add(sToEmail);
                mail.Subject = sSubject;
                mail.Body = sBody;
                mail.IsBodyHtml = true;
                ServicePointManager.ServerCertificateValidationCallback = delegate(object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                oSmtpServer.Send(mail);
            }
            catch (SmtpFailedRecipientException ex)
            {
                LogError(ex.Message, EventLogEntryType.Error, EventCode.ErrorSendingEmail);
                Environment.Exit((int)EventCode.ErrorSendingEmail);
            }
            catch (SmtpException ex)
            {
                LogError(ex.Message, EventLogEntryType.Error, EventCode.ErrorSendingEmail);
                Environment.Exit((int)EventCode.ErrorSendingEmail);
            }
            catch (Exception ex)
            {
                LogError(ex.Message, EventLogEntryType.Error, EventCode.ErrorSendingEmail);
                Environment.Exit((int)EventCode.ErrorSendingEmail);
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

        private static void IsValidEmail(string sEmailAddress)
        {
            try
            {
                MailMessage oTestMail = new MailMessage();
                try
                {
                    oTestMail.To.Add(sEmailAddress);
                }
                catch (Exception)
                {
                }
                oTestMail.Dispose();
            }
            catch (Exception ex)
            {
                LogError(ex.Message, EventLogEntryType.Error, EventCode.BadEmailAddress);
                Environment.Exit((int)EventCode.BadEmailAddress);
            }

        }

    }
}
