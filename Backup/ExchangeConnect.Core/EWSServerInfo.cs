using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExchangeConnect;
using System.Configuration;
using System.Net;
using Microsoft.Exchange.WebServices.Data;


namespace ExchangeConnect.Core
{
    public static class EWSServerInfo
    {
        public static string EWSPath { get; private set; }
        public static string EWSUser { get; private set; }
        public static string EWSPassword { get; private set; }
        public static string EWSDomain { get; private set; }
        public static string EWSAutoDiscoverURL { get; private set; }

        static EWSServerInfo()
        {
            EWSPath = ConfigurationSettings.AppSettings["EWSPATH"];    
            EWSUser = ConfigurationSettings.AppSettings["EWSUSER"];
            EWSPassword = ConfigurationSettings.AppSettings["EWSPWD"];
            EWSDomain = ConfigurationSettings.AppSettings["EWSDOMAIN"];
            EWSAutoDiscoverURL = ConfigurationSettings.AppSettings["EWSAUTODISCOVERURL"];
        }

        public static ExchangeService  GetExchangeProxy()
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Credentials = new NetworkCredential(EWSUser, EWSPassword, EWSDomain);
            service.AutodiscoverUrl(EWSAutoDiscoverURL);
            return service;
        }

        public static ExchangeService GetExchangeProxy(string user, string password, string domain, string autoDiscoverURL)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
            service.Credentials = new NetworkCredential(user, password, domain);
            service.AutodiscoverUrl(autoDiscoverURL);
            return service;
        }

    }
}
