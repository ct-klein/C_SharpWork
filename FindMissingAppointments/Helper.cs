 using CommonLibraries.V2;
  using Microsoft.Xrm.Tooling.Connector;
  using System;
  using System.Net;
  using System.Threading;
  
 namespace Helper
  {
      public static class CrmServiceHelper
       {
           private static string crmUrl;
           private static string clientId;
           private static string clientSecret;
   
          public static void Initialize(string url, string id, string secret)
           {
               crmUrl = url;
               clientId = id;
               clientSecret = secret;
           }
   
          public static CrmServiceClient GetCrmServiceClient()
           {
               const int sleepCounterMax = 3;
               const int sleepMillSeconds = 5000;
               ServicePointManager.SecurityProtocol = SecurityProtocolType.SystemDefault;
   
              try
               {
                  ValidateConnectionComponents();
   
                  for (var sleepCounter = 1; sleepCounter <= sleepCounterMax; sleepCounter++)
                   {
                       try
                       {
                           var svc = AttemptToGetCRMService();
   
                          if (svc != null && svc.IsReady)
                           {
                               ConfigureServiceClient(svc);
                               return svc;
                           }
   
                          Console.WriteLine($"CrmServiceClient not ready .... Retry #{sleepCounter}");
                           Thread.Sleep(sleepMillSeconds);
                       }
                       catch (Exception ex)
                       {
                           Console.WriteLine(ex.Message);
                           throw;
                       }
                   }
   
                  Console.WriteLine($"CrmServiceClient exceeded {sleepCounterMax} retry attempts");
                   Console.WriteLine($"CrmServiceClient Connection Failed.");
                 throw new ArgumentException("CrmServiceClient connection attempts exceeded the max number of retry attempts. Connection has Failed.");
               }
               catch (ArgumentException ex)
               {
                   Console.WriteLine(ex.Message);
                   throw;
               }
           }
   
          private static void ValidateConnectionComponents()
           {
               if (string.IsNullOrEmpty(crmUrl) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret))
               {
                   Console.WriteLine($"Some CrmServiceClient Connection Components are Empty - " +
                                     $"Url : {crmUrl} - " +
                                     $"Client Id : {clientId} - " +
                                     $"Client Secret : {clientSecret}");
   
                  throw new ArgumentException("CrmServiceClient connection components cannot be null or empty.");
               }
           }
   
          private static CrmServiceClient AttemptToGetCRMService()
           {
               CrmServiceClient.MaxConnectionTimeout = new TimeSpan(0, 420, 0);
               var svc = CrmConnection.GetCRMService(crmUrl, clientId, clientSecret);
               return svc;
           }
   
          private static void ConfigureServiceClient(CrmServiceClient svc)
           {
               svc.MaxRetryCount = 3;
           }
       }
   }   