 using Azure.Storage.Blobs;
  using CommonLibraries.V2;
  using Microsoft.Crm.Sdk.Messages;
  using Microsoft.Xrm.Sdk.Query;
  using System;
  using System.Collections.Generic;
  using System.Configuration;
  using System.Linq;
  using System.IO;
   using Helper;
   
  namespace FindMissingAppointments
   {
       class Program
       {
           private static readonly string crmUrl = ConfigurationManager.AppSettings["CRMURL"];
           private static readonly string connectionString = ConfigurationManager.AppSettings["Azure:BlobStorageConnectionString"];
           private static readonly string keyVaultIdentifier = ConfigurationManager.AppSettings["Azure:KeyVaultIdentifier"];
           private static readonly string clientId = CrmConnection.GetKeyVaultSecret(keyVaultIdentifier, ConfigurationManager.AppSettings["Azure:CRMID"]);
           private static readonly string clientSecret = CrmConnection.GetKeyVaultSecret(keyVaultIdentifier, ConfigurationManager.AppSettings["Azure:CRMSEC"]);
   
          private const string BlobContainerName = "scheduled-appointments";
           private const string OutputDirectory = @"C:\temp\";
   
          private static void Main()
           {
               DateTime targetDate = DateTime.Today.AddDays(-0); // Yesterday's date in UTC
   
              try
               {
                   CrmServiceHelper.Initialize(crmUrl, clientId, clientSecret);
   
                  using (var service = CrmServiceHelper.GetCrmServiceClient())
                   {
                       if (service == null)
                       {
                           Console.WriteLine("Failed to Established Connection!!!");
                           return;
                       }
   
                      var whoAmIResponse = service.Execute(new WhoAmIRequest()) as WhoAmIResponse;
                       if (whoAmIResponse == null || whoAmIResponse.UserId == Guid.Empty)
                       {
                           Console.WriteLine("Failed to retrieve CRM user ID.");
                           return;
                       }
   
                      var blobServiceClient = new BlobServiceClient(connectionString);
                       var containerClient = blobServiceClient.GetBlobContainerClient(BlobContainerName);
   
                      var confirmationNumbers = new List<string>();
                       foreach (var blobItem in containerClient.GetBlobs())
                       {
                           //Console.WriteLine($"Blob Name: {blobItem.Name}, Last Modified: {blobItem.Properties.LastModified.Value.ToLocalTime().Date}");
                           if (blobItem.Name.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) &&
                               blobItem.Properties.LastModified.HasValue &&
                               blobItem.Properties.LastModified.Value.ToLocalTime().Date == targetDate.Date)
                          {
                               string name = blobItem.Name.Substring(0, blobItem.Name.Length - 4); // Removes last 4 characters
                               confirmationNumbers.Add(name);
                           }
                       }
   
                      if (!confirmationNumbers.Any())
                       {
                           Console.WriteLine("No confirmation numbers found for the target date.");
                           return;
                       }
   
                      var query = new QueryExpression("appointment")
                       {
                           ColumnSet = new ColumnSet("activityid", "subject", "ct_confirmationnumber")};
   
                      query.Criteria.AddCondition("ct_confirmationnumber", ConditionOperator.In, confirmationNumbers.ToArray());
   
                      var appointments = service.RetrieveMultiple(query);
   
                      var foundConfirmationNumbers = new HashSet<string>(
                           appointments.Entities.Select(a => a.GetAttributeValue<string>("ct_confirmationnumber"))
                               .Where(cn => !string.IsNullOrEmpty(cn))
                       );
   
                      var notFound = confirmationNumbers.Except(foundConfirmationNumbers).ToList();
   
                      if (!Directory.Exists(OutputDirectory))
                       {
                           Directory.CreateDirectory(OutputDirectory);
                       }
   
                      string fileName = Path.Combine(OutputDirectory, $"{targetDate:yyyy-MM-dd}.csv");
   
                      if (notFound.Any())
                       {
                           File.WriteAllLines(fileName, notFound);
                           Console.WriteLine($"Missing confirmation numbers written to {fileName}");
                       }
                       else
                       {
                           Console.WriteLine("No missing confirmation numbers found. No file was created.");
                       }
                   }
               }
               catch (Exception ex)
               {
                   Console.WriteLine($"Exception caught - {ex}");
               }
           }
       }
   }