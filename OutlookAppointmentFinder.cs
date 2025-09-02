   using Azure.Identity;
   using Microsoft.Extensions.Logging;
   using System;
   using System.Net.Http;
   using System.Net.Http.Headers;
   using System.Text;
   using System.Text.Json;
   using System.Threading.Tasks;
   using Azure.Storage.Blobs;
   using Azure.Storage.Blobs.Models;
   using System.Globalization;
   using CsvHelper;
   using System.Text;
   using System.Formats.Asn1;
   using CommonCommunication;
   
   using System.Net.Mail;
    
   public class OutlookAppointmentFinder
   {
        private const string CsvHeader = 
            "FirstName,LastName,Email,StartDateTime,EndDateTime,ConfirmationNumber,Phone,CompanyName,AdditionalComments,Status";
 
        public static async Task Main(string[] args)
         {
          // Parse the input string as a DateTime
         DateTime utcTime = DateTime.Parse("5/2/2025 1:00:00 PM");
 
           // Format the UTC time as an ISO 8601 string
            string iso8601Time = utcTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
    
           // Example usage of FindAppointmentUIDAsync
           DateTime startTime = DateTime.Parse("2025-06-27 01:00:00");    //Time MUST be in this format
           DateTime endTime = DateTime.Parse("2025-06-27 23:30:00");      //
   
          // Replace with a valid user email
           string userEmail = "ct_klein@hotmail.com";
   
          string connectionString = "";
           DateTime targetDate = DateTime.Today.AddDays(-1); // Yesterday's date in UTC
   
          string outputDirectory = @"C:\Temp"; // Change as needed
   
          string[] containerNames = { "scheduled-appointments", "canceled-appointments" };
           string[] filePrefixes = { "New Advice Desk Appointments", "Canceled Advice Desk Appointments" };

          for (int i = 0; i < containerNames.Length; i++)
           {
               await DownloadCsvsAndCombineAsync(
                   connectionString,
                   containerNames[i],
                   targetDate,
                   outputDirectory,
                   filePrefixes[i]);
           }
   
  
          //var appointmentUIDs = await FindAppointmentUIDAsync(userEmail, startTime, endTime);
           //foreach (var uid in appointmentUIDs)
           //{
           //    await SetAppointmentToFreeAsync(userEmail, uid);
           //}
   
          //DateTime newEndTime = endTime.AddMinutes(15);
   
          //var appointmentUIDs = await FindAppointmentUIDAsync(userEmail, startTime, endTime);
           //foreach (var uid in appointmentUIDs)
           //{
           //    await UpdateAppointmentEndTimeAsync(userEmail, uid, newEndTime);
           //}
       }
       private static async Task<string> GetAccessTokenAsync()
       {
           var clientId = "";          //Azure:ClientID
           var tenantId = "";          //Azure:TenantID
           var clientSecret = "";  //Azure:ClientSecret
   
          var clientSecretCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
           var tokenRequestContext = new Azure.Core.TokenRequestContext(new[] { "https://graph.microsoft.com/.default" });
           var token = await clientSecretCredential.GetTokenAsync(tokenRequestContext);
   
          return token.Token;
       }
   
      public static async Task<List<string>> FindAppointmentUIDAsync(string userEmail, DateTime startTime, DateTime endTime)
       {
           var accessToken = await GetAccessTokenAsync();
   
          using var httpClient = new HttpClient();
           httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
   
          //var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/calendarView?startDateTime={startTime:O}&endDateTime={endTime:O}";
           var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/events?$filter=start/dateTime ge '{startTime}' and end/dateTime le '{endTime}'";
   
          var response = await httpClient.GetAsync(url);
           response.EnsureSuccessStatusCode();
   
          var jsonResponse = await response.Content.ReadAsStringAsync();
          var events = JsonDocument.Parse(jsonResponse).RootElement.GetProperty("value");
   
					var appointmentUIDs = new List<string>();
    
           //Event Name: Appointment Scheduler
    
           foreach (var calendarEvent in events.EnumerateArray())
            {
                var eventStart = DateTime.Parse(calendarEvent.GetProperty("start").GetProperty("dateTime").GetString());
                var eventEnd = DateTime.Parse(calendarEvent.GetProperty("end").GetProperty("dateTime").GetString());
   
               var bodyPreview = calendarEvent.GetProperty("bodyPreview").GetString();
               var uid = calendarEvent.GetProperty("id").GetString();
   
               Console.WriteLine($"Subject: {calendarEvent.GetProperty("subject").GetString()}");
   
               if (bodyPreview != null && bodyPreview.Contains("Work"))  //Make this a environmental setting in Azure
                {
                    appointmentUIDs.Add(uid);
                    Console.WriteLine($"Subject: {calendarEvent.GetProperty("subject").GetString()}");
               }
                else
                {
                   //Skip this appointment
                   continue;
               }
   
          }
   
          return appointmentUIDs;
       }
   
      public static async Task<bool> UpdateAppointmentEndTimeAsync(
           string userEmail,
           string appointmentUID,
           DateTime newEndTime)
       {
           var accessToken = await GetAccessTokenAsync();
   
          using var httpClient = new HttpClient();
           httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
   
          var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/events/{appointmentUID}";
   
          // Prepare the PATCH payload
           var updatePayload = new
           {
               end = new
               {
                   dateTime = newEndTime.ToString("o"), // ISO 8601 format
                   timeZone = "UTC"
               }
           };
   
          var content = new StringContent(JsonSerializer.Serialize(updatePayload), Encoding.UTF8, "application/json");
           var method = new HttpMethod("PATCH");
           var request = new HttpRequestMessage(method, url) { Content = content };
   
          var response = await httpClient.SendAsync(request);
   
          if (response.IsSuccessStatusCode)
           {
               Console.WriteLine("Appointment end time updated successfully.");
               return true;
           }
           else
           {
               var errorDetails = await response.Content.ReadAsStringAsync();
               Console.WriteLine($"Failed to update appointment. Status code: {response.StatusCode}, Error: {errorDetails}");
               return false;
           }
       }
   
  
      public static async Task DeleteAppointmentAsync(string userEmail, string appointmentUID)
       {
           var accessToken = await GetAccessTokenAsync();
   
          using var httpClient = new HttpClient();
           httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
   
          var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/events/{appointmentUID}";
  
          var response = await httpClient.DeleteAsync(url);
   
          if (response.IsSuccessStatusCode)
           {
               Console.WriteLine("Appointment deleted successfully.");
           }
           else
           {
               Console.WriteLine($"Failed to delete appointment. Status code: {response.StatusCode}");
               var errorDetails = await response.Content.ReadAsStringAsync();
               Console.WriteLine($"Error details: {errorDetails}");
           }
       }
   
      public static async Task SetAppointmentToFreeAsync(string userEmail, string appointmentUID)
       {
           Console.WriteLine($"Setting appointment with UID: {appointmentUID} to free time for user: {userEmail}");
   
          try
           {
               var accessToken = await GetAccessTokenAsync();
   
              using var httpClient = new HttpClient();
               httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
               httpClient.DefaultRequestHeaders.Add("Prefer", "outlook.suppress-notifications");
   
              var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/events/{appointmentUID}" ;
   
              // Update the event's showAs property to "free"
               var updateEvent = new
               {
                   showAs = "free"
               };
   
              var content = new StringContent(JsonSerializer.Serialize(updateEvent), Encoding.UTF8, "application/json");
               var response = await httpClient.PatchAsync(url, content);
   
              if (response.IsSuccessStatusCode)
               {
                   Console.WriteLine($"Successfully set appointment with UID: {appointmentUID} to free time.");
               }
               else
               {
                   var errorDetails = await response.Content.ReadAsStringAsync();
                   Console.WriteLine($"Failed to set appointment to free time. Status code: {response.StatusCode}, Error: {errorDetails}");
               }
           }
           catch (Exception ex)
           {
               Console.WriteLine($"Unexpected error while setting appointment to free time. Error: {ex.Message}");
           }
       }
   
      public static async Task DownloadCsvsAndCombineAsync(
           string connectionString,
           string containerName,
           DateTime targetDate,
          string outputDirectory,
           string outputFileNamePrefix)
       {
           var blobServiceClient = new BlobServiceClient(connectionString);
           var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
   
          // List blobs for the given day
           var blobs = new List<BlobItem>();
           await foreach (var blobItem in containerClient.GetBlobsAsync())
           {
               Console.WriteLine($"Blob Name: {blobItem.Name}, Last Modified: {blobItem.Properties.LastModified.Value.ToLocalTime().Date}");
               if (blobItem.Name.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) &&
                   blobItem.Properties.LastModified.HasValue &&
                   blobItem.Properties.LastModified.Value.ToLocalTime().Date == targetDate.Date)
               {
                   blobs.Add(blobItem);
               }
           }
   
          var outputFilePath = Path.Combine(
               outputDirectory,
               $"{outputFileNamePrefix} {targetDate:yyyy-MM-dd}.csv"
           );
   
          // Explicitly delete the file if it exists
           if (File.Exists(outputFilePath))
           {
               File.Delete(outputFilePath);
           }
   
          using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
           {
              bool headerWritten = false;
   
              if (blobs.Count == 0)
               {
                   Console.WriteLine("No CSV files found for the given date.");
                   await writer.WriteLineAsync(CsvHeader);
               }
               else
               {
                   foreach (var blob in blobs)
                   {
                       var blobClient = containerClient.GetBlobClient(blob.Name);
                       using (var stream = await blobClient.OpenReadAsync())
                       using (var reader = new StreamReader(stream))
                       {
                           string? headerLine = await reader.ReadLineAsync();
   
                          if (!headerWritten && headerLine != null)
                           {
                               await writer.WriteLineAsync(headerLine);
                               headerWritten = true;
                           }
   
                          // Write the rest of the lines (data rows)
                           while (!reader.EndOfStream)
                           {
                               var line = await reader.ReadLineAsync();
                               if (!string.IsNullOrWhiteSpace(line))
                               {
                                   await writer.WriteLineAsync(line);
                               }
                           }
                       }
                   }
               }
          }
           Console.WriteLine($"Combined CSV saved to: {outputFilePath}");
       }
   } 
