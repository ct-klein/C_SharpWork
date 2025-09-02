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
 121                   //Skip this appointment
 122                   continue;
 123               }
 124   
125           }
 126   
127           return appointmentUIDs;
 128       }
 129   
130       public static async Task<bool> UpdateAppointmentEndTimeAsync(
 131           string userEmail,
 132           string appointmentUID,
 133           DateTime newEndTime)
 134       {
 135           var accessToken = await GetAccessTokenAsync();
 136   
137           using var httpClient = new HttpClient();
 138           httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
 139   
140           var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/events/{appointmentUID}";
 141   
142           // Prepare the PATCH payload
 143           var updatePayload = new
 144           {
 145               end = new
 146               {
 147                   dateTime = newEndTime.ToString("o"), // ISO 8601 format
 148                   timeZone = "UTC"
 149               }
 150           };
 151   
152           var content = new StringContent(JsonSerializer.Serialize(updatePayload), Encoding.UTF8, "application/json");
 153           var method = new HttpMethod("PATCH");
 154           var request = new HttpRequestMessage(method, url) { Content = content };
 155   
156           var response = await httpClient.SendAsync(request);
 157   
158           if (response.IsSuccessStatusCode)
 159           {
 160               Console.WriteLine("Appointment end time updated successfully.");
 161               return true;
 162           }
 163           else
 164           {
 165               var errorDetails = await response.Content.ReadAsStringAsync();
 166               Console.WriteLine($"Failed to update appointment. Status code: {response.StatusCode}, Error: {errorDetails}");
 167               return false;
 168           }
 169       }
 170   
171   
172       public static async Task DeleteAppointmentAsync(string userEmail, string appointmentUID)
 173       {
 174           var accessToken = await GetAccessTokenAsync();
 175   
176           using var httpClient = new HttpClient();
 177           httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
 178   
179           var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/events/{appointmentUID}";
180   
181           var response = await httpClient.DeleteAsync(url);
 182   
183           if (response.IsSuccessStatusCode)
 184           {
 185               Console.WriteLine("Appointment deleted successfully.");
 186           }
 187           else
 188           {
 189               Console.WriteLine($"Failed to delete appointment. Status code: {response.StatusCode}");
 190               var errorDetails = await response.Content.ReadAsStringAsync();
 191               Console.WriteLine($"Error details: {errorDetails}");
 192           }
 193       }
 194   
195       public static async Task SetAppointmentToFreeAsync(string userEmail, string appointmentUID)
 196       {
 197           Console.WriteLine($"Setting appointment with UID: {appointmentUID} to free time for user: {userEmail}");
 198   
199           try
 200           {
 201               var accessToken = await GetAccessTokenAsync();
 202   
203               using var httpClient = new HttpClient();
 204               httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
 205               httpClient.DefaultRequestHeaders.Add("Prefer", "outlook.suppress-notifications");
 206   
207               var url = $"https://graph.microsoft.com/v1.0/users/{userEmail}/events/{appointmentUID}" ;
 208   
209               // Update the event's showAs property to "free"
 210               var updateEvent = new
 211               {
 212                   showAs = "free"
 213               };
 214   
215               var content = new StringContent(JsonSerializer.Serialize(updateEvent), Encoding.UTF8, "application/json");
 216               var response = await httpClient.PatchAsync(url, content);
 217   
218               if (response.IsSuccessStatusCode)
 219               {
 220                   Console.WriteLine($"Successfully set appointment with UID: {appointmentUID} to free time.");
 221               }
 222               else
 223               {
 224                   var errorDetails = await response.Content.ReadAsStringAsync();
 225                   Console.WriteLine($"Failed to set appointment to free time. Status code: {response.StatusCode}, Error: {errorDetails}");
 226               }
 227           }
 228           catch (Exception ex)
 229           {
 230               Console.WriteLine($"Unexpected error while setting appointment to free time. Error: {ex.Message}");
 231           }
 232       }
 233   
234       public static async Task DownloadCsvsAndCombineAsync(
 235           string connectionString,
 236           string containerName,
 237           DateTime targetDate,
238           string outputDirectory,
 239           string outputFileNamePrefix)
 240       {
 241           var blobServiceClient = new BlobServiceClient(connectionString);
 242           var containerClient = blobServiceClient.GetBlobContainerClient(containerName);
 243   
244           // List blobs for the given day
 245           var blobs = new List<BlobItem>();
 246           await foreach (var blobItem in containerClient.GetBlobsAsync())
 247           {
 248               Console.WriteLine($"Blob Name: {blobItem.Name}, Last Modified: {blobItem.Properties.LastModified.Value.ToLocalTime().Date}");
 249               if (blobItem.Name.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) &&
 250                   blobItem.Properties.LastModified.HasValue &&
 251                   blobItem.Properties.LastModified.Value.ToLocalTime().Date == targetDate.Date)
 252               {
 253                   blobs.Add(blobItem);
 254               }
 255           }
 256   
257           var outputFilePath = Path.Combine(
 258               outputDirectory,
 259               $"{outputFileNamePrefix} {targetDate:yyyy-MM-dd}.csv"
 260           );
 261   
262           // Explicitly delete the file if it exists
 263           if (File.Exists(outputFilePath))
 264           {
 265               File.Delete(outputFilePath);
 266           }
 267   
268           using (var writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
 269           {
 270               bool headerWritten = false;
 271   
272               if (blobs.Count == 0)
 273               {
 274                   Console.WriteLine("No CSV files found for the given date.");
 275                   await writer.WriteLineAsync(CsvHeader);
 276               }
 277               else
 278               {
 279                   foreach (var blob in blobs)
 280                   {
 281                       var blobClient = containerClient.GetBlobClient(blob.Name);
 282                       using (var stream = await blobClient.OpenReadAsync())
 283                       using (var reader = new StreamReader(stream))
 284                       {
 285                           string? headerLine = await reader.ReadLineAsync();
 286   
287                           if (!headerWritten && headerLine != null)
 288                           {
 289                               await writer.WriteLineAsync(headerLine);
 290                               headerWritten = true;
 291                           }
 292   
293                           // Write the rest of the lines (data rows)
 294                           while (!reader.EndOfStream)
 295                           {
 296                               var line = await reader.ReadLineAsync();
 297                               if (!string.IsNullOrWhiteSpace(line))
 298                               {
 299                                   await writer.WriteLineAsync(line);
 300                               }
 301                           }
 302                       }
 303                   }
 304               }
305           }
 306           Console.WriteLine($"Combined CSV saved to: {outputFilePath}");
 307       }
 308   }
 309   
