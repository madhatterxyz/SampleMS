/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Resources;
using System.IO;
using Microsoft.SharePoint.Client;
using System.Security;
using GraphResources;
using static GraphResources.Event;

namespace Microsoft_Graph_REST_ASPNET_Connect.Models
{

    // This sample shows how to:
    //    - Get the current user's email address
    //    - Get the current user's profile photo
    //    - Attach the photo as a file attachment to an email message
    //    - Upload the photo to the user's root drive
    //    - Get a sharing link for the file and add it to the message
    //    - Send the email
    public class GraphService
    {

        public async Task<UserInfo> getMe(string accessToken)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me";
            string queryParameter = "?$select=businessPhones, displayName, givenName, jobTitle, mail, mobilePhone, officeLocation" +
                ", preferredLanguage, surname, userPrincipalName, id, skills";
            UserInfo me = new UserInfo();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                  

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            me = JsonConvert.DeserializeObject<UserInfo>(stringResult);
                        }
                        return me;
                    }
                }
            }
        }

        public async Task<ListMessages> getMessages(string accessToken)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me/messages";
            string queryParameter = "";
            ListMessages messages = new ListMessages();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    


                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            messages = JsonConvert.DeserializeObject<ListMessages>(stringResult);
                        }
                        return messages;
                    }
                }
            }
        }

        // Get the current user's email address from their profile.
        public async Task<string> GetMyEmailAddress(string accessToken)
        {

            // Get the current user. 
            // The app only needs the user's email address, so select the mail and userPrincipalName properties.
            // If the mail property isn't defined, userPrincipalName should map to the email for all account types. 
            string endpoint = "http://functionapp20180321080621.azurewebsites.net/api/function1";
            string queryParameter = $"?token={accessToken}";
            UserInfo me = new UserInfo();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    //request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string text = await response.Content.ReadAsStringAsync();
                            //me.Address = !string.IsNullOrEmpty(json.GetValue("mail").ToString()) ? json.GetValue("mail").ToString() : json.GetValue("userPrincipalName").ToString();
                        }
                        return me.Address?.Trim();
                    }
                }
            }
        }

        // Get the current user's profile photo.
        public async Task<Stream> GetMyProfilePhoto(string accessToken)
        {

            // Get the profile photo of the current user (from the user's mailbox on Exchange Online). 
            // This operation in version 1.0 supports only a user's work or school mailboxes and not personal mailboxes. 
            string endpoint = "https://graph.microsoft.com/v1.0/me/photo/$value";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint))
                {
                    //request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    var response = await client.SendAsync(request);

                    // If successful, Microsoft Graph returns a 200 OK status code and the photo's binary data. If no photo exists, returns 404 Not Found.
                    if (response.IsSuccessStatusCode)
                    {
                        return await response.Content.ReadAsStreamAsync();
                    }
                    else
                    {
                        // If no photo exists, the sample uses a local file.
                        return System.IO.File.OpenRead(System.Web.Hosting.HostingEnvironment.MapPath("/Content/test.jpg"));
                    }
                }
            }
        }

        // Upload a file to OneDrive.
        // This call creates or updates the file.
        public async Task<GraphResources.FileInfo> UploadFile(string accessToken, Stream file)
        {

            // This operation only supports files up to 4MB in size.
            // To upload larger files, see `https://developer.microsoft.com/graph/docs/api-reference/v1.0/api/item_createUploadSession`.
            string endpoint = "https://graph.microsoft.com/v1.0/me/drive/root/children/mypic.jpg/content";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Put, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StreamContent(file);
                    request.Content.Headers.ContentType = new MediaTypeHeaderValue("image/jpg");
                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            return JsonConvert.DeserializeObject<GraphResources.FileInfo>(stringResult);
                        }
                        else return null;
                    }
                }
            }
        }

        
        // Send an email message from the current user.
        public async Task<string> SendEmail(string accessToken, MessageRequest email)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me/sendMail";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                    request.Content = new StringContent(JsonConvert.SerializeObject(email), Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return Resource.Graph_SendMail_Success_Result;
                        }
                        return response.ReasonPhrase;
                    }
                }
            }
        }

        // Create the email message.
        
        public async Task<SharePointSite> LoadSPSite(string accessToken, string sharePointDomain, string siteName)
        {

            // Get the specific SharePoint site. 
            string endpoint = string.Format("https://graph.microsoft.com/v1.0/sites/{0}:/sites/{1}/", sharePointDomain, siteName);
            string queryParameter = "";
            SharePointSite spSite = new SharePointSite();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            spSite = JsonConvert.DeserializeObject<SharePointSite>(stringResult);
                        }
                            return spSite;
                    }
                }
            }
        }

        public async Task<List<SharePointList>> LoadLists(string accessToken, SharePointSite spSite)
        {

            // Get all lists
            string endpoint = "https://graph.microsoft.com/v1.0/sites/" + spSite.id + "/lists/";
            string queryParameter = "";
            ResultLists lists = new ResultLists();
            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            lists = JsonConvert.DeserializeObject<ResultLists>(stringResult);
                        }
                    }
                    return new List<SharePointList>(lists.SharePointLists);
                }
            }
        }

        
        public async Task<List<Item>> GetItems(string accessToken, SharePointSite spSite, SharePointList spList)
        {

            // Get all items 
            string endpoint = "https://graph.microsoft.com/v1.0/sites/" + spSite.id + "/lists/" + spList.Id + "/items/";
            string queryParameter = "?expand=fields";
            ListItems ListItems = null;

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            ListItems = JsonConvert.DeserializeObject<ListItems>(stringResult);
                        }
                    }
                    return ListItems.Items;
                }
            }
        }

        public async Task<string> RegisterToEatAPI(string accessToken, SharePointSite spSite, SharePointList spList)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/sites/" + spSite.id + "/lists/" + spList.Id + "/items/";
            UserInfo me = await getMe(accessToken);

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
            
                    ItemCreated item = new ItemCreated()
                    {
                        fields = new FieldsCreated
                        {
                            Title = "Je veux manger",
                            DisplayName = me.displayName,
                            UPN = me.userPrincipalName
                        }
                    };
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    request.Content = new StringContent(JsonConvert.SerializeObject(item), Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return Resource.Graph_RegisterToEatAPI_Success_Result;
                        }
                        return response.ReasonPhrase;
                    }
                }
            }
        }
        
        
        public async Task<UserInfo> getUser(string accessToken, string upn)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/users/" + upn;
            string queryParameter = "?$select=businessPhones, displayName, givenName, jobTitle, mail, mobilePhone, officeLocation" +
                ", preferredLanguage, surname, userPrincipalName, id, skills";
            UserInfo user = new UserInfo();

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    //request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            user = JsonConvert.DeserializeObject<UserInfo>(stringResult);
                        }
                        return user;
                    }
                }
            }
        }

        public async Task<string> Matchpeople(string accessToken, SharePointSite spSite, SharePointList spList)
        {

            // Get all items 
            string endpoint = "https://graph.microsoft.com/v1.0/sites/" + spSite.id + "/lists/" + spList.Id + "/items/";
            string queryParameter = "?expand=fields";
            ListItems ListItems = null;

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Get, endpoint + queryParameter))
                {
                    request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            string stringResult = await response.Content.ReadAsStringAsync();
                            ListItems = JsonConvert.DeserializeObject<ListItems>(stringResult);
                        }
                    }

                    //return await SetupDateWithoutMySkill(accessToken, ListItems.Items);
                    //return await SetupDateWithMySkill(accessToken, ListItems.Items);
                    return await SetupDateAnyone(accessToken, ListItems.Items);
                }
            }
        }

        public async Task<string> SetupDateAnyone(string accessToken, List<Item> items)
        {
            string endpoint = "http://functionapp20180321080621.azurewebsites.net/api/function1";
            string queryParameter = $"?token={accessToken}";

            List<UserInfo> date = new List<UserInfo>();
            UserInfo me = await getMe(accessToken);
            date.Add(me);

            foreach (var item in items)
            {
                if (!item.fields.UPN.Equals(me.userPrincipalName))
                {
                    date.Add(await getUser(accessToken, item.fields.UPN));
                    using (var client = new HttpClient())
                    {
                        using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint + queryParameter))
                        {
                            request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                            //request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                            // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                            request.Headers.Add("SampleID", "aspnet-connect-rest-sample");

                            using (var response = await client.SendAsync(request))
                            {
                                if (response.IsSuccessStatusCode)
                                {
                                    string text = await response.Content.ReadAsStringAsync();
                                    //me.Address = !string.IsNullOrEmpty(json.GetValue("mail").ToString()) ? json.GetValue("mail").ToString() : json.GetValue("userPrincipalName").ToString();
                                }
                                return me.Address?.Trim();
                            }
                        }
                    }
                    return await BookInCalendar(accessToken, me, date[1]);
                }
            }
            return Resource.Graph_BookInCalendar_Failure_Result;

            
        }

        public async Task<string> SetupDateWithMySkill(string accessToken, List<Item> items)
        {
            List<UserInfo> date = new List<UserInfo>();
            UserInfo me = await getMe(accessToken);
            date.Add(me);

            foreach (var item in items)
            {
                UserInfo otherUser = await getUser(accessToken, item.fields.UPN);

                if (!otherUser.userPrincipalName.Equals(me.userPrincipalName))
                {
                    foreach(string skill in me.Skills)
                    {
                        if (otherUser.Skills.Contains(skill))
                        {
                            date.Add(otherUser);
                            return await BookInCalendar(accessToken, me, date[1]);

                        }
                    }
                }
            }
            return Resource.Graph_BookInCalendar_Failure_Result;
        }

        public async Task<string> SetupDateWithoutMySkill(string accessToken, List<Item> items)
        {
            List<UserInfo> date = new List<UserInfo>();
            UserInfo me = await getMe(accessToken);
            date.Add(me);
            UserInfo usertWithoutMySkill = null;

            foreach (var item in items)
            {
                UserInfo otherUser = await getUser(accessToken, item.fields.UPN);
                if (!otherUser.userPrincipalName.Equals(me.userPrincipalName))
                {
                    foreach (string skill in me.Skills)
                    {
                        if (otherUser.Skills.Contains(skill))
                        {
                            usertWithoutMySkill = null;
                            break;
                        }
                        else
                        {
                            usertWithoutMySkill = otherUser;
                        }
                    }
                }
            }
            if(usertWithoutMySkill != null)
            {
                date.Add(usertWithoutMySkill);
                return await BookInCalendar(accessToken, me, date[1]);
            }

            return Resource.Graph_BookInCalendar_Failure_Result;
        }


        public async Task<Boolean> CheckRegistration(string accessToken, SharePointSite spSite, SharePointList spList)
        {
            UserInfo me = await getMe(accessToken);
            List<Item> items = await GetItems(accessToken, spSite, spList);

            foreach (var item in items)
            {
                if (item.fields.UPN.Equals(me.userPrincipalName))
                {
                    return true;
                }
            }
            return false;
        }

        public async Task<string> BookInCalendar(string accessToken, UserInfo me, UserInfo user)
        {
            string endpoint = "https://graph.microsoft.com/v1.0/me/events";

            using (var client = new HttpClient())
            {
                using (var request = new HttpRequestMessage(HttpMethod.Post, endpoint))
                {
                    Event lunch = new Event();
                    lunch.subject = Resource.DefaultSubject;
                    lunch.body = new Body()
                    {
                        contentType = "HTML",
                        content = Resource.DefaultBody
                    };
                    lunch.start = new Start()
                    {
                        dateTime = DateTime.Today,
                        timeZone = "(UTC-04:00) Asuncion" // Resource.DefaultTimeZone
                    };
                    lunch.end = new End()
                    {
                        dateTime = lunch.start.dateTime.AddHours(2),
                        timeZone = "(UTC-04:00) Asuncion" //Resource.DefaultTimeZone
                    };
                    lunch.location = new Location() { displayName = Resource.DefaultLocation };
                    lunch.attendees = new List<Attendee>();
                    lunch.attendees.Add(new Attendee()
                    {
                        emailAddress = new EmailAddress() { address = user.userPrincipalName, name = user.displayName },
                        type = "required"
                    });
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    request.Content = new StringContent(JsonConvert.SerializeObject(lunch), Encoding.UTF8, "application/json");
                    using (var response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            return Resource.Graph_BookInCalendar_Success_Result;
                        }
                        return Resource.Graph_BookInCalendar_Failure_Result;
                    }
                }
            }
        }
        private static SecureString SecurePassword(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            securePassword.MakeReadOnly();
            return securePassword;

        }

    }
}
 