/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft_Graph_REST_ASPNET_Connect.Helpers;
using Microsoft_Graph_REST_ASPNET_Connect.Models;
using Resources;
using System;
using System.IO;
using System.Collections.Generic;
using GraphResources;
using Newtonsoft.Json;

namespace Microsoft_Graph_REST_ASPNET_Connect.Controllers
{
    public class HomeController : Controller
    {
        GraphService graphService = new GraphService();
        //Get a SharePoint domain
        string sharePointDomain = Resource.SharePointDomainDev;
        //Get a site name
        string siteName = Resource.SiteNameDev;


        public async Task<ActionResult> Index()
        {
            if (Request.IsAuthenticated)
            {
                try
                {

                    string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();
                    
                }
                catch(Exception e)
                {
                    if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                    return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
                }

            }
            return View("Test");
        }

        public async Task<ActionResult> GetMe()
        {
            try
            {

                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();
                var items = await graphService.getMe(accessToken);
                ViewBag.Values = items.ToString();
                ViewBag.API = "getme/";

                return View("Test");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        public async Task<ActionResult> SendMessage()
        {
            try
            {

                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();
                UserInfo addresse = await graphService.getMe(accessToken);
                Recipient recipient = new Recipient
                {
                    EmailAddress = addresse
                };
                List<Recipient> recipients = new List<Recipient>();
                recipients.Add(recipient);
                ItemBody body = new ItemBody
                {
                    Content = "Is a test"
                };

                Message message = new Message
                {
                    ToRecipients = recipients,
                    Body = body,
                    Subject = "Test API"
                };
                MessageRequest messageRequest = new MessageRequest
                {
                    Message = message,
                    SaveToSentItems = true

                };
                var items = await graphService.SendEmail(accessToken, messageRequest);
                ViewBag.API = "Postmessage/";
                ViewBag.Values = items.ToString();
                return View("Test");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        public async Task<ActionResult> GetMessages()
        {
            try
            {

                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                //
                var items = await graphService.getMessages(accessToken);
                ViewBag.API = "getmessage/";
                ViewBag.Values = items.ToString();
                return View("Test");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        [Authorize]
        // Get the current user's email address from their profile.
        public async Task<ActionResult> GetMyEmailAddress()
        {
            try
            {

                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                //Load SharePoint site
                SharePointSite spSite = await graphService.LoadSPSite(accessToken, sharePointDomain, siteName);
                spSite.Lists = await graphService.LoadLists(accessToken, spSite);

                SharePointList WantEatTodayList = null;
                //List of lists
                foreach (SharePointList list in spSite.Lists)
                {
                    if (list.Name.Equals(Resource.ListName))
                    {
                        WantEatTodayList = list;
                    }
                }
                List<Item> items = await graphService.GetItems(accessToken, spSite, WantEatTodayList);
                ViewBag.Items = items;

               
                // Get the current user's email address. 
//                ViewBag.Email = await graphService.GetListItem(accessToken);
                return View("Test");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        public async Task<ActionResult> About()
        {
            try
            {

                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                //
                var items = await graphService.GetMyEmailAddress(accessToken);
                ViewBag.Items = items;


                // Get the current user's email address. 
                //                ViewBag.Email = await graphService.GetListItem(accessToken);
                return View();
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        [Authorize]
        // Send mail on behalf of the current user.
        public async Task<ActionResult> RegisterToEat()
        {
            try
            {
                // Get an access token.
                string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

                SharePointSite spSite = await graphService.LoadSPSite(accessToken, sharePointDomain, siteName);
                spSite.Lists = await graphService.LoadLists(accessToken, spSite);

                SharePointList WantEatTodayList = null;
                //List of lists
                foreach (SharePointList list in spSite.Lists)
                {
                    if (list.Name.Equals(Resource.ListName))
                    {
                        WantEatTodayList = list;
                    }
                }

                await graphService.RegisterToEatAPI(accessToken, spSite, WantEatTodayList);

                // Build the email message. 
                return View("Test");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        public async Task<ActionResult> Matchme()
        {
            string accessToken = await SampleAuthProvider.Instance.GetUserAccessTokenAsync();

            SharePointSite spSite = await graphService.LoadSPSite(accessToken, sharePointDomain, siteName);
            spSite.Lists = await graphService.LoadLists(accessToken, spSite);

            SharePointList WantEatTodayList = null;
            //List of lists
            foreach (SharePointList list in spSite.Lists)
            {
                if (list.Name.Equals(Resource.ListName))
                {
                    WantEatTodayList = list;
                }
            }

            //Est-ce que je suis déjà enregistré dans la liste ?
            if (!await graphService.CheckRegistration(accessToken, spSite, WantEatTodayList))
            {
                //Si non je m'enregistre puis je tente un match
                string result = await graphService.RegisterToEatAPI(accessToken, spSite, WantEatTodayList);
                if (!result.Equals(Resource.Graph_RegisterToEatAPI_Success_Result))
                {
                    return RedirectToAction("Index", "Error", result);
                }
                ViewBag.Registration = true;

            }
            //Si oui, je vais tenter un match
            try
            {
                ViewBag.Message = await graphService.Matchpeople(accessToken, spSite, WantEatTodayList);

                return View("Test");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

        public ActionResult RemoveMe()
        {
            try
            {
                return View("Test");
            }
            catch (Exception e)
            {
                if (e.Message == Resource.Error_AuthChallengeNeeded) return new EmptyResult();
                return RedirectToAction("Index", "Error", new { message = Resource.Error_Message + Request.RawUrl + ": " + e.Message });
            }
        }

    }
}