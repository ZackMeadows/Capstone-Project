using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OpenIdConnect;
using System.Configuration;
using System.Threading.Tasks;
using System.Security.Claims;
using Microsoft.Identity.Client;
using Capstone.TokenStorage;
using Microsoft.Graph;
using System.Net.Http.Headers;
using System.Data.OleDb;
using Capstone.Models;
using System.Globalization;
using System.IO;
using Capstone.Classes;
using System.Diagnostics;
using Newtonsoft.Json;

namespace Capstone.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            if (Request.IsAuthenticated)
            {
                string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;
                string userName = ClaimsPrincipal.Current.FindFirst(ClaimTypes.Email).Value;
                if (string.IsNullOrEmpty(userId))
                {
                    // Invalid principal, sign out
                    return RedirectToAction("SignOut");
                }

                // Since we cache tokens in the session, if the server restarts
                // but the browser still has a cached cookie, we may be
                // authenticated but not have a valid token cache. Check for this
                // and force signout.
                SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
                if (!tokenCache.HasData())
                {
                    // Cache is empty, sign out
                    return RedirectToAction("SignOut");
                }

                Session["USER"] = userName;
            }
            return View();
        }

        public async Task<ActionResult> Inbox()
        {
            if (Request.IsAuthenticated)
            {
                APIManager api = new APIManager(Session["USER"].ToString());
                string token = await api.GetAccessToken(HttpContext);
                if (string.IsNullOrEmpty(token))
                {
                    // If there's no token in the session, redirect to Home
                    return Redirect("/");
                }

                GraphServiceClient client = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization =
                                new AuthenticationHeaderValue("Bearer", token);

                            return Task.FromResult(0);
                        }));
                try
                {
                    var mailResults = await client.Me.MailFolders.Inbox.Messages.Request()
                                        .OrderBy("receivedDateTime DESC")
                                        .Select("subject,receivedDateTime,from")
                                        .Top(10)
                                        .GetAsync();

                    return View(mailResults.CurrentPage);
                }
                catch (ServiceException ex)
                {
                    return RedirectToAction("Error", "Home", new { message = "ERROR retrieving messages", debug = ex.Message });
                }
            }
            else { return RedirectToAction("SignOut", "Home", null); }
        }

        public async Task<ActionResult> Calendar()
        {
            if (Request.IsAuthenticated)
            {
                APIManager api = new APIManager(Session["USER"].ToString());
                string token = await api.GetAccessToken(HttpContext);
                if (string.IsNullOrEmpty(token))
                {
                    // If there's no token in the session, redirect to Home
                    return Redirect("/");
                }

                GraphServiceClient client = new GraphServiceClient(
                    new DelegateAuthenticationProvider(
                        (requestMessage) =>
                        {
                            requestMessage.Headers.Authorization =
                                new AuthenticationHeaderValue("Bearer", token);

                            return Task.FromResult(0);
                        }));

                try
                {
                    var eventResults = await client.Me.Events.Request()
                                        .OrderBy("start/dateTime DESC")
                                        .Select("subject,start,end")
                                        .Top(10)
                                        .GetAsync();

                    return View(eventResults.CurrentPage);
                }
                catch (ServiceException ex)
                {
                    return RedirectToAction("Error", "Home", new { message = "ERROR retrieving events", debug = ex.Message });
                }
            }
            else { return RedirectToAction("SignOut", "Home", null); }
        }
        
        public void SignIn()
        {
            if (!Request.IsAuthenticated)
            {
                HttpContext.GetOwinContext().Authentication.Challenge(
                    new AuthenticationProperties { RedirectUri = "/" },
                    OpenIdConnectAuthenticationDefaults.AuthenticationType);
            }
        }

        public void SignOut()
        {
            if (Request.IsAuthenticated)
            {
                string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

                if (!string.IsNullOrEmpty(userId))
                {
                    // Get the user's token cache and clear it
                    SessionTokenCache tokenCache = new SessionTokenCache(userId, HttpContext);
                    tokenCache.Clear();
                }
            }

            Session["USER"] = null;
            // Send an OpenID Connect sign-out request. 
            HttpContext.GetOwinContext().Authentication.SignOut(
                CookieAuthenticationDefaults.AuthenticationType);
            Response.Redirect("/");
        }
        public ActionResult Preferences()
        {
            if (Request.IsAuthenticated)
            {
                string user = Session["USER"].ToString();
                DBManager db = new DBManager();

                string uploadDir = db.GetUserUploadDirectory(user);
                string examDir = db.GetUserExamDirectory(user);
                string calGen = db.GetUserGenCalendars(user);

                ViewBag.uploadDir = uploadDir;
                ViewBag.examDir = examDir;
                ViewBag.calGen = calGen;

                return View();
            }
            else { return RedirectToAction("SignOut", "Home", null); }
        }
        [HttpPost]
        public ActionResult SavePreferences(string uploadDir, string examDir, bool calGen = false)
        {
            if (Request.IsAuthenticated)
            {
                string user = Session["USER"].ToString();
                DBManager db = new DBManager();

                db.SavePreferences(user, uploadDir, examDir, calGen.ToString());
                return Redirect("/");
            }
            else { return RedirectToAction("SignOut", "Home", null); }
        }
        public ActionResult History()
        {
            if (Request.IsAuthenticated)
            {
                DBManager db = new DBManager();
                ViewBag.HistoryList = db.GetHistoryList();
                return View();
            }
            else { return RedirectToAction("SignOut", "Home", null); }
        }
        public ActionResult DriveSelect()
        {
            if (Request.IsAuthenticated)
            {
                APIManager drive = new APIManager(Session["USER"].ToString());
                IDriveItemChildrenCollectionPage driveItems = Task.Run(() => drive.GetDriveItems(HttpContext)).Result;

                DBManager db = new DBManager();
                string dir = db.GetUserUploadDirectory(Session["USER"].ToString());

                ViewBag.Dir = dir;
                ViewBag.DriveList = driveItems;
                return View();
            }
            else { return RedirectToAction("SignOut", "Home", null); }
        }
    }
}