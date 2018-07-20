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
using Capstone.Controllers;
using System.Diagnostics;
using Newtonsoft.Json;
using Capstone.Classes.GeneratorClasses;

namespace Capstone.Classes
{
    public class APIManager
    {
        private string user;
        public APIManager(string user)
        {
            this.user = user;
        }

        public async Task<string> GetAccessToken(HttpContextBase httpContextBase)
        {
            string token = null;

            // Load the app config from web.config
            string appId = ConfigurationManager.AppSettings["ida:AppId"];
            string appPassword = ConfigurationManager.AppSettings["ida:AppPassword"];
            string redirectUri = ConfigurationManager.AppSettings["ida:RedirectUri"];
            string[] scopes = ConfigurationManager.AppSettings["ida:AppScopes"]
                .Replace(' ', ',').Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            // Get the current user's ID
            string userId = ClaimsPrincipal.Current.FindFirst(ClaimTypes.NameIdentifier).Value;

            if (!string.IsNullOrEmpty(userId))
            {
                // Get the user's token cache
                SessionTokenCache tokenCache = new SessionTokenCache(userId, httpContextBase);

                ConfidentialClientApplication cca = new ConfidentialClientApplication(
                    appId, redirectUri, new ClientCredential(appPassword), tokenCache.GetMsalCacheInstance(), null);

                // Call AcquireTokenSilentAsync, which will return the cached
                // access token if it has not expired. If it has expired, it will
                // handle using the refresh token to get a new one.
                AuthenticationResult result = await cca.AcquireTokenSilentAsync(scopes, cca.Users.FirstOrDefault());
                token = result.AccessToken;
            }
            return token;
        }
        public async Task<string> UploadSheet(HttpContextBase httpContextBase, string path, string saveName, bool examSheet = false)
        {
            // This whole process is returning a strange exception that doesn't seem to actually be doing anything.
            // File is uploaded properly without corruption and application continues normally.
            // Exception thrown: 'System.InvalidOperationException' in mscorlib.dll

            string token = await GetAccessToken(httpContextBase);
            if (string.IsNullOrEmpty(token))
                return null;

            GraphServiceClient client = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);

                        return Task.FromResult(0);
                    }));
            
            byte[] data = System.IO.File.ReadAllBytes(path);
            // Writeable stream from byte array for drive upload
            Stream stream = new MemoryStream(data);

            DBManager db = new DBManager();

            string dir = "";
            if (examSheet)
                dir = db.GetUserExamDirectory(user);
            else
                dir = db.GetUserUploadDirectory(user);

            Microsoft.Graph.DriveItem file = client.Me.Drive.Root.ItemWithPath(dir + saveName).Content.Request().PutAsync<DriveItem>(stream).Result;
            return file.WebUrl;
        }

        public async Task<IDriveItemChildrenCollectionPage> GetDriveItems(HttpContextBase httpContextBase)
        {
            // This whole process is returning a strange exception that doesn't seem to actually be doing anything.
            // File is uploaded properly without corruption and application continues normally.
            // Exception thrown: 'System.InvalidOperationException' in mscorlib.dll

            string token = await GetAccessToken(httpContextBase);
            if (string.IsNullOrEmpty(token))
                return null;

            GraphServiceClient client = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);

                        return Task.FromResult(0);
                    }));

            DBManager db = new DBManager();
            string dir = db.GetUserUploadDirectory(user);

            //Get all items in drive
            IDriveItemChildrenCollectionPage items = null;
            try
            {
                items = await client.Me.Drive.Root.ItemWithPath(dir).Children.Request().GetAsync();
            }
            catch(Exception e) { }
            return items;
        }
        
        public async Task<bool> DownloadSheet(HttpContextBase httpContextBase, string driveItemID, string path)
        {
            string token = await GetAccessToken(httpContextBase);
            if (string.IsNullOrEmpty(token))
                return false;

            GraphServiceClient client = new GraphServiceClient(
             new DelegateAuthenticationProvider(
             (requestMessage) =>
             {
                 requestMessage.Headers.Authorization =
                     new AuthenticationHeaderValue("Bearer", token);

                 return Task.FromResult(0);
             }));

            Stream stream = await client.Me.Drive.Items[driveItemID].Content.Request().GetAsync();

            FileStream fs = System.IO.File.Create(path, (int)stream.Length);
            byte[] bytesInStream = new byte[stream.Length];
            stream.Read(bytesInStream, 0, bytesInStream.Length);
            fs.Write(bytesInStream, 0, bytesInStream.Length);
            fs.Close();
            fs.Dispose();
            stream.Close();
            stream.Dispose();
            return true;
        }
        public async Task<string> CreateCalendarEvent(HttpContextBase httpContextBase, string name, List<Exam> exams)
        {
            string token = await GetAccessToken(httpContextBase);
            if (string.IsNullOrEmpty(token))
                return null;

            GraphServiceClient client = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);

                        return Task.FromResult(0);
                    }));

            DBManager db = new DBManager();
            if (db.GetUserGenCalendars(user).ToUpper() == "TRUE")
            {
                try
                {
                    Debug.WriteLine("Create Calendar ... ");
                    Microsoft.Graph.Calendar examCalendar = new Microsoft.Graph.Calendar
                    {
                        Name = name,
                        Events = new CalendarEventsCollectionPage(),
                    };

                    Debug.WriteLine("Check Calendars ... ");
                    var calendars = await client.Me.Calendars.Request().GetAsync();
                    bool exists = false;
                    foreach (Microsoft.Graph.Calendar calendar in calendars)
                    {
                        if (calendar.Name == name)
                        {
                            Debug.WriteLine("Delete Calendar ... ");
                            exists = true;
                            await client.Me.Calendars[calendar.Id].Request().DeleteAsync();
                        }
                    }
                    if (!exists)
                    {
                        Debug.WriteLine("Add Calendar ... ");
                        await client.Me.Calendars.Request().AddAsync(examCalendar);
                    }
                    foreach (Exam exam in exams)
                    {
                        DateTime examTime = Convert.ToDateTime(exam.Start);
                        DateTime start = new DateTime(
                            DateTime.Now.Year,
                            DateTime.Now.Month,
                            DateTime.Now.Day,
                            examTime.Hour,
                            examTime.Minute,
                            examTime.Second,
                            examTime.Millisecond,
                            DateTimeKind.Utc);

                        foreach (System.DayOfWeek weekday in Enum.GetValues(typeof(System.DayOfWeek)))
                        {
                            if (exam.Day.ToUpper() == weekday.ToString().ToUpper())
                            {
                                start = start.AddDays((weekday + 1) - (DateTime.Now.DayOfWeek + 1));
                                break;
                            }
                        }
                        TimeSpan span = Convert.ToDateTime(exam.End).Subtract(Convert.ToDateTime(exam.Start));
                        DateTime end = start.AddHours(span.Hours);
                        Event entry = new Event
                        {
                            Calendar = examCalendar,
                            Start = new DateTimeTimeZone
                            {
                                DateTime = string.Format("{0:s}", start),
                                TimeZone = TimeZone.CurrentTimeZone.StandardName
                            },
                            End = new DateTimeTimeZone
                            {
                                DateTime = string.Format("{0:s}", end),
                                TimeZone = TimeZone.CurrentTimeZone.StandardName
                            },
                            Body = new ItemBody
                            {
                                Content = exam.Room + "\n Faculty: " + exam.Faculty + "\n Proctor: " + exam.Proctor
                            },
                            Subject = exam.Code + ": " + exam.Name + ", Section " + exam.Section,
                            SingleValueExtendedProperties = new EventSingleValueExtendedPropertiesCollectionPage
                            {
                                new SingleValueLegacyExtendedProperty
                                {
                                    Id = "String " + Guid.NewGuid().ToString() + " Name TruckleSoft1",
                                    Value = "CLM_MidweekMeeting"
                                }
                            }
                        };
                        Debug.WriteLine("New Event ... ");
                        //await client.Me.Calendars[examCalendar.Id].Events.Request().AddAsync(entry);
                        await client.Me.Events.Request().AddAsync(entry);
                    }
                    //Debug.WriteLine(JsonConvert.SerializeObject(calendars));
                    //await client.Me.Calendars["Test"].Events.Request().AddAsync(entry);

                    //await client.Me.Events.Request().AddAsync(entry);
                }
                catch(Exception e)
                {
                    Debug.WriteLine(e.ToString());
                    Debug.WriteLine(e.InnerException);
                }
            }
            return "";
        }
    }
}