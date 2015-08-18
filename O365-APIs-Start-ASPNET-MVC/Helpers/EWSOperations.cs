using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using model = O365_APIs_Start_ASPNET_MVC.Models;
using System.Threading.Tasks;
using O365_APIs_Start_ASPNET_MVC.Utils;

namespace O365_APIs_Start_ASPNET_MVC.Helpers
{
    internal class EWSOperations
    {
        #region working code for exchange service
        internal async Task<List<model.EWSTaskItem>> getEWSTasks()
        {
            string ewsResourceId = SettingsHelper.EwsResourceId;
            var ewsAuthToken = await AuthenticationHelper.EnsureResourceClientCreatedAsync(ewsResourceId);
            ExchangeService serviceObj = getExchangeService(ewsAuthToken.AccessToken, SettingsHelper.EwsWebServiceUri);

            List<model.EWSTaskItem> ewsTasks = new List<Models.EWSTaskItem>();
            var ewsTaskList = FindIncompleteTask(serviceObj);
            foreach (Microsoft.Exchange.WebServices.Data.Task eTask in ewsTaskList)
            {
                model.EWSTaskItem ewsTaskItemObj = new model.EWSTaskItem(eTask);
                ewsTasks.Add(ewsTaskItemObj);
            }

            return ewsTasks;
        }

        public ExchangeService getExchangeService(string ewsAuthToken, string EwsWebServiceUri)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
            service.Credentials = new OAuthCredentials(ewsAuthToken);
            //service.HttpHeaders.Add("Authorization", "Bearer " + tokenx);
            service.PreAuthenticate = true;
            service.SendClientLatencies = true;
            service.EnableScpLookup = false;
            service.Url = new Uri(EwsWebServiceUri);

            return service;
        }

        static IEnumerable<Microsoft.Exchange.WebServices.Data.Task> FindIncompleteTask(ExchangeService service)
        {
            // Specify the folder to search, and limit the properties returned in the result.
            TasksFolder tasksfolder = TasksFolder.Bind(service,
                                                       WellKnownFolderName.Tasks,
                                                       new PropertySet(BasePropertySet.IdOnly, FolderSchema.TotalCount));

            // Set the number of items to the smaller of the number of items in the Contacts folder or 1000.
            int numItems = tasksfolder.TotalCount < 1000 ? tasksfolder.TotalCount : 1000;

            // Instantiate the item view with the number of items to retrieve from the contacts folder.
            ItemView view = new ItemView(numItems);

            // To keep the request smaller, send only the display name.
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, TaskSchema.Subject, TaskSchema.Status, TaskSchema.StartDate);

            var filter = new SearchFilter.IsGreaterThan(TaskSchema.DateTimeCreated, DateTime.Now.AddYears(-10));

            // Retrieve the items in the Tasks folder with the properties you selected.
            FindItemsResults<Microsoft.Exchange.WebServices.Data.Item> taskItems = service.FindItems(WellKnownFolderName.Tasks, filter, view);

            // If the subject of the task matches only one item, return that task item.
            var results = new List<Microsoft.Exchange.WebServices.Data.Task>();
            foreach (var task in taskItems)
            {
                var result = new Microsoft.Exchange.WebServices.Data.Task(service);
                result = task as Microsoft.Exchange.WebServices.Data.Task;
                results.Add(result);
            }
            return results;
        }

        internal async Task<string> getRoomsAndAvailability()
        {
            string output = string.Empty;
            string ewsResourceId = SettingsHelper.EwsResourceId;
            var ewsAuthToken = await AuthenticationHelper.EnsureResourceClientCreatedAsync(ewsResourceId);
            ExchangeService serviceObj = getExchangeService(ewsAuthToken.AccessToken, SettingsHelper.EwsWebServiceUri);



            return output;
        }

        private static GetUserAvailabilityResults GetSuggestedMeetingTimesAndFreeBusyInfo(ExchangeService service)
        {
            // Create a collection of attendees. 
            List<AttendeeInfo> attendees = new List<AttendeeInfo>();

            attendees.Add(new AttendeeInfo()
            {
                SmtpAddress = "mack@contoso.com", //get the current user email
                AttendeeType = MeetingAttendeeType.Organizer
            });

            attendees.Add(new AttendeeInfo()
            {
                SmtpAddress = "sadie@contoso.com",
                AttendeeType = MeetingAttendeeType.Required
            });

            // Specify options to request free/busy information and suggested meeting times.
            AvailabilityOptions availabilityOptions = new AvailabilityOptions();
            availabilityOptions.GoodSuggestionThreshold = 49;
            availabilityOptions.MaximumNonWorkHoursSuggestionsPerDay = 0;
            availabilityOptions.MaximumSuggestionsPerDay = 2;
            // Note that 60 minutes is the default value for MeetingDuration, but setting it explicitly for demonstration purposes.
            availabilityOptions.MeetingDuration = 60;
            availabilityOptions.MinimumSuggestionQuality = SuggestionQuality.Good;
            availabilityOptions.DetailedSuggestionsWindow = new TimeWindow(DateTime.Now.AddDays(1), DateTime.Now.AddDays(2));
            availabilityOptions.RequestedFreeBusyView = FreeBusyViewType.FreeBusy;

            // Return free/busy information and a set of suggested meeting times. 
            // This method results in a GetUserAvailabilityRequest call to EWS.
            GetUserAvailabilityResults results = null;
            results = service.GetUserAvailability(attendees,
                                                  availabilityOptions.DetailedSuggestionsWindow,
                                                  AvailabilityData.FreeBusyAndSuggestions,
                                                  availabilityOptions);
            return results;
            #region find the suggestions and availability
            // Display suggested meeting times. 
            //Console.WriteLine("Availability for {0} and {1}", attendees[0].SmtpAddress, attendees[1].SmtpAddress);
            //Console.WriteLine();

            //foreach (Suggestion suggestion in results.Suggestions)
            //{
            //    Console.WriteLine("Suggested date: {0}\n", suggestion.Date.ToShortDateString());
            //    Console.WriteLine("Suggested meeting times:\n");
            //    foreach (TimeSuggestion timeSuggestion in suggestion.TimeSuggestions)
            //    {
            //        Console.WriteLine("\t{0} - {1}\n",
            //                          timeSuggestion.MeetingTime.ToShortTimeString(),
            //                          timeSuggestion.MeetingTime.Add(TimeSpan.FromMinutes(availabilityOptions.MeetingDuration)).ToShortTimeString());



            //    }
            //}

            //int i = 0;

            //// Display free/busy times.
            //foreach (AttendeeAvailability availability in results.AttendeesAvailability)
            //{
            //    Console.WriteLine("Availability information for {0}:\n", attendees[i].SmtpAddress);

            //    foreach (CalendarEvent calEvent in availability.CalendarEvents)
            //    {
            //        Console.WriteLine("\tBusy from {0} to {1} \n", calEvent.StartTime.ToString(), calEvent.EndTime.ToString());
            //    }

            //    i++;
            //}

            #endregion
        }

        #endregion

        #region old code
        //public string runNewcodewithToken(string token)
        //{
        //    try
        //    {
        //        string tokenx = token;

        //        ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013);
        //        //service.TraceListener = ITraceListener
        //        service.TraceEnabled = true;
        //        service.TraceFlags = TraceFlags.All;
        //        //service.Credentials = new OAuthCredentials(tokenx);
        //        service.HttpHeaders.Add("Authorization", "Bearer " + tokenx);

        //        service.PreAuthenticate = true;
        //        service.SendClientLatencies = true;
        //        service.EnableScpLookup = false;
        //        service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");

        //        service.GetRoomLists();

        //        IEnumerable<Microsoft.Exchange.WebServices.Data.Task> tasks = FindIncompleteTask(service);

        //    }
        //    catch (Exception ex)
        //    { Console.Write(ex); }

        //    return string.Empty;
        //}
        #endregion

    }
}
