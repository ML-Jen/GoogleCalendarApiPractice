using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using System.ComponentModel.DataAnnotations;

namespace GoogleCalendarLibrary
{
    public class GoogleCalendar_Setting
    {
        [Required]
        public string client_email { get; set; }
        [Required]
        public string private_key { get; set; }
        [Required]
        public string calendar_id { get; set; }
        [Required]
        public string delegatingUser { get; set; }
        [Required]
        public string applicationName { get; set; }
    }
    public class GoogleCalendar
    {
        public CalendarService GoogleCalendarServiceSetting(GoogleCalendar_Setting setting)
        {
            string GoogleOAuth2EmailAddress = setting.client_email;
            string private_key = setting.private_key;
            string calendar_id = setting.calendar_id;

            var credential =
                new ServiceAccountCredential(
                new ServiceAccountCredential.Initializer(GoogleOAuth2EmailAddress)
                {
                    Scopes = new string[] { CalendarService.Scope.Calendar },
                }.FromPrivateKey(private_key));
            //https://cloud.google.com/dotnet/docs/reference/Google.Apis/latest/Google.Apis.Auth.OAuth2.GoogleCredential#Google_Apis_Auth_OAuth2_GoogleCredential_CreateWithUser_System_String_
            GoogleCredential googleCredential = GoogleCredential.FromServiceAccountCredential(credential).CreateWithUser(setting.delegatingUser);
            CalendarService service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = googleCredential,
                ApplicationName = setting.applicationName
            });
            return service;
        }
        public bool InsertEvent(GoogleCalendar_Setting setting,Event calendar_event,ref string add_event_id)
        {
            var service = GoogleCalendarServiceSetting(setting);
            EventsResource eventResource = new EventsResource(service);
            var insertEntryRequest = eventResource.Insert(calendar_event, setting.calendar_id);
            try
            {
                var insertResult = insertEntryRequest.Execute();
                add_event_id = insertResult.Id;
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool DeleteEvent(GoogleCalendar_Setting setting,string event_id)
        {
            var service = GoogleCalendarServiceSetting(setting);
            EventsResource eventResource = new EventsResource(service);
            var deleteEventRequest = eventResource.Delete(setting.calendar_id, event_id);
            try
            {
                var deleteResult = deleteEventRequest.Execute();
                return true;
            }
            catch
            {
                return false;
            }
        }

        public IList<Event> GetEventList(GoogleCalendar_Setting setting)
        {
            var service = GoogleCalendarServiceSetting(setting);
            var list = service.Events.List(setting.calendar_id);
            IList<Event> events = new List<Event>();
            try
            {
                events = list.Execute().Items;
            }
            catch { }
            return events;
        }
    }
}