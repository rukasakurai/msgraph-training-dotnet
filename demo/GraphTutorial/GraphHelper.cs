// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using System.Text.Json;

class GraphHelper
{
    #region User-auth
    // <UserAuthConfigSnippet>
    // Settings object
    private static Settings? _settings;
    // User auth token credential
    private static DeviceCodeCredential? _deviceCodeCredential;
    // Client configured with user authentication
    private static GraphServiceClient? _userClient;

    public static void InitializeGraphForUserAuth(Settings settings,
        Func<DeviceCodeInfo, CancellationToken, Task> deviceCodePrompt)
    {
        _settings = settings;

        _deviceCodeCredential = new DeviceCodeCredential(deviceCodePrompt,
            settings.AuthTenant, settings.ClientId);

        _userClient = new GraphServiceClient(_deviceCodeCredential, settings.GraphUserScopes);
    }
    // </UserAuthConfigSnippet>

    // <GetUserTokenSnippet>
    public static async Task<string> GetUserTokenAsync()
    {
        // Ensure credential isn't null
        _ = _deviceCodeCredential ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Ensure scopes isn't null
        _ = _settings?.GraphUserScopes ?? throw new System.ArgumentNullException("Argument 'scopes' cannot be null");

        // Request token with given scopes
        var context = new TokenRequestContext(_settings.GraphUserScopes);
        var response = await _deviceCodeCredential.GetTokenAsync(context);
        return response.Token;
    }
    // </GetUserTokenSnippet>

    // <GetUserSnippet>
    public static Task<User> GetUserAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            .Request()
            .Select(u => new
            {
                // Only request specific properties
                u.DisplayName,
                u.Mail,
                u.UserPrincipalName
            })
            .GetAsync();
    }
    // </GetUserSnippet>

    // <GetInboxSnippet>
    public static Task<IMailFolderMessagesCollectionPage> GetInboxAsync()
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        return _userClient.Me
            // Only messages from Inbox folder
            .MailFolders["Inbox"]
            .Messages
            .Request()
            .Select(m => new
            {
                // Only request specific properties
                m.From,
                m.IsRead,
                m.ReceivedDateTime,
                m.Subject
            })
            // Get at most 25 results
            .Top(25)
            // Sort by received time, newest first
            .OrderBy("ReceivedDateTime DESC")
            .GetAsync();
    }
    // </GetInboxSnippet>

    // <SendMailSnippet>
    public static async Task SendMailAsync(string subject, string body, string recipient)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        // Create a new message
        var message = new Message
        {
            Subject = subject,
            Body = new ItemBody
            {
                Content = body,
                ContentType = BodyType.Text
            },
            ToRecipients = new Recipient[]
            {
                new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipient
                    }
                }
            }
        };

        // Send the message
        await _userClient.Me
            .SendMail(message)
            .Request()
            .PostAsync();
    }
    // </SendMailSnippet>
    #endregion

    #region App-only
    // <AppOnyAuthConfigSnippet>
    // App-ony auth token credential
    private static ClientSecretCredential? _clientSecretCredential;
    // Client configured with app-only authentication
    private static GraphServiceClient? _appClient;

    private static void EnsureGraphForAppOnlyAuth()
    {
        // Ensure settings isn't null
        _ = _settings ??
            throw new System.NullReferenceException("Settings cannot be null");

        if (_clientSecretCredential == null)
        {
            _clientSecretCredential = new ClientSecretCredential(
                _settings.TenantId, _settings.ClientId, _settings.ClientSecret);
        }

        if (_appClient == null)
        {
            _appClient = new GraphServiceClient(_clientSecretCredential,
                // Use the default scope, which will request the scopes
                // configured on the app registration
                new[] {"https://graph.microsoft.com/.default"});
        }
    }
    // </AppOnyAuthConfigSnippet>

    // <GetUsersSnippet>
    public static Task<IGraphServiceUsersCollectionPage> GetUsersAsync()
    {
        EnsureGraphForAppOnlyAuth();
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        return _appClient.Users
            .Request()
            .Select(u => new
            {
                // Only request specific properties
                u.DisplayName,
                u.Id,
                u.Mail
            })
            // Get at most 25 results
            .Top(25)
            // Sort by display name
            .OrderBy("DisplayName")
            .GetAsync();
    }
    // </GetUsersSnippet>
    #endregion

    #pragma warning disable CS1998
    // <MakeGraphCallSnippet>
    // This function serves as a playground for testing Graph snippets
    // or other code
    public async static Task MakeGraphCallAsync()
    {
        // INSERT YOUR CODE HERE
        // Note: if using _appClient, be sure to call EnsureGraphForAppOnlyAuth
        // before using it.
        // EnsureGraphForAppOnlyAuth();
    }
    // </MakeGraphCallSnippet>

    public async static Task ListMembersInGroupAsync()
    {
        EnsureGraphForAppOnlyAuth();
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var members = await _appClient.Groups["4b5c4fad-ba66-4628-9d0b-46ada2a47345"].Members
            .Request()
            .GetAsync();

        foreach(User member in members){
            //Console.WriteLine(JsonSerializer.Serialize(member));
            //Console.WriteLine(member.GetType());
            Console.WriteLine(member.DisplayName);
        }
    }

    //https://docs.microsoft.com/en-us/graph/api/calendar-getschedule?view=graph-rest-1.0&tabs=csharp
   public async static Task GetScheduleAsync(string userEmail)
    {
        EnsureGraphForAppOnlyAuth();
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var schedules = new List<String>()
        {
            userEmail
        };

        var startTime = new DateTimeTimeZone
        {
            DateTime = "2019-03-15T09:00:00",
            TimeZone = "Pacific Standard Time"
        };

        var endTime = new DateTimeTimeZone
        {
            DateTime = "2019-03-15T18:00:00",
            TimeZone = "Pacific Standard Time"
        };

        var availabilityViewInterval = 60;

        var results = await _appClient.Users[userEmail].Calendar
            .GetSchedule(schedules,endTime,startTime,availabilityViewInterval)
            .Request()
            .Header("Prefer","outlook.timezone=\"Pacific Standard Time\"")
            .PostAsync();

        //Console.WriteLine(results.GetType());
        foreach(object result in results){
            Console.WriteLine(JsonSerializer.Serialize(result));
            //Console.WriteLine(result.GetType());
            //Console.WriteLine(result.DisplayName);
        }
    }
    

    // https://docs.microsoft.com/en-us/graph/api/user-findmeetingtimes?view=graph-rest-1.0&tabs=http
    public async static Task FindMeetingTimes(string userEmail)
    {
        // Ensure client isn't null
        _ = _userClient ??
            throw new System.NullReferenceException("Graph has not been initialized for user auth");

        MeetingTimeSuggestionsResult meetingResponse = await _userClient.Me
            .FindMeetingTimes()
            .Request()
            .Header("Prefer", "outlook.timezone=\"W. Europe Standard Time\"")
            .PostAsync();

        Console.WriteLine(JsonSerializer.Serialize(meetingResponse));
        //Console.WriteLine(member.DisplayName);
    }

    // https://docs.microsoft.com/en-us/graph/api/calendar-post-events?view=graph-rest-1.0&tabs=csharp#example-2-create-and-enable-an-event-as-an-online-meeting
    public async static Task CreateEventAsync(string userEmail)
    {
        EnsureGraphForAppOnlyAuth();
        // Ensure client isn't null
        _ = _appClient ??
            throw new System.NullReferenceException("Graph has not been initialized for app-only auth");

        var @event = new Event
        {
            Subject = "Let's go for lunch",
            Body = new ItemBody
            {
                ContentType = BodyType.Html,
                Content = "Does noon work for you?"
            },
            Start = new DateTimeTimeZone
            {
                DateTime = "2022-07-15T12:00:00",
                TimeZone = "Pacific Standard Time"
            },
            End = new DateTimeTimeZone
            {
                DateTime = "2022-07-15T14:00:00",
                TimeZone = "Pacific Standard Time"
            },
            Location = new Location
            {
                DisplayName = "Harry's Bar"
            },
            Attendees = new List<Attendee>()
            {
                new Attendee
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = "samanthab@contoso.onmicrosoft.com",
                        Name = "Samantha Booth"
                    },
                    Type = AttendeeType.Required
                }
            },
            AllowNewTimeProposals = true,
            IsOnlineMeeting = true,
            OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness
        };

        // Send the message
        await _appClient.Users[userEmail]
            .Events.Request().Header("Prefer","outlook.timezone=\"Pacific Standard Time\"").AddAsync(@event);
    }
    
}
