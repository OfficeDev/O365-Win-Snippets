# O365-Win-Snippets

**Table of contents**

* [Introduction](#introduction)
* [Register and configure the apps](#register)
* [Build and debug](#build)
* [Add a snippet](#add-a-snippet)
* [Next steps](#next-steps)
* [Additional resources](#additional-resources)

<a name="introduction"></a>
##Introduction

The Office 365 Windows Snippets project contains a repository of code snippets that show you how to use the client libraries from the Office 365 API tools to interact with Office 365 objects, including users, groups, calendars, contacts, mail, files, and folders.

You can use this project to learn how to work with entities and data in Office 365 services. You can also use it as a source of code snippets for your own projects. 

The image below shows what you'll see when you launch the app.

![](/Readme-images/MainPage.png "Windows Phone interface for the O365-WinPlatform-Connect sample")

After you choose to run all or run selected snippets, you'll be prompted to authenticate with your Office 365 account credentials, and the snippets will run.

**Note:** This project contains code that authenticates and connects a user to Office 365, but if you want to learn about authentication specifically, look at the [Connecting to Office 365 in Windows Store, Phone, and universal apps](https://github.com/OfficeDev/O365-Win-Connect).

<a name="register"></a>
###Register and configure the apps

You can register each app with the Office 365 API Tools for Visual Studio. Be sure to download and install the [Office 365 API tools](http://aka.ms/k0534n) from the Visual Studio Gallery.

**Note:** If you see any errors while installing packages during step 7 (for example, *Unable to find "Microsoft.IdentityModel.Clients.ActiveDirectory"*) make sure the local path where you placed the solution is not too long/deep. Moving the solution closer to the root of your drive resolves this issue.

   1. Open the O365-Win-Snippets.sln file using Visual Studio 2013.
   2. In the Solution Explorer window, right-click each project name and select Add -> Connected Service.
   3. A Services Manager dialog box will appear. Choose **Office 365** and then **Register your app**.
   4. On the sign-in dialog box, enter the user name and password for your Office 365 tenant. This user name will often follow the pattern <your-name>@<tenant-name>.onmicrosoft.com. If you don't already have an Office 365 tenant, you can get a free Developer Site as part of your MSDN Benefits or sign up for a free trial.
   5. After you're signed in, you will see a list of all the services. No permissions will be selected, since the app is not registered to use any services yet. 
   6. To register for the services used in this sample, choose the following services and permissions:

 	- **Calendar** - Have full access to users' calendars.
	- **Contacts** - Have full access to users' contacts.
	- **Mail** - Read users' mail. Read and write access to users' mail. Send mail as a user.
	- **My Files** - Edit or delete users' files.
	- **Users and Groups** - Access your organization's directory. 
	

The dialog will look like this:
![](/Readme-images/ConnectedServices.PNG "Windows Phone interface for the O365-WinPlatform-Connect sample")
   7. After you click **OK** in the Services Manager dialog box, assemblies for connecting to Office 365 APIs will be added to your project.

<a name="build"></a>
## Build and debug ##

After you've loaded the solution in Visual Studio, press F5 to build and debug.
Run the solution and sign in with your organizational account to Office 365.

<a name="add-a-snippet"></a>
##Add a snippet

If you have a snippet of your own and you would like to run it in this project, or even contribute it to this sample, follow these steps:

1. Identify the operations file where your snippet belongs. For example, if it works with Calendars, choose the Calendar\CalendarSnippets.cs file. The project also includes ContactsSnippets.cs, EmailSnippets.cs, FilesSnippets.cs, and UsersAndGroupsSnippets.cs files.
2. Add your snippet to the snippets file. Be sure to include a try/catch block. The snippet below is an example of a simple snippet that gets one page of calendar events:

        public static async Task<List<IEvent>> GetCalendarEventsAsync()
        {
            try
            {
                // Make sure we have a reference to the Exchange client
                OutlookServicesClient client = await GetOutlookClientAsync();

                IPagedCollection<IEvent> eventsResults = await client.Me.Calendar.Events.ExecuteAsync();

                // You can access each event as follows.
                if (eventsResults.CurrentPage.Count > 0)
                {
                    string eventId = eventsResults.CurrentPage[0].Id;
                    Debug.WriteLine("First event:" + eventId);
                }

                return eventsResults.CurrentPage.ToList();
            }
            catch { return null; }
        }
3. Create a story that implements your snippet and add it to the associated stories file. For example, the `TryCreateCalendarEventAsync()` story implements the `AddCalendarEventAsync ()` snippet inside the Calendar\CalendarStories.cs file:

        public static async Task<bool> TryCreateCalendarEventAsync()
        {
            var newEventId = await CalendarOperations.AddCalendarEventAsync();

            if (newEventId == null)
                return false;

            //Cleanup
            await CalendarOperations.DeleteCalendarEventAsync(newEventId);

            return true;
        }
Sometimes your story will need to run snippets in addition to the one that you're implementing. For example, if you want to update an event, you first need to use the `AddCalendarEventAsync()` method to create an event. Then you can update it. Always be sure to use snippets that already exist in the operations file. If the operation you need doesn't exist, you'll have to create it and then include it in your story. It's a best practice to delete any entities that you create in a story, especially if you're working on anything other than a test or developer tenant.

4. Add your story to the story collection in MainPageXaml.cs:

	`StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Create", RunStoryAsync = CalendarStories.TryCreateCalendarEventAsync });`

5. Test your snippet. When you run the app, your snippet will appear as a new box in the grid. Select the box for your snippet, and then run it. Use this as an opportunity to debug your snippet.

<a name="next-steps"></a>
## Next steps ##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/O365-Win-Snippets/issues).
- If you're interested in a sample that has a richer interface for interacting with the Office 365 services in a Windows app, look at the [Office 365 Starter Project for Windows Store App](https://github.com/OfficeDev/O365-Windows-Start).
- For more details on what else you can do with the Office 365 services in your Windows app, start with the [Getting started](http://aka.ms/rpx192) page on dev.office.com.

<a name="additional-resources"></a>
## Additional resources ##

- [Office 365 APIs documentation](http://aka.ms/kbwa5c)
- [Office 365 APIs starter projects and code samples](http://aka.ms/x1kpnz)
- [Office developer code samples](http://aka.ms/afh45z)
- [Office dev center](http://aka.ms/uftrm1)