# Office 365 Code Snippets for Windows

**Table of contents**

* [Introduction](#introduction)
* [Prerequisites](#prerequisites)
* [Register and configure the app](#register)
* [Build and debug](#build)
* [Add a snippet](#add-a-snippet)
* [Troubleshooting](#troubleshooting)
* [Questions and comments](#questions)
* [Next steps](#next-steps)
* [Additional resources](#additional-resources)

<a name="introduction"></a>
##Introduction

The Office 365 Windows Snippets project contains a repository of code snippets that show you how to use the client libraries from the Office 365 API tools to interact with Office 365 objects, including users, groups, calendars, contacts, mail, files, and folders.

These snippets are simple and self-contained, and you can copy and paste them into your own code, whenever appropriate, or use them as a resource for learning how to use the client libraries.

The image below shows what you'll see when you launch the app.

![](/Readme-images/MainPage.png "Launch page of O365 Window snippet app")

You can choose to run all of the snippets, or just the ones you select. After you choose to run, you’ll be prompted to authenticate with your Office 365 account credentials, and the snippets will run.

**Note:** This project contains code that authenticates and connects a user to Office 365, but if you want to learn about authentication specifically, look at the [Connecting to Office 365 in Windows Store, Phone, and universal apps](https://github.com/OfficeDev/O365-Win-Connect).

<a name="prerequisites"></a>
## Prerequisites ##

This sample requires the following:  
  - Visual Studio 2013 with Update 4.  
  - [Office 365 API Tools version 1.4.50428.2](http://aka.ms/k0534n).  
  - An Office 365 account. You can sign up for [an Office 365 Developer subscription](http://aka.ms/ro9c62) that includes the resources that you need to start building Office 365 apps.
 

<a name="register"></a>
###Register and configure the app

You can register the app with the Office 365 API Tools for Visual Studio. Be sure to download and install the [Office 365 API tools](http://aka.ms/k0534n) from the Visual Studio Gallery.

**Note:** If you see any errors while installing packages (for example, *Unable to find "Microsoft.IdentityModel.Clients.ActiveDirectory"*) make sure the local path where you placed the solution is not too long/deep. Moving the solution closer to the root of your drive resolves this issue.

   1. Open the O365-Win-Snippets.sln file using Visual Studio 2013.
   2. Build the solution. The NuGet Package Restore feature will load the assemblies listed in the packages.config file. You should do this before adding connected services in the following steps so that you don't get older versions of some assemblies.
   3. In the Solution Explorer window, right-click each project name and select Add -> Connected Service.
   4. A Services Manager dialog box will appear. Choose **Office 365** and then **Register your app**.
   5. On the sign-in dialog box, enter the user name and password for your Office 365 tenant. This user name will often follow the pattern <your-name>@<tenant-name>.onmicrosoft.com. If you don't already have an Office 365 tenant, you can get a free Developer Site as part of your MSDN Benefits or sign up for a free trial. After you're signed in, you will see a list of all the services. No permissions will be selected, since the app is not registered to use any services yet. 
   
   6. To register for the services used in this sample, choose the following services and permissions:

 	- **Calendar** - Read and write user calendars.
	- **Contacts** - Read and write user contacts.
	- **Mail** - Read and write user mail. Send mail as a user.
	- **My Files** - Read and write user files.
	- **Users and Groups** - Enable sign-on and read users' profiles. Access your organization's directory. 
	

The dialog will look like this:
![](/Readme-images/ConnectedServices.PNG "Windows Store connected services")

After you click **OK** in the Services Manager dialog box, you can select **Build Solution** from the **Build menu** to load the Microsoft.IdentityModel.Clients.ActiveDirectory assembly, or you can wait until you debug.


<a name="build"></a>
## Build and debug ##

After you've loaded the solution in Visual Studio, press F5 to build and debug.
Run the solution and sign in with your organizational account to Office 365.

<a name="add-a-snippet"></a>
##Add a snippet

This project includes five snippets files: Calendar\CalendarSnippets.cs, Contacts\ContactsSnippets.cs, Email\EmailSnippets.cs, Files\FilesSnippets.cs, and UsersAndGroups\UsersAndGroupsSnippets.cs.

If you have a snippet of your own and you would like to run it in this project, just follow these three steps:

1. **Add your snippet to the snippets file.** Be sure to include a try/catch block. The snippet below is an example of a simple snippet that gets one page of calendar events:

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
2. **Create a story that uses your snippet and add it to the associated stories file.** For example, the `TryCreateCalendarEventAsync()` story uses the `AddCalendarEventAsync ()` snippet inside the Calendar\CalendarStories.cs file:

        public static async Task<bool> TryCreateCalendarEventAsync()
        {
            var newEventId = await CalendarSnippets.AddCalendarEventAsync();

            if (newEventId == null)
                return false;

            //Cleanup
            await CalendarSnippets.DeleteCalendarEventAsync(newEventId);

            return true;
        }
Sometimes your story will need to run snippets in addition to the one that you're implementing. For example, if you want to update an event, you first need to use the `AddCalendarEventAsync()` method to create an event. Then you can update it. Always be sure to use snippets that already exist in the snippets file. If the operation you need doesn't exist, you'll have to create it and then include it in your story. It's a best practice to delete any entities that you create in a story, especially if you're working on anything other than a test or developer tenant.

3. **Add your story to the story collection in MainPageXaml.cs** (inside the `CreateTestList()` method):

	`StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Create", RunStoryAsync = CalendarStories.TryCreateCalendarEventAsync });`

Now you can test your snippet. When you run the app, your snippet will appear as a new box in the grid. Select the box for your snippet, and then run it. Use this as an opportunity to debug your snippet.

<a name="troubleshooting"></a>
## Troubleshooting ##

- You run the Windows App Certification Kit against the installed app, and the app fails the supported APIs test. This likely happened because the Visual Studio tools installed older versions of some assemblies. Check the entries for Microsoft.Azure.ActiveDirectory.GraphClient and the Microsoft.OData assemblies in your project's packages.config file. Make sure that the version numbers for those assemblies match the version numbers in [this repo's version of packages.config](https://github.com/OfficeDev/O365-Win-Snippets/blob/master/src/packages.config). When you rebuild and reinstall the solution with the updated assemblies, the app should pass the supported APIs test.

<a name="questions"></a>
##Questions and comments##

- If you have any trouble running this sample, please [log an issue](https://github.com/OfficeDev/O365-Win-Snippets/issues).
- For more general feedback, send and email to [docthis@microsoft.com](mailto:docthis@microsoft.com?subject=Feedback%20on%20the%20Office%20365%20Windows%20snippet%20app).


<a name="next-steps"></a>
## Next steps ##

- If you're interested in a sample that has a richer interface for interacting with the Office 365 services in a Windows app, look at the [Office 365 Starter Project for Windows Store App](https://github.com/OfficeDev/O365-Windows-Start).
- For more details on what else you can do with the Office 365 services in your Windows app, start with the [Getting started](http://aka.ms/rpx192) page on dev.office.com.

<a name="additional-resources"></a>
## Additional resources ##

- [Office 365 APIs documentation](http://aka.ms/kbwa5c)
- [Office 365 APIs starter projects and code samples](http://aka.ms/x1kpnz)
- [Office developer code samples](http://aka.ms/afh45z)
- [Office dev center](http://aka.ms/uftrm1)
