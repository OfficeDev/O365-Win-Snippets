using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace O365_Win_Snippets
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public List<StoryDefinition> StoryCollection { get; private set; }
        public MainPage()
        {
            this.InitializeComponent();
            CreateTestList();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            // Developer code - if you haven't registered the app yet, we warn you. 
            if (!App.Current.Resources.ContainsKey("ida:ClientID"))
            {
                Debug.WriteLine("Oops - App not registered with Office 365. To run this sample, you must register it with Office 365. You can do that through the 'Add | Connected services' dialog in Visual Studio. See Readme for more info");
 
            }
        }
        private void CreateTestList()
        {
            StoryCollection = new List<StoryDefinition>();

            // These stories require your app to have permission to access your organization's directory. 
            // Comment them if you're not going to run the app with that permission level.

            StoryCollection.Add(new StoryDefinition() { GroupName = "Users & Groups", Title = "Client", RunStoryAsync = UsersAndGroupsStories.TryGetAadGraphClientAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users & Groups", Title = "Read Users", RunStoryAsync = UsersAndGroupsStories.TryGetUsersAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users & Groups", Title = "Tenant Details", RunStoryAsync = UsersAndGroupsStories.TryGetTenantAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Users & Groups", Title = "Read Groups", RunStoryAsync = UsersAndGroupsStories.TryGetGroupsAsync });

            StoryCollection.Add(new StoryDefinition() { GroupName = "Contacts", Title = "Client", RunStoryAsync = ContactsStories.TryGetOutlookClientAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Contacts", Title = "Read", RunStoryAsync = ContactsStories.TryGetContactsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Contacts", Title = "Get contact", RunStoryAsync = ContactsStories.TryGetContactAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Contacts", Title = "Create", RunStoryAsync = ContactsStories.TryAddNewContactAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Contacts", Title = "Delete", RunStoryAsync = ContactsStories.TryDeleteContactAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Contacts", Title = "Update", RunStoryAsync = ContactsStories.TryUpdateContactAsync });


            StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Client", RunStoryAsync = CalendarStories.TryGetOutlookClientAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Read", RunStoryAsync = CalendarStories.TryGetCalendarEventsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Create", RunStoryAsync = CalendarStories.TryCreateCalendarEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Create with args", RunStoryAsync = CalendarStories.TryCreateCalendarEventWithArgsAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Update", RunStoryAsync = CalendarStories.TryUpdateCalendarEventAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Calendar", Title = "Delete", RunStoryAsync = CalendarStories.TryDeleteCalendarEventAsync });


            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Client", RunStoryAsync = EmailStories.TryGetOutlookClientAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Read Inbox", RunStoryAsync = EmailStories.TryGetInboxMessagesAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Read messages", RunStoryAsync = EmailStories.TryGetMessagesAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "SendMail", RunStoryAsync = EmailStories.TrySendMessageAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Reply", RunStoryAsync = EmailStories.TryReplyMessageAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Reply All", RunStoryAsync = EmailStories.TryReplyAllAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Forward", RunStoryAsync = EmailStories.TryForwardMessageAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Create draft", RunStoryAsync = EmailStories.TryCreateDraftAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Update", RunStoryAsync = EmailStories.TryUpdateMessageAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Delete", RunStoryAsync = EmailStories.TryDeleteMessageAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Move", RunStoryAsync = EmailStories.TryMoveMessageAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Email", Title = "Copy", RunStoryAsync = EmailStories.TryCopyMessageAsync });


            StoryCollection.Add(new StoryDefinition() { GroupName = "Mail folder", Title = "Read Folders", RunStoryAsync = EmailStories.TryGetMailFoldersAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Mail folder", Title = "Create", RunStoryAsync = EmailStories.TryCreateMailFolderAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Mail folder", Title = "Rename", RunStoryAsync = EmailStories.TryUpdateMailFolderAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Mail folder", Title = "Move", RunStoryAsync = EmailStories.TryMoveMailFolderAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Mail folder", Title = "Copy", RunStoryAsync = EmailStories.TryCopyMailFolderAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Mail folder", Title = "Delete", RunStoryAsync = EmailStories.TryDeleteMailFolderAsync });



            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Client", RunStoryAsync = FilesStories.TryGetSharePointClientAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Read folders", RunStoryAsync = FilesStories.TryGetFolderChildrenAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Create folder", RunStoryAsync = FilesStories.TryCreateFolderAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Delete folder", RunStoryAsync = FilesStories.TryDeleteFolderAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Create file", RunStoryAsync = FilesStories.TryCreateFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Update content", RunStoryAsync = FilesStories.TryUpdateFileContentAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Delete file", RunStoryAsync = FilesStories.TryDeleteFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Download", RunStoryAsync = FilesStories.TryDownloadFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Copy file", RunStoryAsync = FilesStories.TryCopyFileAsync });
            StoryCollection.Add(new StoryDefinition() { GroupName = "Files", Title = "Rename file", RunStoryAsync = FilesStories.TryRenameFileAsync });

            var result = from story in StoryCollection group story by story.GroupName into api orderby api.Key select api;
            StoriesByApi.Source = result; 
        }


        private async void RunSelectedStories_Click(object sender, RoutedEventArgs e)
        {

            await runSelectedAsync();
        }

        private async Task runSelectedAsync()
        {
            ResetStories();

            foreach (var story in StoryGrid.SelectedItems)
            {
                StoryDefinition currentStory = story as StoryDefinition;
                currentStory.IsRunning = true;
                Stopwatch sw = new Stopwatch();
                sw.Start();
                currentStory.Result = await currentStory.RunStoryAsync();
                sw.Stop();
                currentStory.DurationMS = sw.ElapsedMilliseconds;
                currentStory.IsRunning = false;


                Debug.WriteLine(String.Format("{0} {1}", currentStory.Title, (currentStory.Result.HasValue && currentStory.Result.Value) ? "passed" : "failed"));

            }

            // To shut down this app when the Stories complete, uncomment the following line. 
            //Application.Current.Exit();
        }

        private async void RunAll_Click(object sender, RoutedEventArgs e)
        {
            StoryGrid.SelectedItems.Clear();
            foreach (var item in StoryGrid.Items)
            {
                StoryGrid.SelectedItems.Add(item);
            }
            await runSelectedAsync();
        }

        private void ResetStories()
        {
            foreach (var story in StoryCollection)
            {
                story.Result = null;
                story.DurationMS = null;
            }
        }

        private void ClearSelection_Click(object sender, RoutedEventArgs e)
        {
            StoryGrid.SelectedItems.Clear();
        }

        private void Disconnect_Click(object sender, RoutedEventArgs e)
        {
            AuthenticationHelper.SignOut();
            StoryGrid.SelectedItems.Clear();
        }
    }
}
