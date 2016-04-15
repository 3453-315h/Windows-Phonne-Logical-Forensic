using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using Windows.ApplicationModel.Contacts;
using Windows.Security.ExchangeActiveSyncProvisioning;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;
using Windows.ApplicationModel.Appointments;


// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace Contact_Appointements
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
       
        string filename = "contact_thumbnail_";
        string todayDateTime = DateTime.Now.ToString().Replace("/", "").Replace(":", "").Replace(" ", "_");
        StorageFolder globalStorageFolder = KnownFolders.PicturesLibrary;
        public MainPage()
        {

            this.InitializeComponent();

        }

        /// <summary>
        /// Invoked when this page is about to be displayed in a Frame.
        /// </summary>
        /// <param name="e">Event data that describes how this page was reached.
        /// This parameter is typically used to configure the page.</param>
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            // TODO: Prepare page for display here.

            // TODO: If your application contains multiple pages, ensure that you are
            // handling the hardware Back button by registering for the
            // Windows.Phone.UI.Input.HardwareButtons.BackPressed event.
            // If you are using the NavigationHelper provided by some templates,
            // this event is handled for you.
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            if (chckbContacts.IsChecked == true)
            {
                StartSearchContacts();
            }
            if (chckbAppointments.IsChecked == true)
            {
                StartSearchAppointments();
            }
        }
        private void buttonExit_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Exit();
        }
        private async void StartSearchContacts()
        {
            listLog.Items.Add("Searching Contacts...");
            listLog.SelectedIndex = 0;
            ContactStore contactStore = await ContactManager.RequestStoreAsync();
            IReadOnlyList<Contact> contacts = null;
            contacts = await contactStore.FindContactsAsync();
            listLog.Items.Add("The number of contacts is: " + contacts.Count.ToString());
            Save_Contacts(contacts);
        }

    
        private string getdeviceInfo(int totalItems)
        {
            EasClientDeviceInformation deviceInfo = new EasClientDeviceInformation();
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<strong>Phone Timezone:</strong> " + TimeZoneInfo.Local + "<br>");
            sb.AppendLine("<strong>FriendlyName:</strong>  " + deviceInfo.FriendlyName + "<br>");
            sb.AppendLine("<strong>OS:</strong>  " + deviceInfo.OperatingSystem);
            sb.AppendLine("<strong>Firmware Version:</strong>  " + deviceInfo.SystemFirmwareVersion + "<br>");
            sb.AppendLine("<strong>Hardware Version:</strong>  " + deviceInfo.SystemHardwareVersion + "<br>");
            sb.AppendLine("<strong>Product Name:</strong>  " + deviceInfo.SystemProductName + "<br>");
            sb.AppendLine("<strong>Store Keeping Unit:</strong>  " + deviceInfo.SystemSku + "<br>");
            sb.AppendLine("<strong>Total Items Extracted:</strong>  " + totalItems + "<br>");
            sb.AppendLine("<br>");
            return sb.ToString();
        }
        private string getThumbnailName(Contact c)
        {
            if (c != null)
            {
                return c.FirstName + "_" + c.LastName + ".jpg";

            }
            return null;
        }

        async void Save_Contacts(IReadOnlyList<Contact> cntctsList)
        {
            try
            {
                

                StorageFolder localFolderStorage = await globalStorageFolder.CreateFolderAsync("WPLogical_" + todayDateTime, CreationCollisionOption.OpenIfExists);
                StorageFile contactHtmlFile = await localFolderStorage.CreateFileAsync("contacts_" + todayDateTime + ".html", CreationCollisionOption.OpenIfExists);

                var streamHTML = await contactHtmlFile.OpenAsync(FileAccessMode.ReadWrite);

                var writerHTML = new DataWriter(streamHTML.GetOutputStreamAt(0));
                StringBuilder htmlBuilder = new StringBuilder();
                StringBuilder tempStringBuilder = new StringBuilder();

                //Generate HTML table headers
                htmlBuilder.AppendLine("<html>");
                htmlBuilder.AppendLine("<head>");
                htmlBuilder.AppendLine("<style type=\"text/css\">");
                htmlBuilder.AppendLine(".tg  {border-collapse:collapse;border-spacing:0;}");
                htmlBuilder.AppendLine(".tg td{font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;}");
                htmlBuilder.AppendLine(".tg th{font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;}");
                htmlBuilder.AppendLine(".tg .tg-qk49{font-weight:bold;background-color:#306f98;color:#ffffff;vertical-align:top}");
                htmlBuilder.AppendLine(".tg .tg-yw4l{vertical-align:top}");
                htmlBuilder.AppendLine("</style>");
                htmlBuilder.AppendLine("<meta charset=\"UTF-8\">");
                htmlBuilder.AppendLine("</head>");
                htmlBuilder.AppendLine("<body>");
                //Table details starts
                htmlBuilder.AppendLine("<table class=\"tg\">");
                htmlBuilder.AppendLine("  <tr>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Acquisition details</th>");
                htmlBuilder.AppendLine("  </tr>");
                htmlBuilder.AppendLine("  <tr>");
                htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + getdeviceInfo(cntctsList.Count) + "</td>");
                htmlBuilder.AppendLine("  </tr>");
                //Table details ends
                htmlBuilder.AppendLine("</table>");
                htmlBuilder.AppendLine("<table class=\"tg\">");
                htmlBuilder.AppendLine("  <tr>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Display Name</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">First Name</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Middle Name</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Last Name</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Phones</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Important Dates</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Emails</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Websites</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Job Info</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Adresses</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Notes</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Thumbnail</th>");
                htmlBuilder.AppendLine("  </tr>");
                foreach (Contact cntct in cntctsList)
                {

                    if (cntct.Thumbnail != null)
                    {
                        IRandomAccessStream fileStream = await cntct.Thumbnail.OpenReadAsync();
                    }

                    //Save acquired contact details to a formatted HTML table

                    htmlBuilder.AppendLine("  <tr>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (cntct.DisplayName != null ? cntct.DisplayName : "NULL") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (cntct.FirstName != null ? cntct.FirstName : "NULL") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (cntct.MiddleName != null ? cntct.MiddleName : "NULL") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (cntct.LastName != null ? cntct.LastName : "NULL") + "</td>");

                    if (cntct.Phones != null)
                    {
                        foreach (ContactPhone item in cntct.Phones)
                        {
                            tempStringBuilder.AppendLine("Kind: " + item.Kind + " Number: " + item.Number);
                        }
                    }

                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (tempStringBuilder.ToString() != null ? tempStringBuilder.ToString() : "NULL") + "</td>");

                    if (cntct.ImportantDates != null)
                    {
                        tempStringBuilder.Clear();
                        foreach (ContactDate item in cntct.ImportantDates)
                        {
                            tempStringBuilder.AppendLine("Kind: " + item.Kind + " D: " + item.Day + "M: " + item.Month + "Y: " + item.Year + " Description: " + item.Description);
                        }
                    }
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (tempStringBuilder.ToString() != null ? tempStringBuilder.ToString() : "NULL") + "</td>");

                    if (cntct.Emails != null)
                    {
                        tempStringBuilder.Clear();
                        foreach (ContactEmail item in cntct.Emails)
                        {
                            tempStringBuilder.AppendLine("Kind: " + item.Kind + " Email: " + item.Address + " Description: " + item.Description);
                        }
                    }

                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (tempStringBuilder.ToString() != null ? tempStringBuilder.ToString() : "NULL") + "</td>");

                    if (cntct.Websites != null)
                    {
                        tempStringBuilder.Clear();
                        foreach (ContactWebsite item in cntct.Websites)
                        {
                            tempStringBuilder.AppendLine("URL: " + item.Uri.AbsoluteUri.ToString() + " Description: " + item.Description);
                        }
                    }
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (tempStringBuilder.ToString() != null ? tempStringBuilder.ToString() : "NULL") + "</td>");

                    if (cntct.JobInfo != null)
                    {
                        tempStringBuilder.Clear();
                        foreach (ContactJobInfo item in cntct.JobInfo)
                        {
                            tempStringBuilder.AppendLine("Title: " + item.Title + " Office" + item.Office + " CompanyName: " + item.CompanyName + " Description: " + item.Description);
                        }
                    }

                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (tempStringBuilder.ToString() != null ? tempStringBuilder.ToString() : "NULL") + "</td>");

                    if (cntct.Addresses != null)
                    {
                        tempStringBuilder.Clear();
                        foreach (ContactAddress item in cntct.Addresses)
                        {
                            tempStringBuilder.AppendLine("Kind: " + item.Kind + " Locality" + item.Locality + " PostalCode: " + item.PostalCode + " Region: " + item.Region + " StreetAddress: " + item.StreetAddress + " Description: " + item.Description);
                        }
                    }
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (tempStringBuilder.ToString() != null ? tempStringBuilder.ToString() : "NULL") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (cntct.Notes != null ? cntct.Notes : "NULL") + "</td>");

                    try
                    {
                        IRandomAccessStreamWithContentType streamThumb = await cntct.Thumbnail.OpenReadAsync();

                        if (streamThumb != null && streamThumb.Size > 0)
                        {

                            byte[] buffer = new byte[streamThumb.Size];
                            IBuffer test = await streamThumb.ReadAsync(buffer.AsBuffer(), (uint)streamThumb.Size, InputStreamOptions.None);


                            var thumbB64 = System.Convert.ToBase64String(test.ToArray());

                            htmlBuilder.AppendLine("    <td class=\"tg-yw4l\"><a href=\"data: image / jpeg; base64," + thumbB64 + "\" download =\"" + filename + (cntct.FirstName != null ? cntct.FirstName : "NULL") + "_" + (cntct.LastName != null ? cntct.LastName : "NULL") + ".jpg\" ><img src=\"data: image / jpeg; base64," + thumbB64 + "\" height=\"256\" width=\"256\"   > </a></td>");
                        }
                    }
                    catch (Exception)
                    {
                        htmlBuilder.AppendLine("    <td class=\"tg-yw4l\"> N/A </td>");

                    }


                    //htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">"Thumbnail"</td>");
                    htmlBuilder.AppendLine("  </tr>");


                    listLog.Items.Add("Contacts saved: " + cntct.FirstName + " " + cntct.LastName);
                    listLog.SelectedIndex = listLog.Items.Count - 1;

                }
                htmlBuilder.AppendLine("</table>");
                htmlBuilder.AppendLine("</body> ");
                htmlBuilder.AppendLine("</html>");

                writerHTML.WriteString(htmlBuilder.ToString());
                await writerHTML.StoreAsync();
                await writerHTML.FlushAsync();
                listLog.SelectedIndex = listLog.Items.Count - 1;
                ShowMessageFinished("Contacts");
            }
            catch (Exception ee)
            {
                MessageDialog msgbox = new MessageDialog(ee.Message);
                await msgbox.ShowAsync();
            }
        }

        private async void StartSearchAppointments()
        {
           
            listLog.Items.Add("Searching  Appointments...");
            listLog.SelectedIndex = listLog.Items.Count - 1;

            AppointmentStore appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadOnly);
            FindAppointmentsOptions options = new FindAppointmentsOptions();

            options.FetchProperties.Add(AppointmentProperties.Subject);
            options.FetchProperties.Add(AppointmentProperties.Location);
            options.FetchProperties.Add(AppointmentProperties.Invitees);
            options.FetchProperties.Add(AppointmentProperties.Details);
            options.FetchProperties.Add(AppointmentProperties.StartTime);
            options.FetchProperties.Add(AppointmentProperties.ReplyTime);
            options.FetchProperties.Add(AppointmentProperties.Duration);
            options.FetchProperties.Add(AppointmentProperties.IsCanceledMeeting);
            options.FetchProperties.Add(AppointmentProperties.IsOrganizedByUser);
            options.FetchProperties.Add(AppointmentProperties.OnlineMeetingLink);
            options.FetchProperties.Add(AppointmentProperties.Organizer);
            options.FetchProperties.Add(AppointmentProperties.OriginalStartTime);
            options.FetchProperties.Add(AppointmentProperties.Sensitivity);


            var utcDateTime = new DateTime(DateTime.Now.Year, 01, 01, 00, 00, 0, DateTimeKind.Utc);
            IReadOnlyList<Appointment> apointments = await appointmentStore.FindAppointmentsAsync(utcDateTime, TimeSpan.FromDays(365), options);
            listLog.Items.Add("Total appointments from : " + utcDateTime + " to " + utcDateTime.AddDays(365).ToUniversalTime() + " is: " + apointments.Count.ToString());
            Save_Appointements(apointments);
        }

        private async void ShowMessageFinished(string input)
        {

            MessageDialog msgbox = new MessageDialog("All "+ input + " Extracted.");
            await msgbox.ShowAsync();


        }
        async void Save_Appointements(IReadOnlyList<Appointment> appointmentsList)
        {

            try
            {

                StorageFolder localFolderStorage = await globalStorageFolder.CreateFolderAsync("WPLogical_" + todayDateTime, CreationCollisionOption.OpenIfExists);
                StorageFile appointmentsHtmlFile = await localFolderStorage.CreateFileAsync("appointments_" + todayDateTime + ".html", CreationCollisionOption.OpenIfExists);

                var streamHTMLAppts = await appointmentsHtmlFile.OpenAsync(FileAccessMode.ReadWrite);

                var writerHTMLAppts = new DataWriter(streamHTMLAppts.GetOutputStreamAt(0));
                StringBuilder htmlBuilder = new StringBuilder();
                StringBuilder tempStringBuilderAp = new StringBuilder();

                //Generate HTML table headers
                htmlBuilder.AppendLine("<html>");
                htmlBuilder.AppendLine("<head>");
                htmlBuilder.AppendLine("<style type=\"text/css\">");
                htmlBuilder.AppendLine(".tg  {border-collapse:collapse;border-spacing:0;}");
                htmlBuilder.AppendLine(".tg td{font-family:Arial, sans-serif;font-size:14px;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;}");
                htmlBuilder.AppendLine(".tg th{font-family:Arial, sans-serif;font-size:14px;font-weight:normal;padding:10px 5px;border-style:solid;border-width:1px;overflow:hidden;word-break:normal;}");
                htmlBuilder.AppendLine(".tg .tg-qk49{font-weight:bold;background-color:#306f98;color:#ffffff;vertical-align:top}");
                htmlBuilder.AppendLine(".tg .tg-yw4l{vertical-align:top}");
                htmlBuilder.AppendLine("</style>");
                htmlBuilder.AppendLine("<meta charset=\"UTF-8\">");
                htmlBuilder.AppendLine("</head>");
                htmlBuilder.AppendLine("<body>");
                //Table details starts
                htmlBuilder.AppendLine("<table class=\"tg\">");
                htmlBuilder.AppendLine("  <tr>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Acquisition details</th>");
                htmlBuilder.AppendLine("  </tr>");
                htmlBuilder.AppendLine("  <tr>");
                htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + getdeviceInfo(appointmentsList.Count) + "</td>");
                htmlBuilder.AppendLine("  </tr>");
                //Table details ends
                htmlBuilder.AppendLine("</table>");
                htmlBuilder.AppendLine("<table class=\"tg\">");
                htmlBuilder.AppendLine("  <tr>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Subject</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Location</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Organizer</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Invitees</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Start Time (UTC)</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Original Start Time </th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Duration (in hours)</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Sensitivity</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Replay Time</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Is Organized by User?</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">Is Canceled?</th>");
                htmlBuilder.AppendLine("    <th class=\"tg-qk49\">More Details</th>");
                htmlBuilder.AppendLine("  </tr>");


                foreach (var appointment in appointmentsList)
                {

                    htmlBuilder.AppendLine("  <tr>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.Subject != null ? appointment.Subject : "N/A") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.Location != null ? appointment.Location : "N/A") + "</td>");
                    if (appointment.Organizer != null)
                    {
                        htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + "Address: " + (appointment.Organizer.Address != null ? appointment.Organizer.Address : "N/A") + " Name: " + (appointment.Organizer.DisplayName != null ? appointment.Organizer.DisplayName : "N/A") + "</td>");

                    }
                    else
                    {
                        htmlBuilder.AppendLine("    <td class=\"tg-yw4l\"> N/A </ td > ");
                    }

                    if (appointment.Invitees != null)
                    {
                        tempStringBuilderAp.Clear();
                        foreach (var item in appointment.Invitees)
                        {
                            tempStringBuilderAp.AppendLine("Adress: " + item.Address + "<br> Name: " + item.DisplayName + "<br> Response: " + item.Response + "<br> Role: " + item.Role + "<br>");
                        }
                    }
                    else
                    {
                        htmlBuilder.AppendLine("    <td class=\"tg-yw4l\"> N/A </ td > ");
                    }

                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (tempStringBuilderAp.ToString() != null ? tempStringBuilderAp.ToString() : "N/A") + "</td>");

                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.StartTime.ToUniversalTime().ToString() != null ? appointment.StartTime.ToUniversalTime().ToString() : "N/A") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.OriginalStartTime.ToString() != null ? appointment.OriginalStartTime.ToString() : "N/A") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.Duration.TotalHours) + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.Sensitivity.ToString() != null ? appointment.Sensitivity.ToString() : "N/A") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.ReplyTime.ToString() != null ? appointment.ReplyTime.ToString() : "N/A") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.IsOrganizedByUser.ToString() != null ? appointment.IsOrganizedByUser.ToString() : "N/A") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.IsCanceledMeeting.ToString() != null ? appointment.IsCanceledMeeting.ToString() : "N/A") + "</td>");
                    htmlBuilder.AppendLine("    <td class=\"tg-yw4l\">" + (appointment.Details.ToString() != null ? appointment.Details.ToString() : "N/A") + "</td>");
                    htmlBuilder.AppendLine("  </tr>");


                    listLog.Items.Add("Appointment saved: " + appointment.Subject);
                    listLog.SelectedIndex = listLog.Items.Count - 1;
                }
                htmlBuilder.AppendLine("</table>");
                htmlBuilder.AppendLine("</body> ");
                htmlBuilder.AppendLine("</html>");

                writerHTMLAppts.WriteString(htmlBuilder.ToString());
                await writerHTMLAppts.StoreAsync();
                await writerHTMLAppts.FlushAsync();

                listLog.Items.Add("Appointmemnts acquisition done and file saved.");
                listLog.SelectedIndex = listLog.Items.Count - 1;
                ShowMessageFinished("Appointments");
            }

            catch (Exception ee)
            {
                MessageDialog msgbox = new MessageDialog(ee.Message);
                await msgbox.ShowAsync();
            }
        }
    }
}
