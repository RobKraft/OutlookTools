using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutLook = Microsoft.Office.Interop.Outlook;
//https://www.slipstick.com/developer/vba-set-existing-contacts-custom-form/
//https://www.c-sharpcorner.com/article/outlook-integration-in-C-Sharp/
namespace OutlookTools
{
	public partial class Form1 : Form
	{
		OutLook._Application outlookObj = new OutLook.Application();
		public Form1()
		{
			InitializeComponent();
			GetMessageClasses();
		}
		private List<MyContact> GetMessageClasses()
		{
			var contacts = new List<MyContact>();
			Dictionary<string, int> msgClasses = new Dictionary<string, int>();
			OutLook.MAPIFolder fldContacts =
				(OutLook.MAPIFolder)outlookObj.Session.GetDefaultFolder(
					OutLook.OlDefaultFolders.olFolderContacts);
			foreach (var item in fldContacts.Items)
			{
				if (item is Microsoft.Office.Interop.Outlook._ContactItem)
				{
					Microsoft.Office.Interop.Outlook._ContactItem contactItem = (Microsoft.Office.Interop.Outlook._ContactItem)item;
					if (!msgClasses.ContainsKey(contactItem.MessageClass))
					{
						msgClasses.Add(contactItem.MessageClass, 0);
					}
					msgClasses[contactItem.MessageClass]++;

					MyContact contact = new MyContact();
					contact.FirstName = (contactItem.FirstName == null) ? string.Empty : contactItem.FirstName;
					contact.LastName = (contactItem.LastName == null) ? string.Empty : contactItem.LastName;
					contact.EmailAddress = contactItem.Email1Address;
					contact.Phone = contactItem.Business2TelephoneNumber;
					contact.Address = contactItem.BusinessAddress;
					contact.MessageClass = contactItem.MessageClass;
					contacts.Add(contact);
				}
			}
			dataGridView1.Columns.Add("MessageClass", "Message Class");
			dataGridView1.Columns.Add("Number", "Number of Contacts with this Class");
			foreach (var item in msgClasses)
			{
				dataGridView1.Rows.Add(item.Key, item.Value);
				comboBox1.Items.Add(item.Key);
				comboBox1.SelectedIndex = 0;
			}
			return contacts;
		}
		private void button1_Click(object sender, EventArgs e)
		{
			string desiredMessageClass = comboBox1.SelectedItem.ToString();
			OutLook.MAPIFolder fldContacts =
				(OutLook.MAPIFolder)outlookObj.Session.GetDefaultFolder(
					OutLook.OlDefaultFolders.olFolderContacts);
			foreach (var item in fldContacts.Items)
			{
				//Some items are distribution lists, not contacts
				if (item is Microsoft.Office.Interop.Outlook._ContactItem)
				{
					Microsoft.Office.Interop.Outlook._ContactItem contactItem = (Microsoft.Office.Interop.Outlook._ContactItem)item;
					if (!contactItem.MessageClass.Equals(desiredMessageClass))
					{
						contactItem.MessageClass = desiredMessageClass;
						contactItem.Save();
					}
				}
			}
		}
	}
	public class MyContact
	{
		public string FirstName { get; set; }
		public string LastName { get; set; }
		public string EmailAddress { get; set; }
		public string Phone { get; set; }
		public string Address { get; set; }
		public string MessageClass { get; set; }
	}
}
