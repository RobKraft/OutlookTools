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
		Dictionary<string, int> msgClasses = new Dictionary<string, int>();
		List<MyContact> contacts = new List<MyContact>();
		public Form1()
		{
			InitializeComponent();
		}
		private void GetMessageClasses()
		{
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
					contact.Phone = contactItem.PrimaryTelephoneNumber;
					contact.Address = contactItem.HomeAddress;
					contact.MessageClass = contactItem.MessageClass;
					contacts.Add(contact);
					UpdateProgressMessageOnUI("Loading " + contact.FirstName + " " + contact.LastName);
				}
			}
			UpdateUI(msgClasses, contacts);
		}
		public void UpdateProgressMessageOnUI(string message)
		{
			Action del = delegate
			{
				label1.Text = message;
			};
			Invoke(del);
			Application.DoEvents();
		}
		public void UpdateUI(Dictionary<string, int> msgClasses, List<MyContact> contacts)
		{
			Action del = delegate
			{
				foreach (var item in msgClasses)
				{
					dataGridView1.Rows.Add(item.Key, item.Value);
					comboBox1.Items.Add(item.Key);
					comboBox1.SelectedIndex = 0;
				}
				dataGridView2.DataSource = contacts;
			};
			Invoke(del);
			Application.DoEvents();
		}
		private void button1_Click(object sender, EventArgs e)
		{
			label1.Text = "Working...";
			button1.Enabled = false;
			button2.Enabled = false;
			string desiredMessageClass = comboBox1.SelectedItem.ToString();
			UpdateContactMessageClass(desiredMessageClass);
			ClearAndReloadData();
		}

		private void UpdateContactMessageClass(string desiredMessageClass, bool selected = false)
		{
			MyContact selectedContact = null;
			if (selected)
			{
				selectedContact = (MyContact)dataGridView2.SelectedRows[0].DataBoundItem;
			}
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
						if (selected)
						{
							if (selectedContact.EmailAddress == contactItem.Email1Address
								&& selectedContact.Phone == contactItem.PrimaryTelephoneNumber
								&& selectedContact.FirstName == contactItem.FirstName
								&& selectedContact.LastName == contactItem.LastName
								&& selectedContact.Address == contactItem.HomeAddress)
							{
								contactItem.MessageClass = desiredMessageClass;
								contactItem.Save();
							}
						}
						else
						{
							contactItem.MessageClass = desiredMessageClass;
							contactItem.Save();
						}
					}
				}
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			label1.Text = "Working...";
			button1.Enabled = false;
			button2.Enabled = false;
			string desiredMessageClass = comboBox1.SelectedItem.ToString();
			UpdateContactMessageClass(desiredMessageClass, true);
			ClearAndReloadData();
		}
		private void ClearAndReloadData()
		{
			dataGridView1.Rows.Clear();
			dataGridView2.DataSource = null;
			contacts.Clear();
			comboBox1.Items.Clear();
			msgClasses.Clear();
			GetMessageClasses();
			label1.Text = "";
			button1.Enabled = true;
			button2.Enabled = true;

		}

		private void Form1_Load(object sender, EventArgs e)
		{
			dataGridView1.Columns.Add("MessageClass", "Message Class");
			dataGridView1.Columns.Add("Number", "Number of Contacts with this Class");

			backgroundWorker1.RunWorkerAsync();
		}
		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			GetMessageClasses();
		}
		private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			//Application.DoEvents();

			label1.Text = "";
			button1.Enabled = true;
			button2.Enabled = true;
			comboBox1.Enabled = true;
		}
		private bool sortAscending = false;
		private void dataGridView2_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
		{
			if (e.ColumnIndex == 1)
			{
				if (sortAscending)
					dataGridView2.DataSource = contacts.OrderBy(i => i.FirstName).ToList();
				else
					dataGridView2.DataSource = contacts.OrderBy(i => i.FirstName).Reverse().ToList();
			}
			if (e.ColumnIndex == 2)
			{
				if (sortAscending)
					dataGridView2.DataSource = contacts.OrderBy(i => i.LastName).ToList();
				else
					dataGridView2.DataSource = contacts.OrderBy(i => i.LastName).Reverse().ToList();
			}
			if (e.ColumnIndex == 3)
			{
				if (sortAscending)
					dataGridView2.DataSource = contacts.OrderBy(i => i.EmailAddress).ToList();
				else
					dataGridView2.DataSource = contacts.OrderBy(i => i.EmailAddress).Reverse().ToList();
			}
			if (e.ColumnIndex == 4)
			{
				if (sortAscending)
					dataGridView2.DataSource = contacts.OrderBy(i => i.Phone).ToList();
				else
					dataGridView2.DataSource = contacts.OrderBy(i => i.Phone).Reverse().ToList();
			}
			if (e.ColumnIndex == 5)
			{
				if (sortAscending)
					dataGridView2.DataSource = contacts.OrderBy(i => i.Address).ToList();
				else
					dataGridView2.DataSource = contacts.OrderBy(i => i.Address).Reverse().ToList();
			}
			if (e.ColumnIndex == 0)
			{
				if (sortAscending)
					dataGridView2.DataSource = contacts.OrderBy(i => i.MessageClass).ToList();
				else
					dataGridView2.DataSource = contacts.OrderBy(i => i.MessageClass).Reverse().ToList();
			}
			sortAscending = !sortAscending;
		}

	}
	public class MyContact
	{
		public string MessageClass { get; set; }
		public string FirstName { get; set; }
		public string LastName { get; set; }
		public string EmailAddress { get; set; }
		public string Phone { get; set; }
		public string Address { get; set; }
	}
}
