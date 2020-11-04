using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Windows.Forms;
using OutLook = Microsoft.Office.Interop.Outlook;
//https://www.slipstick.com/developer/vba-set-existing-contacts-custom-form/
//https://www.c-sharpcorner.com/article/outlook-integration-in-C-Sharp/
namespace OutlookTools
{
	public partial class Form1 : Form
	{
		#region Properties
		OutLook._Application _outlookObj = new OutLook.Application();
		Dictionary<string, int> _msgClasses = new Dictionary<string, int>();
		List<IOutlookContact> _contacts = new List<IOutlookContact>();
		List<IOutlookTask> _tasks = new List<IOutlookTask>();
		OutlookMessageType _selectedOutlookType = OutlookMessageType.Contact;
		private List<IOutlookUIInfo> _propertyMaps = new List<IOutlookUIInfo>();
		private string userAccounts = "";
		#endregion

		#region PatternCompleted NeedsSimpleChanges
		private OutLook.MAPIFolder GetFolderForType(OutlookMessageType messageType)
		{
			OutLook.Stores stores = null;
			try
			{
				stores = _outlookObj.Session.Stores;
			}
			catch ( Exception ex)
			{
				//https://stackoverflow.com/questions/18373260/outlook-interop-exception
				userAccounts = ex.Message + " " + ex.InnerException + Environment.NewLine + Environment.NewLine;
			}
			foreach (OutLook.Store store in stores)
			{
				var c = store.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);
				//Console.WriteLine(c.Items.Count.ToString() + " contacts in " + store.DisplayName + " for " + store.FilePath);
				userAccounts += c.Items.Count.ToString() + " contacts in " + store.DisplayName + " for " + store.FilePath + Environment.NewLine;
				userAccounts += Environment.NewLine;
			}
			var accounts = _outlookObj.Session.Accounts;
			foreach (OutLook.Account account in accounts)
			{
				Console.WriteLine(account.DisplayName);
			}
			switch (messageType)
			{
				case OutlookMessageType.Contact:
					return (OutLook.MAPIFolder)_outlookObj.Session.GetDefaultFolder(
						OutLook.OlDefaultFolders.olFolderContacts);
				case OutlookMessageType.Task:
					return (OutLook.MAPIFolder)_outlookObj.Session.GetDefaultFolder(
						OutLook.OlDefaultFolders.olFolderTasks);
			}
			return null;
		}
		private void GetMessageClassesForType()
		{
			switch (_selectedOutlookType)
			{
				case OutlookMessageType.Contact:
					GetMessageClassesForContacts();
					break;
				case OutlookMessageType.Task:
					GetMessageClassesForTasks();
					break;
			}
		}
		private void ClearCurrentListOfItems()
		{
			switch (_selectedOutlookType)
			{
				case OutlookMessageType.Contact:
					_contacts.Clear();
					break;
				case OutlookMessageType.Task:
					_tasks.Clear();
					break;
			}
		}
		private bool sortAscending = false;
		private void dataGridView2_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
		{
			string sortColumn = GetSortColumn(e);

			switch (_selectedOutlookType)
			{
				case OutlookMessageType.Contact:
					_contacts = Extensions.OrderByMe(_contacts.AsQueryable(), sortColumn, sortAscending).ToList();
					break;
				case OutlookMessageType.Task:
					_tasks = Extensions.OrderByMe(_tasks.AsQueryable(), sortColumn, sortAscending).ToList();
					break;
			}
			SetDataSourceOfGrid();
			sortAscending = !sortAscending;
		}
		private object GetObjectsWeAreWorkingWith()
		{
			switch (_selectedOutlookType)
			{
				case OutlookMessageType.Contact:
					return _contacts;
				case OutlookMessageType.Task:
					return _tasks;
			}
			return null;
		}
		private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
		{
			bool loadData = false;
			switch (comboBox2.SelectedIndex)
			{
				case 0:
					_selectedOutlookType = OutlookMessageType.Contact;
					if (_contacts.Count == 0)
						loadData = true;
					break;
				case 1:
					_selectedOutlookType = OutlookMessageType.Task;
					if (_tasks.Count == 0)
						loadData = true;
					break;
			}
			if (loadData)
			{
				DisableUI();
				backgroundWorker1.RunWorkerAsync();
			}
			else
			{
				UpdateUI(_msgClasses);
			}
		}
		#endregion

		#region Code That Is Probably Fully Done
		public Form1()
		{
			InitializeComponent();
		}
		private void Form1_Load(object sender, EventArgs e)
		{
			SetupPropertyMaps();
			comboBox2.SelectedIndex = 1; //1 defaults to Task - this triggers an event!
			dataGridView1.Columns.Add("MessageClass", "Message Class");
			dataGridView1.Columns.Add("Number", "Number of " + _selectedOutlookType.ToString() + "s with this Class");
		}

		private void AddToListOfMessageClasses(string messageClass)
		{
			if (!_msgClasses.ContainsKey(messageClass))
			{
				_msgClasses.Add(messageClass, 0);
			}
			_msgClasses[messageClass]++;
		}

		public void UpdateProgressMessageOnUI(string message)
		{
			System.Action del = delegate
			{
				label1.Text = message;
			};
			Invoke(del);
			System.Windows.Forms.Application.DoEvents();
		}
		public void UpdateUI(Dictionary<string, int> msgClasses)
		{
			System.Action del = delegate
			{
				dataGridView1.Rows.Clear();
				comboBox1.Items.Clear();
				foreach (var item in msgClasses)
				{
					dataGridView1.Rows.Add(item.Key, item.Value);
					comboBox1.Items.Add(item.Key);
					comboBox1.SelectedIndex = 0;
				}
				SetDataSourceOfGrid();
				textBox2.Text = userAccounts;
			};
			Invoke(del);
			System.Windows.Forms.Application.DoEvents();
		}
		private void SetDataSourceOfGrid()
		{
			dataGridView2.DataSource = GetObjectsWeAreWorkingWith();
			SetColumnOrdersInGrid();
		}
		private void button1_Click(object sender, EventArgs e)
		{
			DoTheChange(false);
		}
		private void button2_Click(object sender, EventArgs e)
		{
			DoTheChange(true);
		}
		private void DoTheChange(bool forSelectedOnly)
		{
			DisableUI();
			string desiredMessageClass = comboBox1.SelectedItem.ToString();
			switch (_selectedOutlookType)
			{
				case OutlookMessageType.Contact:
					UpdateOutlookItemMessageClassContact(desiredMessageClass, forSelectedOnly);
					break;
				case OutlookMessageType.Task:
					UpdateOutlookItemMessageClassTask(desiredMessageClass, forSelectedOnly);
					break;
			}
			ClearAndReloadData();
		}
		private void DisableUI()
		{
			label1.Text = "Working...";
			button1.Enabled = false;
			button2.Enabled = false;
		}
		private void ClearAndReloadData()
		{
			dataGridView1.Rows.Clear();
			dataGridView2.DataSource = null;
			ClearCurrentListOfItems();
			comboBox1.Items.Clear();
			_msgClasses.Clear();
			GetMessageClassesForType();
			UpdateUI(_msgClasses);
			label1.Text = "";
			button1.Enabled = true;
			button2.Enabled = true;
		}
		private string GetSortColumn(DataGridViewCellMouseEventArgs e)
		{
			return _propertyMaps.FirstOrDefault(i => i.MessageType == _selectedOutlookType && i.Sequence == e.ColumnIndex).PropertyName;
		}

		private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
		{
			GetMessageClassesForType();
			UpdateUI(_msgClasses);
		}
		private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			label1.Text = "";
			button1.Enabled = true;
			button2.Enabled = true;
			comboBox1.Enabled = true;
			comboBox2.Enabled = true;
		}
		private void SetColumnOrdersInGrid()
		{
			foreach (var col in _propertyMaps.Where(i => i.MessageType == _selectedOutlookType))
			{
				dataGridView2.Columns[col.PropertyName].DisplayIndex = col.Sequence;
			}
		}
		#endregion

		#region Messy Code that needs work
		private void UpdateOutlookItemMessageClassContact(string desiredMessageClass, bool selected = false)
		{
			//This line is a safeguard to prevent changing all contacts to IPM.Task
			if (desiredMessageClass.Equals("IPM.Task", StringComparison.OrdinalIgnoreCase))
				return;
			IOutlookContact selectedItem = null;
			if (selected)
			{
				selectedItem = (IOutlookContact)dataGridView2.SelectedRows[0].DataBoundItem;
			}
			OutLook.MAPIFolder folderItems = GetFolderForType(_selectedOutlookType);
			foreach (var item in folderItems.Items)
			{
				//Some items are distribution lists, not contacts
				if (item is Microsoft.Office.Interop.Outlook._ContactItem)
				{
					Microsoft.Office.Interop.Outlook._ContactItem thisItem = (Microsoft.Office.Interop.Outlook._ContactItem)item;
					if (!thisItem.MessageClass.Equals(desiredMessageClass))
					{
						if (selected)
						{
							if (selectedItem.EntryID == thisItem.EntryID)
							{
								thisItem.MessageClass = desiredMessageClass;
								thisItem.Save();
							}
						}
						else
						{
							thisItem.MessageClass = desiredMessageClass;
							thisItem.Save();
						}
					}
				}
			}
		}

		private void UpdateOutlookItemMessageClassTask(string desiredMessageClass, bool selected = false)
		{
			//This line is a safeguard to prevent changing all tasks to IPM.Contact
			if (desiredMessageClass.Equals("IPM.Contact", StringComparison.OrdinalIgnoreCase))
				return;
			IOutlookTask selectedItem = null;
			if (selected)
			{
				selectedItem = (IOutlookTask)dataGridView2.SelectedRows[0].DataBoundItem;
			}
			OutLook.MAPIFolder folderItems = GetFolderForType(_selectedOutlookType);
			foreach (var item in folderItems.Items)
			{
				//Some items are distribution lists, not contacts
				if (item is Microsoft.Office.Interop.Outlook._TaskItem)
				{
					Microsoft.Office.Interop.Outlook._TaskItem thisItem = (Microsoft.Office.Interop.Outlook._TaskItem)item;
					if (!thisItem.MessageClass.Equals(desiredMessageClass))
					{
						if (selected)
						{
							if (selectedItem.EntryID == thisItem.EntryID)
							{
								thisItem.MessageClass = desiredMessageClass;
								thisItem.Save();
							}
						}
						else
						{
							thisItem.MessageClass = desiredMessageClass;
							thisItem.Save();
						}
					}
				}
			}
		}
		private void GetMessageClassesForContacts()
		{
			OutLook.MAPIFolder folderItems = GetFolderForType(_selectedOutlookType);

			foreach (var item in folderItems.Items)
			{
				if (item is Microsoft.Office.Interop.Outlook._ContactItem)
				{
					Microsoft.Office.Interop.Outlook._ContactItem thisItem = (Microsoft.Office.Interop.Outlook._ContactItem)item;
					AddToListOfMessageClasses(thisItem.MessageClass);

					IOutlookContact newItem = new OutlookContact();
					newItem.FirstName = (thisItem.FirstName == null) ? string.Empty : thisItem.FirstName;
					newItem.LastName = (thisItem.LastName == null) ? string.Empty : thisItem.LastName;
					newItem.Email1Address = thisItem.Email1Address;
					newItem.PrimaryTelephoneNumber = thisItem.PrimaryTelephoneNumber;
					newItem.HomeAddress = thisItem.HomeAddress;
					newItem.MessageClass = thisItem.MessageClass;
					newItem.MessageType = OutlookMessageType.Contact;
					newItem.EntryID = thisItem.EntryID;
					_contacts.Add(newItem);
					UpdateProgressMessageOnUI("Loading " + newItem.FirstName + " " + newItem.LastName);
				}
			}
		}
		private void GetMessageClassesForTasks()
		{
			OutLook.MAPIFolder folderItems = GetFolderForType(_selectedOutlookType);

			foreach (var item in folderItems.Items)
			{
				if (item is Microsoft.Office.Interop.Outlook._TaskItem)
				{
					Microsoft.Office.Interop.Outlook._TaskItem thisItem = (Microsoft.Office.Interop.Outlook._TaskItem)item;
					AddToListOfMessageClasses(thisItem.MessageClass);

					IOutlookTask newItem = new OutlookTask();
					newItem.Subject = (thisItem.Subject == null) ? string.Empty : thisItem.Subject;
					newItem.DueDate = thisItem.DueDate;
					newItem.Complete = thisItem.Complete;
					newItem.MessageClass = thisItem.MessageClass;
					newItem.MessageType = OutlookMessageType.Task;
					newItem.EntryID = thisItem.EntryID;
					_tasks.Add(newItem);
					UpdateProgressMessageOnUI("Loading " + newItem.Subject);
				}
			}
		}

		private void SetupPropertyMaps()
		{
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Contact, Sequence = 1, PropertyName = "MessageClass" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Contact, Sequence = 2, PropertyName = "FirstName" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Contact, Sequence = 3, PropertyName = "LastName" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Contact, Sequence = 4, PropertyName = "Email1Address" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Contact, Sequence = 5, PropertyName = "PrimaryTelephoneNumber" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Contact, Sequence = 6, PropertyName = "HomeAddress" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Contact, Sequence = 7, PropertyName = "EntryID" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Task, Sequence = 1, PropertyName = "MessageClass" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Task, Sequence = 2, PropertyName = "Subject" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Task, Sequence = 3, PropertyName = "DueDate" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Task, Sequence = 4, PropertyName = "Complete" });
			_propertyMaps.Add(new OutlookUIInfo() { MessageType = OutlookMessageType.Task, Sequence = 5, PropertyName = "EntryID" });
		}
		#endregion
	}
}