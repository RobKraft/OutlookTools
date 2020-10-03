﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookTools
{
    public class OutlookUIInfo: IOutlookUIInfo
	{
        public OutlookMessageType MessageType { get; set; }
        public string PropertyName { get; set; }
        public int Sequence { get; set; }
    }
    public interface IOutlookUIInfo
	{
        OutlookMessageType MessageType { get; set; }
        string PropertyName { get; set; }
        int Sequence { get; set; }
    }
    public class OutlookContact : IOutlookContact, IOutlookItem
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email1Address { get; set; }
        public string PrimaryTelephoneNumber { get; set; }
        public string HomeAddress { get; set; }
        public string MessageClass { get; set; }
        public OutlookMessageType MessageType { get; set; }
        public string EntryID { get; set; }
    }
    public class OutlookTask : IOutlookTask, IOutlookItem
    {
        public string Subject { get; set; }
        public DateTime DueDate { get; set; }
        public bool Complete { get; set; }
        public string MessageClass { get; set; }
        public OutlookMessageType MessageType { get; set; }
        public string EntryID { get; set; }
    }
    public interface IOutlookContact
    {
        OutlookMessageType MessageType { get; set; }
        string MessageClass { get; set; }
        string FirstName { get; set; }
        string LastName { get; set; }
        string Email1Address { get; set; }
        string PrimaryTelephoneNumber { get; set; }
        string HomeAddress { get; set; }
        string EntryID { get; set; }
    }
    public interface IOutlookTask
    {
        OutlookMessageType MessageType { get; set; }
        string MessageClass { get; set; }
        string Subject { get; set; }
        DateTime DueDate { get; set; }
        bool Complete { get; set; }
        string EntryID { get; set; }
    }
    public interface IOutlookItem
    {
        OutlookMessageType MessageType { get; set; }
        string MessageClass { get; set; }
        string EntryID { get; }
    }
    #region Enum MessageType
    /// <summary>
    /// The message types
    /// </summary>
    public enum OutlookMessageType
    {
        /// <summary>
        /// The message type is unknown
        /// </summary>
        Unknown,

        /// <summary>
        /// The message is a normal E-mail
        /// </summary>
        Email,

        /// <summary>
        /// Non-delivery report for a standard E-mail (REPORT.IPM.NOTE.NDR)
        /// </summary>
        EmailNonDeliveryReport,

        /// <summary>
        /// Delivery receipt for a standard E-mail (REPORT.IPM.NOTE.DR)
        /// </summary>
        EmailDeliveryReport,

        /// <summary>
        /// Delivery receipt for a delayed E-mail (REPORT.IPM.NOTE.DELAYED)
        /// </summary>
        EmailDelayedDeliveryReport,

        /// <summary>
        /// Read receipt for a standard E-mail (REPORT.IPM.NOTE.IPNRN)
        /// </summary>
        EmailReadReceipt,

        /// <summary>
        /// Non-read receipt for a standard E-mail (REPORT.IPM.NOTE.IPNNRN)
        /// </summary>
        EmailNonReadReceipt,

        /// <summary>
        /// The message in an E-mail that is encrypted and can also be signed (IPM.Note.SMIME)
        /// </summary>
        EmailEncryptedAndMaybeSigned,

        /// <summary>
        /// Non-delivery report for a Secure MIME (S/MIME) encrypted and opaque-signed E-mail (REPORT.IPM.NOTE.SMIME.NDR)
        /// </summary>
        EmailEncryptedAndMaybeSignedNonDelivery,

        /// <summary>
        /// Delivery report for a Secure MIME (S/MIME) encrypted and opaque-signed E-mail (REPORT.IPM.NOTE.SMIME.DR)
        /// </summary>
        EmailEncryptedAndMaybeSignedDelivery,

        /// <summary>
        /// The message is an E-mail that is clear signed (IPM.Note.SMIME.MultipartSigned)
        /// </summary>
        EmailClearSigned,

        /// <summary>
        /// The message is a secure read receipt for an E-mail (IPM.Note.Receipt.SMIME)
        /// </summary>
        EmailClearSignedReadReceipt,

        /// <summary>
        /// Non-delivery report for an S/MIME clear-signed E-mail (REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.NDR)
        /// </summary>
        EmailClearSignedNonDelivery,

        /// <summary>
        /// Delivery receipt for an S/MIME clear-signed E-mail (REPORT.IPM.NOTE.SMIME.MULTIPARTSIGNED.DR)
        /// </summary>
        EmailClearSignedDelivery,

        /// <summary>
        /// The message is an E-mail that is generared signed (IPM.Note.BMA.Stub)
        /// </summary>
        EmailBmaStub,

        /// <summary>
        /// The message is a short message service (IPM.Note.Mobile.SMS)
        /// </summary>
        EmailSms,

        /// <summary>
        /// The message is a Microsoft template (IPM.Note.Rules.OofTemplate.Microsoft)
        /// </summary>
        EmailTemplateMicrosoft,

        /// <summary>
        /// The message is an appointment (IPM.Appointment)
        /// </summary>
        Appointment,

        /// <summary>
        /// The message is a notification for an appointment (IPM.Notification.Meeting)
        /// </summary>
        AppointmentNotification,

        /// <summary>
        /// The message is a schedule for an appointment (IPM.Schedule.Meeting)
        /// </summary>
        AppointmentSchedule,

        /// <summary>
        /// The message is a request for an appointment (IPM.Schedule.Meeting.Request)
        /// </summary>
        AppointmentRequest,

        /// <summary>
        /// The message is a request for an appointment (REPORT.IPM.SCHEDULE.MEETING.REQUEST.NDR)
        /// </summary>
        AppointmentRequestNonDelivery,

        /// <summary>
        /// The message is a response to an appointment (IPM.Schedule.Response)
        /// </summary>
        AppointmentResponse,

        /// <summary>
        /// The message is a positive response to an appointment (IPM.Schedule.Resp.Pos)
        /// </summary>
        AppointmentResponsePositive,

        /// <summary>
        /// Non-delivery report for a positive meeting response (accept) (REPORT.IPM.SCHEDULE.MEETING.RESP.POS.NDR)
        /// </summary>
        AppointmentResponsePositiveNonDelivery,

        /// <summary>
        /// The message is a negative response to an appointment (IPM.Schedule.Resp.Neg)
        /// </summary>
        AppointmentResponseNegative,

        /// <summary>
        /// Non-delivery report for a negative meeting response (declinet) (REPORT.IPM.SCHEDULE.MEETING.RESP.NEG.NDR)
        /// </summary>
        AppointmentResponseNegativeNonDelivery,

        /// <summary>
        /// The message is a response to tentatively accept the meeting request (IPM.Schedule.Meeting.Resp.Tent)
        /// </summary>
        AppointmentResponseTentative,

        /// <summary>
        /// Non-delivery report for a Tentative meeting response (REPORT.IPM.SCHEDULE.MEETING.RESP.TENT.NDR)
        /// </summary>
        AppointmentResponseTentativeNonDelivery,

        /// <summary>
        /// The message is a cancelation an appointment (IPM.Schedule.Meeting.Canceled)
        /// </summary>
        AppointmentResponseCanceled,

        /// <summary>
        /// Non-delivery report for a cancelled meeting notification (REPORT.IPM.SCHEDULE.MEETING.CANCELED.NDR)
        /// </summary>
        AppointmentResponseCanceledNonDelivery,

        /// <summary>
        /// The message is a contact card (IPM.Contact)
        /// </summary>
        Contact,

        /// <summary>
        /// The message is a task (IPM.Task)
        /// </summary>
        Task,

        /// <summary>
        /// The message is a task request accept (IPM.TaskRequest.Accept)
        /// </summary>
        TaskRequestAccept,

        /// <summary>
        /// The message is a task request decline (IPM.TaskRequest.Decline)
        /// </summary>
        TaskRequestDecline,

        /// <summary>
        /// The message is a task request update (IPM.TaskRequest.Update)
        /// </summary>
        TaskRequestUpdate,

        /// <summary>
        /// The message is a sticky note (IPM.StickyNote)
        /// </summary>
        StickyNote,

        /// <summary>
        /// The message is Cisco Unity Voice message (IPM.Note.Custom.Cisco.Unity.Voice)
        /// </summary>
        CiscoUnityVoiceMessage,

        /// <summary>
        /// IPM.NOTE.RIGHTFAX.ADV
        /// </summary>
        RightFaxAdv,

        /// <summary>
        /// The message is Skype for Business missed message (IPM.Note.Microsoft.Missed)
        /// </summary>
        SkypeForBusinessMissedMessage,

        /// <summary>
        /// The message is a Skype for Business conversation (IPM.Note.Microsoft.Conversation)
        /// </summary>
        SkypeForBusinessConversation
    }
    #endregion

}
