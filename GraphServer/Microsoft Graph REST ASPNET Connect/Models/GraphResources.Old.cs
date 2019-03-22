using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Microsoft_Graph_REST_ASPNET_Connect.Models.old
{
    public class UserInfo
    {
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }
        public string id { get; set; }
        public List<object> businessPhones { get; set; }
        public string displayName { get; set; }
        public string givenName { get; set; }
        public object jobTitle { get; set; }
        public string mail { get; set; }
        public string mobilePhone { get; set; }
        public object officeLocation { get; set; }
        public string preferredLanguage { get; set; }
        public string surname { get; set; }
        public string userPrincipalName { get; set; }
        public string Address { get; set; }
        public string Manager { get; set; }
        public List<string> Skills { get; set; }
        
    }

    public class FileInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string SharingLink { get; set; }
    }

    public class Message
    {
        public string Subject { get; set; }
        public ItemBody Body { get; set; }
        public List<Recipient> ToRecipients { get; set; }
        public List<Attachment> Attachments { get; set; }
    }

    public class Recipient
    {
        public UserInfo EmailAddress { get; set; }
    }

    public class ItemBody
    {
        public string ContentType { get; set; }
        public string Content { get; set; }
    }

    public class MessageRequest
    {
        public Message Message { get; set; }
        public bool SaveToSentItems { get; set; }
    }

    public class Attachment
    {
        [JsonProperty(PropertyName = "@odata.type")]
        public string ODataType { get; set; }
        public byte[] ContentBytes { get; set; }
        public string Name { get; set; }
    }

    public class PermissionInfo
    {
        public SharingLinkInfo Link { get; set; }
    }

    public class SharingLinkInfo
    {
        public SharingLinkInfo(string type)
        {
            Type = type;
        }

        public string Type { get; set; }
        public string WebUrl { get; set; }
    }
    public class Root
    {
    }

    public class SiteCollection
    {
        public string hostname { get; set; }
    }

    public class SharePointSite
    {
        [JsonProperty("@odata.context")]
        public string OdataContext { get; set; }
        public DateTime createdDateTime { get; set; }
        public string description { get; set; }
        public string id { get; set; }
        public DateTime lastModifiedDateTime { get; set; }
        public string name { get; set; }
        public string webUrl { get; set; }
        public Root root { get; set; }
        public SiteCollection siteCollection { get; set; }
        public string displayName { get; set; }
        public List<SharePointList> Lists { get; set; }

    }

    public class User
    {
        public string displayName { get; set; }
        public string id { get; set; }
    }

    public class CreatedBy
    {
        public User user { get; set; }
    }

    public class ParentReference
    {
    }

    public class List
    {
        public bool contentTypesEnabled { get; set; }
        public bool hidden { get; set; }
        public string template { get; set; }
    }

    public class SharePointList
    {
        public string ODataEtag { get; set; }
        public CreatedBy createdBy { get; set; }
        public DateTime createdDateTime { get; set; }
        public string description { get; set; }
        public string eTag { get; set; }
        public string Id { get; set; }
        public DateTime lastModifiedDateTime { get; set; }
        public string Name { get; set; }
        public ParentReference parentReference { get; set; }
        public string webUrl { get; set; }
        public string displayName { get; set; }
        public List list { get; set; }
    }

    public class ResultLists
    {
        [JsonProperty("@odata.context")]
        public string ODataContext { get; set; }

        [JsonProperty("value")]
        public List<SharePointList> SharePointLists { get; set; }
    }

    public class LastModifiedBy
    {
        public User user { get; set; }
    }


    public class ContentType
    {
        public string id { get; set; }
    }

    public class FieldsCreated
    {
        public string Title { get; set; }
        public string DisplayName { get; set; }

        public string UPN { get; set; }

    }
    public class Fields
    {
        [JsonProperty("@odata.etag")]
        public string odataetag { get; set; }
        public string Title { get; set; }
        public string _x0064_jc8LookupId { get; set; }
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public string UPN { get; set; }
        public string id { get; set; }
        public string ContentType { get; set; }
        public DateTime Modified { get; set; }
        public DateTime Created { get; set; }
        public string AuthorLookupId { get; set; }
        public string EditorLookupId { get; set; }
        public string _UIVersionString { get; set; }
        public bool Attachments { get; set; }
        public string Edit { get; set; }
        public string LinkTitleNoMenu { get; set; }
        public string LinkTitle { get; set; }
        public string ItemChildCount { get; set; }
        public string FolderChildCount { get; set; }
        public string _ComplianceFlags { get; set; }
        public string _ComplianceTag { get; set; }
        public string _ComplianceTagWrittenTime { get; set; }
        public string _ComplianceTagUserId { get; set; }
    }
    public class ItemCreated
    {
        public FieldsCreated fields { get; set; }

    }
    public class Item
    {
        [JsonProperty("@odata.etag")]
        public string odataetag { get; set; }
        public CreatedBy createdBy { get; set; }
        public DateTime createdDateTime { get; set; }
        public string eTag { get; set; }
        public string id { get; set; }
        public LastModifiedBy lastModifiedBy { get; set; }
        public DateTime lastModifiedDateTime { get; set; }
        public ParentReference parentReference { get; set; }
        public string webUrl { get; set; }
        public ContentType contentType { get; set; }
        public string odatacontext { get; set; }
        public Fields fields { get; set; }
    }

    public class ListItems
    {
        [JsonProperty("@odata.context")]
        public string odatacontext { get; set; }

        [JsonProperty("value")]
        public List<Item> Items { get; set; }
    }

    public class Body
    {
        public string contentType { get; set; }
        public string content { get; set; }
    }

    public class Start
    {
        public DateTime dateTime { get; set; }
        public string timeZone { get; set; }
    }

    public class End
    {
        public DateTime dateTime { get; set; }
        public string timeZone { get; set; }
    }

    public class Location
    {
        public string displayName { get; set; }
    }

    public class EmailAddress
    {
        public string address { get; set; }
        public string name { get; set; }
    }

    public class Attendee
    {
        public EmailAddress emailAddress { get; set; }
        public string type { get; set; }
    }

    public class Event
    {
        public string subject { get; set; }
        public Body body { get; set; }
        public Start start { get; set; }
        public End end { get; set; }
        public Location location { get; set; }
        public List<Attendee> attendees { get; set; }
    }
}