unit uOutlookPR;

interface

Const
//This is a full List of email properties (Possible candidates in BOLD):

//  PR_SEARCH_KEY = PT_BINARY or ($300B shl 16);
//  PR_SEARCH_KEY : String = 'http://schemas.microsoft.com/mapi/proptag/0x300B0102';


//  PT_BINARY = ULONG(258); { ($0102) Uninterpreted (counted byte array) }
//  (*PR_SEARCH_KEY
//   Contains a binary-comparable key that identifies correlated objects for a search.
//   This property provides a trace for related objects, such as message copies, and facilitates finding unwanted occurrences, such as duplicate recipients.
//   MAPI uses specific rules for constructing search keys for message recipients. The search key is formed by concatenating the address type (in uppercase characters), the colon character ':', the e-mail address in canonical form, and the terminating null character. Canonical form here means that case-sensitive addresses appear in the correct case, and addresses that are not case-sensitive are converted to uppercase. This is important in preserving correlations among messages.
//   For message objects, this property is available through the IMAPIProp::GetProps method immediately following message creation. For other objects, it is available following the first call to the IMAPIProp::SaveChanges method. Because this property is changeable, it is unreliable to obtain it through GetProps until a SaveChanges call has committed any values set or changed by the IMAPIProp::SetProps method.
//   For profiles, MAPI also furnishes a hard-coded profile section named MUID_PROFILE_INSTANCE, with this property as its single property. This key is guaranteed to be unique among all profiles ever created, and can be more reliable than the PR_PROFILE_NAME (PidTagProfileName) property, which can be, for example,
//  *)
//
//  PR_PARENT_ENTRYID = PT_BINARY or ($0E09 shl 16);
//  (*PR_PARENT_ENTRYID
//   Contains the entry identifier of the folder that contains a folder or message.
//   This property is computed by message stores for all folders and messages.
//   For a message store root folder, this property contains the folder's own entry identifier.
//   PR_PARENT_DISPLAY  (PidTagParentDisplay) and this property are not related to each other. They belong to entirely different contexts.
//  *)
//
//  PR_STORE_ENTRYID = PT_BINARY or ($0FFB shl 16);
//  (*PR_STORE_ENTRYID
//   Contains the unique entry identifier of the message store where an object resides.
//   This property is used to open a message store with the IMAPISession::OpenMsgStore method. It is also used to open any object that is owned by the message store.
//   For a message store, this property is identical to the store's own PR_ENTRYID (PidTagEntryId) property. A client application can compare the two properties using the IMAPISession::CompareEntryIDs method.
//  *)


  PR_MESSAGE_CLASS                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x001A001E'; {PT_TSTRING}
  PR_MESSAGE_CLASS_W                   : String = 'http://schemas.microsoft.com/mapi/proptag/0x001A001E'; {PT_UNICODE}
  PR_MESSAGE_CLASS_A                   : String = 'http://schemas.microsoft.com/mapi/proptag/0x001A001E'; {PT_TSTRING}
  PR_SUBJECT                           : String = 'http://schemas.microsoft.com/mapi/proptag/0x0037001E'; {PT_TSTRING}
  PR_SUBJECT_W                         : String = 'http://schemas.microsoft.com/mapi/proptag/0x0037001E'; {PT_UNICODE}
  PR_SUBJECT_A                         : String = 'http://schemas.microsoft.com/mapi/proptag/0x0037001E'; {PT_TSTRING}
  PR_CLIENT_SUBMIT_TIME                : String = 'http://schemas.microsoft.com/mapi/proptag/0x00390040'; {PT_SYSTIME}
  PR_SENT_REPRESENTING_SEARCH_KEY      : String = 'http://schemas.microsoft.com/mapi/proptag/0x003B0102'; {PT_BINARY}
  PR_SUBJECT_PREFIX                    : String = 'http://schemas.microsoft.com/mapi/proptag/0x003D001E'; {PT_TSTRING}
  PR_SUBJECT_PREFIX_W                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x003D001E'; {PT_UNICODE}
  PR_SUBJECT_PREFIX_A                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x003D001E'; {PT_TSTRING}
  PR_RECEIVED_BY_ENTRYID               : String = 'http://schemas.microsoft.com/mapi/proptag/0x003F0102'; {PT_BINARY}
  PR_RECEIVED_BY_NAME                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x0040001E'; {PT_TSTRING}
  PR_RECEIVED_BY_NAME_W                : String = 'http://schemas.microsoft.com/mapi/proptag/0x0040001E'; {PT_UNICODE}
  PR_RECEIVED_BY_NAME_A                : String = 'http://schemas.microsoft.com/mapi/proptag/0x0040001E'; {PT_TSTRING}
  PR_SENT_REPRESENTING_ENTRYID         : String = 'http://schemas.microsoft.com/mapi/proptag/0x00410102'; {PT_BINARY}
  PR_SENT_REPRESENTING_NAME            : String = 'http://schemas.microsoft.com/mapi/proptag/0x0042001E'; {PT_TSTRING}
  PR_SENT_REPRESENTING_NAME_W          : String = 'http://schemas.microsoft.com/mapi/proptag/0x0042001E'; {PT_UNICODE}
  PR_SENT_REPRESENTING_NAME_A          : String = 'http://schemas.microsoft.com/mapi/proptag/0x0042001E'; {PT_TSTRING}
  PR_REPLY_RECIPIENT_ENTRIES           : String = 'http://schemas.microsoft.com/mapi/proptag/0x004F0102'; {PT_BINARY}
  PR_REPLY_RECIPIENT_NAMES             : String = 'http://schemas.microsoft.com/mapi/proptag/0x0050001E'; {PT_TSTRING}
  PR_REPLY_RECIPIENT_NAMES_W           : String = 'http://schemas.microsoft.com/mapi/proptag/0x0050001E'; {PT_UNICODE}
  PR_REPLY_RECIPIENT_NAMES_A           : String = 'http://schemas.microsoft.com/mapi/proptag/0x0050001E'; {PT_TSTRING}
  PR_RECEIVED_BY_SEARCH_KEY            : String = 'http://schemas.microsoft.com/mapi/proptag/0x00510102'; {PT_BINARY}
  PR_SENT_REPRESENTING_ADDRTYPE        : String = 'http://schemas.microsoft.com/mapi/proptag/0x0064001E'; {PT_TSTRING}
  PR_SENT_REPRESENTING_ADDRTYPE_W      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0064001E'; {PT_UNICODE}
  PR_SENT_REPRESENTING_ADDRTYPE_A      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0064001E'; {PT_TSTRING}
  PR_SENT_REPRESENTING_EMAIL_ADDRESS   : String = 'http://schemas.microsoft.com/mapi/proptag/0x0065001E'; {PT_TSTRING}
  PR_SENT_REPRESENTING_EMAIL_ADDRESS_W : String = 'http://schemas.microsoft.com/mapi/proptag/0x0065001E'; {PT_UNICODE}
  PR_SENT_REPRESENTING_EMAIL_ADDRESS_A : String = 'http://schemas.microsoft.com/mapi/proptag/0x0065001E'; {PT_TSTRING}
  PR_CONVERSATION_TOPIC                : String = 'http://schemas.microsoft.com/mapi/proptag/0x0070001E'; {PT_TSTRING}
  PR_CONVERSATION_TOPIC_W              : String = 'http://schemas.microsoft.com/mapi/proptag/0x0070001E'; {PT_UNICODE}
  PR_CONVERSATION_TOPIC_A              : String = 'http://schemas.microsoft.com/mapi/proptag/0x0070001E'; {PT_TSTRING}
  PR_CONVERSATION_INDEX                : String = 'http://schemas.microsoft.com/mapi/proptag/0x00710102'; {PT_BINARY}
  PR_RECEIVED_BY_ADDRTYPE              : String = 'http://schemas.microsoft.com/mapi/proptag/0x0075001E'; {PT_TSTRING}
  PR_RECEIVED_BY_ADDRTYPE_W            : String = 'http://schemas.microsoft.com/mapi/proptag/0x0075001E'; {PT_UNICODE}
  PR_RECEIVED_BY_ADDRTYPE_A            : String = 'http://schemas.microsoft.com/mapi/proptag/0x0075001E'; {PT_TSTRING}
  PR_RECEIVED_BY_EMAIL_ADDRESS         : String = 'http://schemas.microsoft.com/mapi/proptag/0x0076001E'; {PT_TSTRING}
  PR_RECEIVED_BY_EMAIL_ADDRESS_W       : String = 'http://schemas.microsoft.com/mapi/proptag/0x0076001E'; {PT_UNICODE}
  PR_RECEIVED_BY_EMAIL_ADDRESS_A       : String = 'http://schemas.microsoft.com/mapi/proptag/0x0076001E'; {PT_TSTRING}
  PR_TRANSPORT_MESSAGE_HEADERS         : String = 'http://schemas.microsoft.com/mapi/proptag/0x007D001E'; {PT_TSTRING}
  PR_TRANSPORT_MESSAGE_HEADERS_W       : String = 'http://schemas.microsoft.com/mapi/proptag/0x007D001E'; {PT_UNICODE}
  PR_TRANSPORT_MESSAGE_HEADERS_A       : String = 'http://schemas.microsoft.com/mapi/proptag/0x007D001E'; {PT_TSTRING}
  PR_SENDER_ENTRYID                    : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C190102'; {PT_BINARY}
  PR_SENDER_NAME                       : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1A001E'; {PT_TSTRING}
  PR_SENDER_NAME_W                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1A001E'; {PT_UNICODE}
  PR_SENDER_NAME_A                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1A001E'; {PT_TSTRING}
  PR_SENDER_SEARCH_KEY                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1D0102'; {PT_BINARY}
  PR_SENDER_ADDRTYPE                   : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1E001E'; {PT_TSTRING}
  PR_SENDER_ADDRTYPE_W                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1E001E'; {PT_UNICODE}
  PR_SENDER_ADDRTYPE_A                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1E001E'; {PT_TSTRING}
  PR_SENDER_EMAIL_ADDRESS              : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1F001E'; {PT_TSTRING}
  PR_SENDER_EMAIL_ADDRESS_W            : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1F001E'; {PT_UNICODE}
  PR_SENDER_EMAIL_ADDRESS_A            : String = 'http://schemas.microsoft.com/mapi/proptag/0x0C1F001E'; {PT_TSTRING}
  PR_DISPLAY_BCC                       : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E02001E'; {PT_TSTRING}
  PR_DISPLAY_BCC_W                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E02001E'; {PT_UNICODE}
  PR_DISPLAY_BCC_A                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E02001E'; {PT_TSTRING}
  PR_DISPLAY_CC                        : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E03001E'; {PT_TSTRING}
  PR_DISPLAY_CC_W                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E03001E'; {PT_UNICODE}
  PR_DISPLAY_CC_A                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E03001E'; {PT_TSTRING}
  PR_DISPLAY_TO                        : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E04001E'; {PT_TSTRING}
  PR_DISPLAY_TO_W                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E04001E'; {PT_UNICODE}
  PR_DISPLAY_TO_A                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E04001E'; {PT_TSTRING}
  PR_MESSAGE_DELIVERY_TIME             : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E060040'; {PT_SYSTIME}
  PR_MESSAGE_FLAGS                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E070003'; {PT_LONG}
  PR_MESSAGE_SIZE                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E080003'; {PT_LONG}
  PR_PARENT_ENTRYID                    : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E090102'; {PT_BINARY}
  PR_MESSAGE_RECIPIENTS                : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E12000D'; {PT_OBJECT}
  PR_MESSAGE_ATTACHMENTS               : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E13000D'; {PT_OBJECT}
  PR_HASATTACH                         : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1B000B'; {PT_BOOLEAN}
  PR_NORMALIZED_SUBJECT                : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1D001E'; {PT_TSTRING}
  PR_NORMALIZED_SUBJECT_W              : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1D001E'; {PT_UNICODE}
  PR_NORMALIZED_SUBJECT_A              : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1D001E'; {PT_TSTRING}
  PR_RTF_IN_SYNC                       : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E1F000B'; {PT_BOOLEAN}
  PR_PRIMARY_SEND_ACCT                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E28001E'; {PT_TSTRING}
  PR_PRIMARY_SEND_ACCT_W               : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E28001E'; {PT_UNICODE}
  PR_PRIMARY_SEND_ACCT_A               : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E28001E'; {PT_TSTRING}
  PR_NEXT_SEND_ACCT                    : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E29001E'; {PT_TSTRING}
  PR_NEXT_SEND_ACCT_W                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E29001E'; {PT_UNICODE}
  PR_NEXT_SEND_ACCT_A                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E29001E'; {PT_TSTRING}
  PR_ACCESS                            : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FF40003'; {PT_LONG}
  PR_ACCESS_LEVEL                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FF70003'; {PT_LONG}
  PR_MAPPING_SIGNATURE                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FF80102'; {PT_BINARY}
  PR_RECORD_KEY                        : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FF90102'; {PT_BINARY}
  PR_STORE_RECORD_KEY                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FFA0102'; {PT_BINARY}
  PR_STORE_ENTRYID                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FFB0102'; {PT_BINARY}
  PR_OBJECT_TYPE                       : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FFE0003'; {PT_LONG}
  PR_ENTRYID                           : String = 'http://schemas.microsoft.com/mapi/proptag/0x0FFF0102'; {PT_BINARY}
  PR_BODY                              : String = 'http://schemas.microsoft.com/mapi/proptag/0x1000001F'; {PT_TSTRING}
  PR_BODY_W                            : String = 'http://schemas.microsoft.com/mapi/proptag/0x1000001F'; {PT_UNICODE}
  PR_BODY_A                            : String = 'http://schemas.microsoft.com/mapi/proptag/0x1000001F'; {PT_TSTRING}
  PR_RTF_COMPRESSED                    : String = 'http://schemas.microsoft.com/mapi/proptag/0x10090102'; {PT_BINARY}
  PR_HTML                              : String = 'http://schemas.microsoft.com/mapi/proptag/0x10130102'; {PT_BINARY}
  PR_INTERNET_MESSAGE_ID               : String = 'http://schemas.microsoft.com/mapi/proptag/0x1035001E'; {PT_TSTRING}
  PR_INTERNET_MESSAGE_ID_W             : String = 'http://schemas.microsoft.com/mapi/proptag/0x1035001E'; {PT_UNICODE}
  PR_INTERNET_MESSAGE_ID_A             : String = 'http://schemas.microsoft.com/mapi/proptag/0x1035001E'; {PT_TSTRING}
  PR_LIST_UNSUBSCRIBE                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x1045001E'; {PT_TSTRING}
  PR_LIST_UNSUBSCRIBE_W                : String = 'http://schemas.microsoft.com/mapi/proptag/0x1045001E'; {PT_UNICODE}
  PR_LIST_UNSUBSCRIBE_A                : String = 'http://schemas.microsoft.com/mapi/proptag/0x1045001E'; {PT_TSTRING}
  PR_CREATION_TIME                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x30070040'; {PT_SYSTIME}
  PR_LAST_MODIFICATION_TIME            : String = 'http://schemas.microsoft.com/mapi/proptag/0x30080040'; {PT_SYSTIME}
  PR_SEARCH_KEY                        : String = 'http://schemas.microsoft.com/mapi/proptag/0x300B0102'; {PT_BINARY}
  PR_STORE_SUPPORT_MASK                : String = 'http://schemas.microsoft.com/mapi/proptag/0x340D0003'; {PT_LONG}
  PR_MDB_PROVIDER                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x34140102'; {PT_BINARY}
  PR_INTERNET_CPID                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x3FDE0003'; {PT_LONG}
  PR_MSG_EDITOR_FORMAT                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x59090003'; {PT_LONG}
  PR_NATIVE_BODY_INFO                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x10160003'; {PT_LONG}
  PR_MESSAGE_CODEPAGE                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x3FFD0003'; {PT_LONG}

  PR_ATTACH_ENCODING                   : String = 'http://schemas.microsoft.com/mapi/proptag/0x37020102'; {PT_BINARY}
  PR_ATTACH_ADDITIONAL_INFO            : String = 'http://schemas.microsoft.com/mapi/proptag/0x370F0102'; {PT_BINARY}
  PR_ATTACH_CONTENT_LOCATION           : String = 'http://schemas.microsoft.com/mapi/proptag/0x3713001F'; {PT_TSTRING}
  PR_ATTACH_CONTENT_LOCATION_W         : String = 'http://schemas.microsoft.com/mapi/proptag/0x3713001F'; {PT_UNICODE}
  PR_ATTACH_CONTENT_LOCATION_A         : String = 'http://schemas.microsoft.com/mapi/proptag/0x3713001F'; {PT_TSTRING}
  PR_ATTACH_METHOD                     : String = 'http://schemas.microsoft.com/mapi/proptag/0x37050003'; {PT_LONG}
  PR_ATTACH_DATA_BIN                   : String = 'http://schemas.microsoft.com/mapi/proptag/0x37010102'; {PT_BINARY}
  PR_ATTACH_FLAGS                      : String = 'http://schemas.microsoft.com/mapi/proptag/0x37140003'; {PT_LONG}
  PR_ATTACHMENT_FLAGS                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x37140003'; {PT_LONG}
  PR_ATTACHMENT_LINKID                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x7FFA0003'; {PT_LONG}
  PR_ATTACH_SIZE                       : String = 'http://schemas.microsoft.com/mapi/proptag/0x0E200003'; {PT_LONG}
  PR_ATTACH_MIME_TAG                   : String = 'http://schemas.microsoft.com/mapi/proptag/0x370E001F'; {PT_TSTRING}
  PR_ATTACH_MIME_TAG_W                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x370E001F'; {PT_UNICODE}
  PR_ATTACH_MIME_TAG_A                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x370E001F'; {PT_TSTRING}
  PR_ATTACH_PAYLOAD_CLASS              : String = 'http://schemas.microsoft.com/mapi/proptag/0x371A001F'; {PT_TSTRING}
  PR_ATTACH_PAYLOAD_CLASS_W            : String = 'http://schemas.microsoft.com/mapi/proptag/0x371A001F'; {PT_UNICODE}
  PR_ATTACH_PAYLOAD_CLASS_A            : String = 'http://schemas.microsoft.com/mapi/proptag/0x371A001F'; {PT_TSTRING}
  PR_ATTACH_RENDERING                  : String = 'http://schemas.microsoft.com/mapi/proptag/0x37090102'; {PT_BINARY}
  PR_ATTACH_CONTENT_ID                 : String = 'http://schemas.microsoft.com/mapi/proptag/0x3712001E'; {PT_TSTRING}
  PR_ATTACH_CONTENT_ID_W               : String = 'http://schemas.microsoft.com/mapi/proptag/0x3712001E'; {PT_UNICODE}
  PR_ATTACH_CONTENT_ID_A               : String = 'http://schemas.microsoft.com/mapi/proptag/0x3712001E'; {PT_TSTRING}

Const

  SideEffects                          : String = 'http://schemas.microsoft.com/mapi/proptag/0x80050003';
  InetAcctID                           : String = 'http://schemas.microsoft.com/mapi/proptag/0x802A001E';
  InetAcctName                         : String = 'http://schemas.microsoft.com/mapi/proptag/0x804F001E';
  RemoteEID                            : String = 'http://schemas.microsoft.com/mapi/proptag/0x80660102';
  x_rcpt_to                            : String = 'http://schemas.microsoft.com/mapi/proptag/0x80AD001E';


implementation

end.
