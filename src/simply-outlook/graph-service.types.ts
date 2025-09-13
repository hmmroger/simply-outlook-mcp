import { BodyType, Event, EventType, DateTimeTimeZone, Message } from "@microsoft/microsoft-graph-types";

export type DateTimeRange = {
  startDateTime: string;
  endDateTime?: string;
};

export type NameWithEmail = {
  name: string;
  email: string;
};

export interface ItemBody {
  content: string;
  contentType: BodyType;
}

export type CalendarEventData = Required<Pick<Event, "id">> &
  Pick<
    Event,
    | "createdDateTime"
    | "subject"
    | "categories"
    | "iCalUId"
    | "hasAttachments"
    | "showAs"
    | "isOnlineMeeting"
    | "isOrganizer"
    | "attendees"
    | "onlineMeeting"
    | "organizer"
  > & {
    type: EventType;
    start: DateTimeTimeZone;
    end?: DateTimeTimeZone;
    body?: ItemBody;
  };

export type MailMessageData = Required<Pick<Message, "id">> &
  Pick<
    Message,
    | "createdDateTime"
    | "sentDateTime"
    | "subject"
    | "bodyPreview"
    | "importance"
    | "sender"
    | "from"
    | "toRecipients"
    | "replyTo"
    | "parentFolderId"
    | "isRead"
    | "isDraft"
    | "categories"
  > & {
    receivedDateTime: string;
    body?: ItemBody;
  };

export type MailFolderData = {
  id: string;
  displayName: string;
  wellKnownName?: string;
};
