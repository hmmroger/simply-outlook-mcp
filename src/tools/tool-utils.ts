import { Recipient } from "@microsoft/microsoft-graph-types";
import { CalendarEventData, MailMessageData, NameWithEmail } from "../simply-outlook/graph-service.types.js";

const MAX_MAIL_MESSAGE_RECIPIENTS = 10;

/**
 * Creates a tool result object with text content.
 *
 * @param texts - Array of strings to be joined with newlines
 * @param isError - Optional flag indicating if this is an error result
 * @returns Object with content array containing text and optional isError flag
 */
export const textToolResult = (texts: string[], isError?: boolean) => {
  const text = texts.join("\n");
  return {
    content: [
      {
        type: "text" as const,
        text,
      },
    ],
    isError,
  };
};

/**
 * Creates an error tool result from an exception or error object.
 *
 * @param error - The error object or unknown error to extract message from
 * @param fallbackMessage - Message to use if no error message can be extracted
 * @returns Error tool result object with isError flag set to true
 */
export const getErrorToolResult = (error: unknown, fallbackMessage: string) => {
  const exceptionError = (error as Error).message;
  const errorMessage = exceptionError ? exceptionError : fallbackMessage;
  return textToolResult([errorMessage], true);
};

/**
 * Converts a UTC datetime string to local datetime string representation.
 *
 * @param dateTime - UTC datetime string, optionally with 'Z' suffix
 * @returns Local datetime string representation, or empty string if input is null/undefined
 */
export const utcDateTimeToLocal = (dateTime?: string | null) => {
  if (!dateTime) {
    return "";
  }

  const utcDateTime = dateTime.endsWith("Z") ? dateTime : `${dateTime}Z`;
  return new Date(utcDateTime).toString();
};

/**
 * Gets the current UTC datetime rounded to the nearest hour (minutes, seconds, milliseconds set to 0).
 *
 * @returns ISO string representation of current datetime rounded to the hour
 */
export const getUtcDateTimeToTheHour = () => {
  const now = new Date();
  now.setMinutes(0, 0, 0);
  return now.toISOString();
};

/**
 * Gets the current UTC datetime rounded to the start of the day (hours, minutes, seconds, milliseconds set to 0).
 *
 * @returns ISO string representation of current datetime rounded to the start of the day
 */
export const getUtcDateTimeToTheDay = () => {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return now.toISOString();
};

/**
 * Converts calendar event data from Microsoft Graph API to a standardized result format.
 *
 * @param eventData - Raw calendar event data from Microsoft Graph API
 * @returns Formatted calendar event result with localized datetime strings and simplified structure
 */
export const toCalendarEventResult = (eventData: CalendarEventData) => {
  const subject = eventData.subject || "No subject";
  const eventStartDateTime = utcDateTimeToLocal(eventData.start.dateTime);
  const eventEndDateTime = eventData.end ? utcDateTimeToLocal(eventData.end.dateTime) : undefined;
  const organizer = toNameWithEmail(eventData.organizer);
  const attendees = eventData.attendees?.map((attendee) => toNameWithEmail(attendee));

  const result = {
    id: eventData.id,
    type: eventData.type,
    subject,
    content: eventData.body?.content,
    startDateTime: eventStartDateTime,
    endDateTime: eventEndDateTime,
    organizer,
    isOnlineMeeting: !!eventData.isOnlineMeeting,
    isOrganizedByCurrentUser: !!eventData.isOrganizer,
    attendees: attendees?.length && attendees,
  };

  return result;
};

/**
 * Converts mail message data from Microsoft Graph API to a standardized result format.
 *
 * @param messageData - Raw mail message data from Microsoft Graph API
 * @returns Formatted mail message result with localized datetime strings and limited recipients
 */
export const toMailMessageResult = (messageData: MailMessageData) => {
  const subject = messageData.subject || "No subject";
  const receivedDateTime = utcDateTimeToLocal(messageData.receivedDateTime);
  const senderFallback = toNameWithEmail(messageData.sender);
  const from = toNameWithEmail(messageData.from, senderFallback);
  const toRecipients = messageData.toRecipients?.slice(0, MAX_MAIL_MESSAGE_RECIPIENTS).map((recipient) => toNameWithEmail(recipient));
  const contentPreview = messageData.bodyPreview ? messageData.bodyPreview : undefined;

  const result = {
    id: messageData.id,
    from,
    subject,
    contentPreview,
    content: messageData.body?.content,
    receivedDateTime,
    toRecipients,
    importance: messageData.importance,
    isRead: messageData.isRead,
    isDraft: messageData.isDraft,
  };

  return result;
};

/**
 * Converts a Microsoft Graph Recipient object to a simplified NameWithEmail format.
 *
 * @param recipient - Recipient object from Microsoft Graph API (optional)
 * @param fallback - Fallback NameWithEmail object to use if recipient data is missing (optional)
 * @returns NameWithEmail object with name and email properties, using fallback or defaults if needed
 */
export const toNameWithEmail = (recipient?: Recipient | null, fallback?: NameWithEmail): NameWithEmail => {
  const res = {
    name: recipient?.emailAddress?.name || fallback?.name || "Unknown User",
    email: recipient?.emailAddress?.address || fallback?.email || "",
  };

  return res;
};
