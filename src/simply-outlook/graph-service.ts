import { TokenCredential } from "@azure/identity";
import { Client, PageCollection } from "@microsoft/microsoft-graph-client";
import { Event, Message } from "@microsoft/microsoft-graph-types";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
import { JSDOM } from "jsdom";
import DOMPurify, { WindowLike } from "dompurify";
import { NodeHtmlMarkdown } from "node-html-markdown";
import { Marked } from "marked";
import { ILogger } from "../common/logger.types.js";
import { CalendarEventData, DateTimeRange, MailFolderData, MailMessageData } from "./graph-service.types.js";

const CALENDAR_EVENT_PROPS = [
  "id",
  "createdDateTime",
  "type",
  "subject",
  "start",
  "end",
  "body",
  "organizer",
  "categories",
  "iCalUId",
  "hasAttachments",
  "showAs",
  "isOnlineMeeting",
  "isOrganizer",
  "attendees",
  "onlineMeeting",
];

const MAIL_FOLDER_PROPS = ["id", "displayName", "wellKnownName"];

const MAIL_MESSAGE_PROPS = [
  "id",
  "receivedDateTime",
  "createdDateTime",
  "sentDateTime",
  "subject",
  "importance",
  "sender",
  "from",
  "toRecipients",
  "replyTo",
  "parentFolderId",
  "isRead",
  "isDraft",
  "categories",
];

const MAIL_PREVIEW_MESSAGE_PROPS = MAIL_MESSAGE_PROPS.concat(["bodyPreview"]);
const MAIL_BODY_MESSAGE_PROPS = MAIL_MESSAGE_PROPS.concat(["body"]);

const DEFAULT_MAIL_FOLDERS_LIMIT = 100;

const DELETED_FOLDER_NAME = "deleteditems";
const JUNK_FOLDER_NAME = "junkemail";

export class GraphService {
  private graphClient: Client;
  private domPurify: typeof DOMPurify;
  private nhm = new NodeHtmlMarkdown();
  private marked = new Marked({ gfm: true });
  private mailFolders: MailFolderData[];
  private filterFolderIds: string[] | undefined;

  constructor(private readonly logger: ILogger, private readonly tokenCredential: TokenCredential, private readonly scopes: string[]) {
    const authProvider = new TokenCredentialAuthenticationProvider(tokenCredential, {
      scopes,
    });
    this.graphClient = Client.initWithMiddleware({
      authProvider,
      defaultVersion: "beta",
      fetchOptions: { headers: { "User-Agent": "simply-outlook-mcp" } },
    });

    const window = new JSDOM("").window as unknown as WindowLike;
    this.domPurify = DOMPurify(window);
    this.mailFolders = [];
  }

  public async isAuthenticated(): Promise<boolean> {
    const token = await this.tokenCredential.getToken(this.scopes);
    return !!token;
  }

  public async getCalendarEvents(startDateTimeRange?: DateTimeRange, limit: number = 10, skip?: number): Promise<CalendarEventData[]> {
    const filters: string[] = [];
    const { startDateTime, endDateTime } = startDateTimeRange || {};
    let apiPath: string;
    if (startDateTime && endDateTime) {
      // all events
      apiPath = `/me/calendar/calendarView?startDateTime=${startDateTime}&endDateTime=${endDateTime}`;
    } else {
      if (startDateTime) {
        filters.push(`start/dateTime ge '${startDateTime}'`);
      }
      if (endDateTime) {
        filters.push(`start/dateTime lt '${endDateTime}'`);
      }
      // no event instances
      apiPath = `/me/calendar/events`;
    }

    const filterStr = filters.join(" and ");
    const query = this.graphClient
      .api(apiPath)
      .select(CALENDAR_EVENT_PROPS)
      .top(limit)
      .skip(skip || 0);
    filterStr && query.filter(filterStr);
    const collection: PageCollection = await query.get();
    if (!collection.value) {
      throw new Error("Failed to get events.");
    }

    const events = collection.value
      .filter((event) => this.isCalendarEventData(event))
      .map((event) => {
        if (event.body && event.body.content && event.body.contentType === "html") {
          event.body = {
            contentType: event.body.contentType,
            content: this.parseHtmlToMarkdown(event.body.content),
          };
        }
        return event;
      });
    return events;
  }

  public async createCalendarEvent(
    subject: string,
    content: string,
    utcStartDate: string,
    utcEndDate: string,
    userEmails?: string[],
    location?: string,
    isMeeting?: boolean
  ): Promise<CalendarEventData> {
    const attendees = userEmails ? userEmails.map((email) => ({ emailAddress: { address: email }, type: "required" })) : undefined;
    const eventRequest = {
      subject,
      body: {
        contentType: "html",
        content: this.parseMarkdownToHtml(content),
      },
      location: location ? { displayName: location } : undefined,
      isOnlineMeeting: !!isMeeting,
      start: {
        dateTime: utcStartDate,
        timeZone: "UTC",
      },
      end: {
        dateTime: utcEndDate,
        timeZone: "UTC",
      },
      attendees,
    };

    const event: Event = await this.graphClient.api(`/me/events`).post(eventRequest);
    if (!this.isCalendarEventData(event)) {
      throw new Error("Create event failed.");
    }

    if (event.body && event.body.content && event.body.contentType === "html") {
      event.body = {
        contentType: event.body.contentType,
        content: this.parseHtmlToMarkdown(event.body.content),
      };
    }

    return event;
  }

  public async updateCalendarEvent(
    id: string,
    content?: string,
    subject?: string,
    utcStartDate?: string,
    utcEndDate?: string,
    location?: string
  ): Promise<CalendarEventData> {
    if (
      content === undefined &&
      subject === undefined &&
      utcStartDate === undefined &&
      utcEndDate === undefined &&
      location === undefined
    ) {
      throw new Error("At least one property must be provided to update the calendar event.");
    }

    const updateRequest: Partial<Event> = {};

    if (subject !== undefined) {
      updateRequest.subject = subject;
    }

    if (content !== undefined) {
      updateRequest.body = {
        contentType: "html",
        content: this.parseMarkdownToHtml(content),
      };
    }

    if (utcStartDate !== undefined) {
      updateRequest.start = {
        dateTime: utcStartDate,
        timeZone: "UTC",
      };
    }

    if (utcEndDate !== undefined) {
      updateRequest.end = {
        dateTime: utcEndDate,
        timeZone: "UTC",
      };
    }

    if (location !== undefined) {
      updateRequest.location = location ? { displayName: location } : null;
    }

    const event: Event = await this.graphClient.api(`/me/events/${id}`).patch(updateRequest);
    if (!this.isCalendarEventData(event)) {
      throw new Error("Update event failed.");
    }

    if (event.body && event.body.content && event.body.contentType === "html") {
      event.body = {
        contentType: event.body.contentType,
        content: this.parseHtmlToMarkdown(event.body.content),
      };
    }

    return event;
  }

  public async getOutlookMessages(
    receivedDateTimeRange?: DateTimeRange,
    searchQuery?: string,
    limit: number = 10,
    skip?: number
  ): Promise<MailMessageData[]> {
    const filters: string[] = [];
    const { startDateTime, endDateTime } = receivedDateTimeRange || {};
    if (startDateTime) {
      filters.push(`receivedDateTime ge ${startDateTime}`);
    }

    if (endDateTime) {
      filters.push(`receivedDateTime lt ${endDateTime}`);
    }

    const filterFolders = await this.getFilterFolderIds();
    const folderIdSet = new Set<string>(filterFolders ? filterFolders : []);
    folderIdSet.forEach((folderId) => {
      filters.push(`parentFolderId ne '${folderId}'`);
    });

    const filterStr = filters.join(" and ");
    let query = this.graphClient.api("/me/messages").select(MAIL_PREVIEW_MESSAGE_PROPS).top(limit);
    // Graph search endpoint does not support MSA so use $search with limited functionalities
    const encodedQuery = searchQuery && encodeURIComponent(searchQuery);
    query = encodedQuery
      ? query.search(`"subject:${encodedQuery} OR body:${encodedQuery} OR from:${encodedQuery}"`)
      : query
          .filter(filterStr)
          .skip(skip || 0)
          .orderby("receivedDateTime desc");

    const collection: PageCollection = await query.get();
    if (!collection.value) {
      throw new Error("Failed to get messages.");
    }

    const messages = collection.value
      .filter((message) => this.isMailMessageData(message))
      .filter((message) => {
        if (searchQuery) {
          if (message.parentFolderId && folderIdSet.has(message.parentFolderId)) {
            return false;
          }

          if (startDateTime || endDateTime) {
            const receivedDate = new Date(message.receivedDateTime!);

            if (startDateTime && receivedDate < new Date(startDateTime)) {
              return false;
            }

            if (endDateTime && receivedDate >= new Date(endDateTime)) {
              return false;
            }
          }
        }
        return true;
      })
      .map((message) => {
        if (message.body && message.body.content && message.body.contentType === "html") {
          message.body = {
            contentType: message.body.contentType,
            content: this.parseHtmlToMarkdown(message.body.content),
          };
        }
        return message;
      });

    return messages;
  }

  public async getOutlookMessageById(id: string): Promise<MailMessageData> {
    const mailData = await this.graphClient.api(`/me/messages/${id}`).select(MAIL_BODY_MESSAGE_PROPS).get();
    if (!this.isMailMessageData(mailData)) {
      throw new Error("Get Outlook message failed.");
    }

    if (mailData.body && mailData.body.content && mailData.body.contentType === "html") {
      mailData.body = {
        contentType: mailData.body.contentType,
        content: this.parseHtmlToMarkdown(mailData.body.content),
      };
    }

    return mailData;
  }

  public async sendOutlookMessage(subject: string, content: string, recipientEmails: string[]): Promise<void> {
    const toRecipients = recipientEmails.map((email) => ({ emailAddress: { address: email } }));
    const msgRequest: Message = {
      subject,
      body: {
        contentType: "html",
        content: this.parseMarkdownToHtml(content),
      },
      toRecipients,
    };
    await this.graphClient.api("/me/sendMail").post({ message: msgRequest });
  }

  public async replyOutlookMessage(replyMessageId: string, content: string): Promise<void> {
    const originalMessage = await this.getOutlookMessageById(replyMessageId);

    const originalSender = originalMessage.from?.emailAddress?.name || originalMessage.from?.emailAddress?.address || "Unknown Sender";
    const originalDate = originalMessage.sentDateTime || originalMessage.receivedDateTime;
    const originalSubject = originalMessage.subject || "(No Subject)";
    const originalContent = originalMessage.body?.content || "";

    const replyContent = `${content}\n\n\n\n---\n\n**From:** ${originalSender}  \n**Date:** ${
      originalDate ? new Date(originalDate).toLocaleString() : "Unknown"
    }  \n**Subject:** ${originalSubject}  \n\n\n${originalContent}`;

    const msgRequest: Message = {
      body: {
        contentType: "html",
        content: this.parseMarkdownToHtml(replyContent),
      },
    };
    await this.graphClient.api(`/me/messages/${replyMessageId}/reply`).post({ message: msgRequest });
  }

  private async getMailFolders(limit?: number): Promise<MailFolderData[]> {
    const collection: PageCollection = await this.graphClient
      .api("/me/mailFolders")
      .select(MAIL_FOLDER_PROPS)
      .top(limit || DEFAULT_MAIL_FOLDERS_LIMIT)
      .get();
    if (!collection.value) {
      throw new Error("Failed to get mail folders.");
    }

    return collection.value.filter((value) => this.isMailFolderData(value));
  }

  private async getFilterFolderIds(): Promise<string[] | undefined> {
    if (this.filterFolderIds) {
      return this.filterFolderIds;
    }

    try {
      this.mailFolders = await this.getMailFolders(DEFAULT_MAIL_FOLDERS_LIMIT);
    } catch (error) {
      this.logger.error(`Failed to get mail folders: ${(error as Error).message}`);
      return undefined;
    }

    this.filterFolderIds = this.mailFolders
      .filter((folder) => folder.wellKnownName === DELETED_FOLDER_NAME || folder.wellKnownName === JUNK_FOLDER_NAME)
      .map((folder) => folder.id);

    return this.filterFolderIds;
  }

  private parseHtmlToMarkdown(htmlText: string): string {
    return this.nhm.translate(htmlText);
  }

  private parseMarkdownToHtml(markdownText: string): string {
    const html = this.marked.parse(markdownText, { async: false });
    return this.domPurify.sanitize(html);
  }

  private isCalendarEventData(event: Event): event is CalendarEventData {
    return !!event && !!event.id && !!event.type && !!event.start;
  }

  private isMailMessageData(message: Message): message is MailMessageData {
    return !!message && !!message.id && !!message.receivedDateTime;
  }

  private isMailFolderData(data: unknown): data is MailFolderData {
    const mailFolder = data as MailFolderData;
    return !!mailFolder && !!mailFolder.id && !!mailFolder.displayName;
  }
}
