import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { Client } from '@microsoft/microsoft-graph-client';
import { ConfidentialClientApplication } from '@azure/msal-node';
import * as dotenv from 'dotenv';
import { readFileSync, writeFileSync, existsSync } from 'fs';
import { join } from 'path';
import { homedir } from 'os';
import { z } from 'zod';

dotenv.config();

// Timezone utility functions
function getSystemTimezone() {
  return Intl.DateTimeFormat().resolvedOptions().timeZone;
}

function formatDateTimeForGraph(dateTime, timezone = null) {
  // If timezone is not provided, use system timezone
  const tz = timezone || getSystemTimezone();
  
  // Ensure the datetime string is properly formatted
  let dt = dateTime;
  if (!dt.includes('T')) {
    dt += 'T00:00:00';
  }
  if (!dt.includes('Z') && !dt.includes('+') && !dt.includes('-')) {
    dt += '.000Z';
  }
  
  return {
    dateTime: dt.replace('Z', ''),
    timeZone: tz
  };
}

// Configuration
const config = {
  auth: {
    clientId: process.env.OUTLOOK_CLIENT_ID,
    clientSecret: process.env.OUTLOOK_CLIENT_SECRET,
    authority: `https://login.microsoftonline.com/${process.env.OUTLOOK_TENANT_ID || 'common'}`
  }
};

// Token cache path - using current working directory
const tokenCachePath = '.mcp-outlook-token-cache.json';

// Initialize MSAL
const msalClient = new ConfidentialClientApplication(config);

// Token management
let tokenCache = {};

function loadTokenCache() {
  if (existsSync(tokenCachePath)) {
    try {
      tokenCache = JSON.parse(readFileSync(tokenCachePath, 'utf8'));
    } catch (e) {
      console.error('Failed to load token cache:', e);
    }
  }
}

function saveTokenCache() {
  try {
    writeFileSync(tokenCachePath, JSON.stringify(tokenCache, null, 2));
  } catch (e) {
    console.error('Failed to save token cache:', e);
  }
}

async function getAccessToken() {
  loadTokenCache();
  
  // Check if we have a valid token
  if (tokenCache.accessToken && tokenCache.expiresOn && new Date(tokenCache.expiresOn) > new Date()) {
    return tokenCache.accessToken;
  }

  // Try to use refresh token if available
  if (tokenCache.refreshToken) {
    try {
      const result = await msalClient.acquireTokenByRefreshToken({
        refreshToken: tokenCache.refreshToken,
        scopes: ['https://graph.microsoft.com/.default', 'offline_access']
      });
      
      tokenCache = {
        accessToken: result.accessToken,
        refreshToken: result.refreshToken || tokenCache.refreshToken,
        expiresOn: result.expiresOn
      };
      saveTokenCache();
      return result.accessToken;
    } catch (e) {
      console.error('Refresh token failed:', e);
    }
  }

  throw new Error('No valid authentication. Please run the authentication flow first.');
}

function getGraphClient(accessToken) {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

// Create MCP server
const server = new McpServer({
  name: 'mcp-outlook-server',
  version: '1.0.0'
});

// Authentication
server.tool(
  'outlook_auth',
  {
    authCode: z.string().optional().describe('Authorization code from OAuth flow (leave empty to get auth URL)')
  },
  async ({ authCode }) => {
    if (!authCode) {
      const authUrl = await msalClient.getAuthCodeUrl({
        scopes: ['https://graph.microsoft.com/.default', 'offline_access'],
        redirectUri: 'http://localhost:8080/callback'
      });
      return {
        content: [{
          type: 'text',
          text: `Please visit this URL to authenticate:\n${authUrl}\n\nThen call this tool again with the 'code' parameter from the redirect URL.`
        }]
      };
    } else {
      const result = await msalClient.acquireTokenByCode({
        code: authCode,
        scopes: ['https://graph.microsoft.com/.default', 'offline_access'],
        redirectUri: 'http://localhost:8080/callback'
      });
      
      tokenCache = {
        accessToken: result.accessToken,
        refreshToken: result.refreshToken,
        expiresOn: result.expiresOn
      };
      saveTokenCache();
      
      return {
        content: [{
          type: 'text',
          text: 'Authentication successful! You can now use other Outlook tools.'
        }]
      };
    }
  }
);

// ================================
// EMAIL MANAGEMENT TOOLS
// ================================

server.tool(
  'outlook_list_emails',
  {
    folder: z.string().default('inbox').describe('Folder to list emails from (default: inbox)'),
    limit: z.number().default(10).describe('Maximum number of emails to return'),
    search: z.string().optional().describe('Search query to filter emails')
  },
  async ({ folder, limit, search }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    let endpoint = `/me/mailFolders/${folder}/messages`;
    let query = {
      $top: limit,
      $select: 'id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments',
      $orderby: 'receivedDateTime DESC'
    };
    
    if (search) {
      query.$search = `"${search}"`;
    }
    
    const result = await client.api(endpoint).query(query).get();
    
    const emails = result.value.map(email => ({
      id: email.id,
      subject: email.subject,
      from: email.from?.emailAddress?.address || 'Unknown',
      received: email.receivedDateTime,
      preview: email.bodyPreview,
      isRead: email.isRead,
      importance: email.importance,
      hasAttachments: email.hasAttachments
    }));
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(emails, null, 2)
      }]
    };
  }
);

server.tool(
  'outlook_read_email',
  {
    emailId: z.string().describe('The ID of the email to read')
  },
  async ({ emailId }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const email = await client.api(`/me/messages/${emailId}`)
      .select('subject,from,to,cc,bcc,receivedDateTime,body,attachments,importance,categories')
      .expand('attachments')
      .get();
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          subject: email.subject,
          from: email.from?.emailAddress?.address,
          to: email.to?.map(r => r.emailAddress.address),
          cc: email.cc?.map(r => r.emailAddress.address),
          bcc: email.bcc?.map(r => r.emailAddress.address),
          received: email.receivedDateTime,
          body: email.body.content,
          importance: email.importance,
          categories: email.categories,
          attachments: email.attachments?.map(att => ({
            name: att.name,
            size: att.size,
            contentType: att.contentType
          }))
        }, null, 2)
      }]
    };
  }
);

server.tool(
  'outlook_send_email',
  {
    to: z.array(z.string()).describe('Array of recipient email addresses'),
    subject: z.string().describe('Email subject'),
    body: z.string().describe('Email body (HTML supported)'),
    cc: z.array(z.string()).optional().describe('Array of CC email addresses'),
    bcc: z.array(z.string()).optional().describe('Array of BCC email addresses'),
    importance: z.enum(['low', 'normal', 'high']).default('normal').describe('Email importance level')
  },
  async ({ to, subject, body, cc, bcc, importance }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const message = {
      subject,
      body: {
        contentType: 'HTML',
        content: body
      },
      importance,
      toRecipients: to.map(email => ({
        emailAddress: { address: email }
      }))
    };
    
    if (cc) {
      message.ccRecipients = cc.map(email => ({
        emailAddress: { address: email }
      }));
    }
    
    if (bcc) {
      message.bccRecipients = bcc.map(email => ({
        emailAddress: { address: email }
      }));
    }
    
    await client.api('/me/sendMail').post({ message });
    
    return {
      content: [{
        type: 'text',
        text: 'Email sent successfully!'
      }]
    };
  }
);

server.tool(
  'outlook_reply_email',
  {
    emailId: z.string().describe('The ID of the email to reply to'),
    body: z.string().describe('Reply body (HTML supported)'),
    replyAll: z.boolean().default(false).describe('Whether to reply to all recipients')
  },
  async ({ emailId, body, replyAll }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const replyData = {
      message: {
        body: {
          contentType: 'HTML',
          content: body
        }
      }
    };
    
    const endpoint = replyAll ? `/me/messages/${emailId}/replyAll` : `/me/messages/${emailId}/reply`;
    await client.api(endpoint).post(replyData);
    
    return {
      content: [{
        type: 'text',
        text: `Email ${replyAll ? 'reply all' : 'reply'} sent successfully!`
      }]
    };
  }
);

server.tool(
  'outlook_forward_email',
  {
    emailId: z.string().describe('The ID of the email to forward'),
    to: z.array(z.string()).describe('Array of recipient email addresses'),
    body: z.string().optional().describe('Additional message body (HTML supported)')
  },
  async ({ emailId, to, body }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const forwardData = {
      message: {
        toRecipients: to.map(email => ({
          emailAddress: { address: email }
        }))
      }
    };
    
    if (body) {
      forwardData.message.body = {
        contentType: 'HTML',
        content: body
      };
    }
    
    await client.api(`/me/messages/${emailId}/forward`).post(forwardData);
    
    return {
      content: [{
        type: 'text',
        text: 'Email forwarded successfully!'
      }]
    };
  }
);

server.tool(
  'outlook_delete_email',
  {
    emailId: z.string().describe('The ID of the email to delete')
  },
  async ({ emailId }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    await client.api(`/me/messages/${emailId}`).delete();
    
    return {
      content: [{
        type: 'text',
        text: 'Email deleted successfully!'
      }]
    };
  }
);

server.tool(
  'outlook_mark_email_read',
  {
    emailId: z.string().describe('The ID of the email to mark as read/unread'),
    isRead: z.boolean().describe('Whether to mark as read (true) or unread (false)')
  },
  async ({ emailId, isRead }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    await client.api(`/me/messages/${emailId}`).patch({ isRead });
    
    return {
      content: [{
        type: 'text',
        text: `Email marked as ${isRead ? 'read' : 'unread'} successfully!`
      }]
    };
  }
);

server.tool(
  'outlook_move_email',
  {
    emailId: z.string().describe('The ID of the email to move'),
    destinationFolder: z.string().describe('Name of the destination folder (e.g., "junk", "archive", "drafts")')
  },
  async ({ emailId, destinationFolder }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    // Get folder ID by name
    const folders = await client.api('/me/mailFolders').get();
    const targetFolder = folders.value.find(f => 
      f.displayName.toLowerCase() === destinationFolder.toLowerCase()
    );
    
    if (!targetFolder) {
      throw new Error(`Folder "${destinationFolder}" not found`);
    }
    
    await client.api(`/me/messages/${emailId}/move`).post({
      destinationId: targetFolder.id
    });
    
    return {
      content: [{
        type: 'text',
        text: `Email moved to ${destinationFolder} successfully!`
      }]
    };
  }
);

server.tool(
  'outlook_list_folders',
  {},
  async () => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const folders = await client.api('/me/mailFolders')
      .select('id,displayName,totalItemCount,unreadItemCount')
      .get();
    
    const folderList = folders.value.map(folder => ({
      id: folder.id,
      name: folder.displayName,
      totalItems: folder.totalItemCount,
      unreadItems: folder.unreadItemCount
    }));
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(folderList, null, 2)
      }]
    };
  }
);

// ================================
// CALENDAR MANAGEMENT TOOLS
// ================================

server.tool(
  'outlook_list_calendar_events',
  {
    startDateTime: z.string().optional().describe('Start date/time in ISO format (default: now)'),
    endDateTime: z.string().optional().describe('End date/time in ISO format (default: 7 days from now)'),
    limit: z.number().default(20).describe('Maximum number of events to return')
  },
  async ({ startDateTime, endDateTime, limit }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const now = new Date();
    const weekFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
    
    const start = startDateTime || now.toISOString();
    const end = endDateTime || weekFromNow.toISOString();
    
    const events = await client.api('/me/calendarView')
      .query({
        startDateTime: start,
        endDateTime: end,
        $top: limit,
        $select: 'id,subject,start,end,location,bodyPreview,organizer,attendees,importance,showAs,isAllDay',
        $orderby: 'start/dateTime'
      })
      .get();
    
    const formattedEvents = events.value.map(event => ({
      id: event.id,
      subject: event.subject,
      start: {
        dateTime: event.start.dateTime,
        timeZone: event.start.timeZone
      },
      end: {
        dateTime: event.end.dateTime,
        timeZone: event.end.timeZone
      },
      location: event.location?.displayName,
      preview: event.bodyPreview,
      organizer: event.organizer?.emailAddress?.address,
      attendees: event.attendees?.map(att => att.emailAddress.address),
      importance: event.importance,
      showAs: event.showAs,
      isAllDay: event.isAllDay
    }));
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(formattedEvents, null, 2)
      }]
    };
  }
);

server.tool(
  'outlook_create_calendar_event',
  {
    subject: z.string().describe('Event subject/title'),
    start: z.string().describe('Start date/time in ISO format (YYYY-MM-DDTHH:mm:ss or YYYY-MM-DD for all-day)'),
    end: z.string().describe('End date/time in ISO format (YYYY-MM-DDTHH:mm:ss or YYYY-MM-DD for all-day)'),
    body: z.string().optional().describe('Event description/body'),
    location: z.string().optional().describe('Event location'),
    attendees: z.array(z.string()).optional().describe('Array of attendee email addresses'),
    importance: z.enum(['low', 'normal', 'high']).default('normal').describe('Event importance'),
    showAs: z.enum(['free', 'tentative', 'busy', 'oof', 'workingElsewhere']).default('busy').describe('Show as status'),
    isAllDay: z.boolean().default(false).describe('Whether this is an all-day event'),
    timezone: z.string().optional().describe('Timezone (e.g., "Europe/London", "America/New_York"). Defaults to system timezone')
  },
  async ({ subject, start, end, body, location, attendees, importance, showAs, isAllDay, timezone }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const event = {
      subject,
      importance,
      showAs,
      isAllDay
    };
    
    // Handle timezone properly
    if (isAllDay) {
      // For all-day events, use date only without timezone
      event.start = {
        dateTime: start.split('T')[0],
        timeZone: 'tzone://Microsoft/Custom'
      };
      event.end = {
        dateTime: end.split('T')[0],
        timeZone: 'tzone://Microsoft/Custom'
      };
    } else {
      // For regular events, use proper timezone
      event.start = formatDateTimeForGraph(start, timezone);
      event.end = formatDateTimeForGraph(end, timezone);
    }
    
    if (body) {
      event.body = {
        contentType: 'HTML',
        content: body
      };
    }
    
    if (location) {
      event.location = {
        displayName: location
      };
    }
    
    if (attendees) {
      event.attendees = attendees.map(email => ({
        emailAddress: { address: email },
        type: 'required'
      }));
    }
    
    const result = await client.api('/me/events').post(event);
    
    return {
      content: [{
        type: 'text',
        text: `Calendar event created successfully! Event ID: ${result.id}\nTimezone: ${event.start.timeZone}`
      }]
    };
  }
);

server.tool(
  'outlook_update_calendar_event',
  {
    eventId: z.string().describe('The ID of the event to update'),
    subject: z.string().optional().describe('Updated event subject/title'),
    start: z.string().optional().describe('Updated start date/time in ISO format'),
    end: z.string().optional().describe('Updated end date/time in ISO format'),
    body: z.string().optional().describe('Updated event description/body'),
    location: z.string().optional().describe('Updated event location'),
    importance: z.enum(['low', 'normal', 'high']).optional().describe('Updated event importance'),
    timezone: z.string().optional().describe('Timezone for the event times')
  },
  async ({ eventId, subject, start, end, body, location, importance, timezone }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const updateData = {};
    
    if (subject) updateData.subject = subject;
    if (start) updateData.start = formatDateTimeForGraph(start, timezone);
    if (end) updateData.end = formatDateTimeForGraph(end, timezone);
    if (body) updateData.body = { contentType: 'HTML', content: body };
    if (location) updateData.location = { displayName: location };
    if (importance) updateData.importance = importance;
    
    await client.api(`/me/events/${eventId}`).patch(updateData);
    
    return {
      content: [{
        type: 'text',
        text: 'Calendar event updated successfully!'
      }]
    };
  }
);







server.tool(
  'outlook_delete_calendar_event',
  {
    eventId: z.string().describe('The ID of the event to delete')
  },
  async ({ eventId }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    await client.api(`/me/events/${eventId}`).delete();
    
    return {
      content: [{
        type: 'text',
        text: 'Calendar event deleted successfully!'
      }]
    };
  }
);

server.tool(
  'outlook_get_calendar_event',
  {
    eventId: z.string().describe('The ID of the event to retrieve')
  },
  async ({ eventId }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const event = await client.api(`/me/events/${eventId}`)
      .select('id,subject,start,end,location,body,organizer,attendees,importance,showAs,categories')
      .get();
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          id: event.id,
          subject: event.subject,
          start: event.start.dateTime,
          end: event.end.dateTime,
          location: event.location?.displayName,
          body: event.body?.content,
          organizer: event.organizer?.emailAddress?.address,
          attendees: event.attendees?.map(att => ({
            email: att.emailAddress.address,
            response: att.status.response
          })),
          importance: event.importance,
          showAs: event.showAs,
          categories: event.categories
        }, null, 2)
      }]
    };
  }
);

// ================================
// TASK MANAGEMENT TOOLS
// ================================

server.tool(
  'outlook_list_tasks',
  {
    completed: z.boolean().optional().describe('Filter by completion status'),
    limit: z.number().default(20).describe('Maximum number of tasks to return')
  },
  async ({ completed, limit }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const lists = await client.api('/me/todo/lists').get();
    
    if (lists.value.length === 0) {
      return {
        content: [{
          type: 'text',
          text: 'No task lists found.'
        }]
      };
    }
    
    const listId = lists.value[0].id;
    let tasksQuery = {
      $top: limit,
      $select: 'id,title,body,dueDateTime,importance,status,completedDateTime,createdDateTime'
    };
    
    if (completed !== undefined) {
      tasksQuery.$filter = completed ? 'status eq \'completed\'' : 'status ne \'completed\'';
    }
    
    const tasks = await client.api(`/me/todo/lists/${listId}/tasks`)
      .query(tasksQuery)
      .get();
    
    const formattedTasks = tasks.value.map(task => ({
      id: task.id,
      title: task.title,
      body: task.body?.content,
      dueDateTime: task.dueDateTime?.dateTime,
      importance: task.importance,
      status: task.status,
      completedDateTime: task.completedDateTime?.dateTime,
      createdDateTime: task.createdDateTime
    }));
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(formattedTasks, null, 2)
      }]
    };
  }
);

server.tool(
  'outlook_create_task',
  {
    title: z.string().describe('Task title'),
    body: z.string().optional().describe('Task description'),
    dueDateTime: z.string().optional().describe('Due date/time in ISO format'),
    importance: z.enum(['low', 'normal', 'high']).default('normal').describe('Task importance level')
  },
  async ({ title, body, dueDateTime, importance }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const lists = await client.api('/me/todo/lists').get();
    if (lists.value.length === 0) {
      throw new Error('No task lists found. Please create a task list in Outlook first.');
    }
    
    const listId = lists.value[0].id;
    
    const task = {
      title,
      importance
    };
    
    if (body) {
      task.body = {
        content: body,
        contentType: 'text'
      };
    }
    
    if (dueDateTime) {
      task.dueDateTime = {
        dateTime: dueDateTime,
        timeZone: 'UTC'
      };
    }
    
    const result = await client.api(`/me/todo/lists/${listId}/tasks`).post(task);
    
    return {
      content: [{
        type: 'text',
        text: `Task created successfully! Task ID: ${result.id}`
      }]
    };
  }
);

server.tool(
  'outlook_update_task',
  {
    taskId: z.string().describe('The ID of the task to update'),
    title: z.string().optional().describe('Updated task title'),
    body: z.string().optional().describe('Updated task description'),
    dueDateTime: z.string().optional().describe('Updated due date/time in ISO format'),
    importance: z.enum(['low', 'normal', 'high']).optional().describe('Updated task importance'),
    status: z.enum(['notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred']).optional().describe('Updated task status')
  },
  async ({ taskId, title, body, dueDateTime, importance, status }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const lists = await client.api('/me/todo/lists').get();
    const listId = lists.value[0].id;
    
    const updateData = {};
    
    if (title) updateData.title = title;
    if (body) updateData.body = { content: body, contentType: 'text' };
    if (dueDateTime) updateData.dueDateTime = { dateTime: dueDateTime, timeZone: 'UTC' };
    if (importance) updateData.importance = importance;
    if (status) updateData.status = status;
    
    await client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).patch(updateData);
    
    return {
      content: [{
        type: 'text',
        text: 'Task updated successfully!'
      }]
    };
  }
);

server.tool(
  'outlook_delete_task',
  {
    taskId: z.string().describe('The ID of the task to delete')
  },
  async ({ taskId }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const lists = await client.api('/me/todo/lists').get();
    const listId = lists.value[0].id;
    
    await client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).delete();
    
    return {
      content: [{
        type: 'text',
        text: 'Task deleted successfully!'
      }]
    };
  }
);

server.tool(
  'outlook_complete_task',
  {
    taskId: z.string().describe('The ID of the task to mark as completed')
  },
  async ({ taskId }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const lists = await client.api('/me/todo/lists').get();
    const listId = lists.value[0].id;
    
    await client.api(`/me/todo/lists/${listId}/tasks/${taskId}`).patch({
      status: 'completed'
    });
    
    return {
      content: [{
        type: 'text',
        text: 'Task marked as completed successfully!'
      }]
    };
  }
);

// ================================
// CONTACT MANAGEMENT TOOLS
// ================================

server.tool(
  'outlook_list_contacts',
  {
    limit: z.number().default(50).describe('Maximum number of contacts to return'),
    search: z.string().optional().describe('Search query to filter contacts')
  },
  async ({ limit, search }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    let query = {
      $top: limit,
      $select: 'id,displayName,emailAddresses,phoneNumbers,companyName,jobTitle'
    };
    
    if (search) {
      query.$filter = `startswith(displayName,'${search}') or startswith(givenName,'${search}') or startswith(surname,'${search}')`;
    }
    
    const contacts = await client.api('/me/contacts').query(query).get();
    
    const formattedContacts = contacts.value.map(contact => ({
      id: contact.id,
      name: contact.displayName,
      emails: contact.emailAddresses?.map(e => e.address),
      phones: contact.phoneNumbers?.map(p => ({ type: p.type, number: p.number })),
      company: contact.companyName,
      jobTitle: contact.jobTitle
    }));
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify(formattedContacts, null, 2)
      }]
    };
  }
);

server.tool(
  'outlook_create_contact',
  {
    displayName: z.string().describe('Contact display name'),
    emailAddress: z.string().optional().describe('Primary email address'),
    phoneNumber: z.string().optional().describe('Primary phone number'),
    companyName: z.string().optional().describe('Company name'),
    jobTitle: z.string().optional().describe('Job title')
  },
  async ({ displayName, emailAddress, phoneNumber, companyName, jobTitle }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const contact = {
      displayName
    };
    
    if (emailAddress) {
      contact.emailAddresses = [{
        address: emailAddress,
        name: displayName
      }];
    }
    
    if (phoneNumber) {
      contact.phoneNumbers = [{
        type: 'mobile',
        number: phoneNumber
      }];
    }
    
    if (companyName) contact.companyName = companyName;
    if (jobTitle) contact.jobTitle = jobTitle;
    
    const result = await client.api('/me/contacts').post(contact);
    
    return {
      content: [{
        type: 'text',
        text: `Contact created successfully! Contact ID: ${result.id}`
      }]
    };
  }
);

server.tool(
  'outlook_delete_contact',
  {
    contactId: z.string().describe('The ID of the contact to delete')
  },
  async ({ contactId }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    await client.api(`/me/contacts/${contactId}`).delete();
    
    return {
      content: [{
        type: 'text',
        text: 'Contact deleted successfully!'
      }]
    };
  }
);

// ================================
// UTILITY TOOLS
// ================================

server.tool(
  'outlook_get_user_profile',
  {},
  async () => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    const profile = await client.api('/me')
      .select('id,displayName,mail,userPrincipalName,officeLocation,jobTitle,department')
      .get();
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          id: profile.id,
          name: profile.displayName,
          email: profile.mail,
          userPrincipalName: profile.userPrincipalName,
          office: profile.officeLocation,
          jobTitle: profile.jobTitle,
          department: profile.department
        }, null, 2)
      }]
    };
  }
);

server.tool(
  'outlook_search_all',
  {
    query: z.string().describe('Search query to find across emails, events, and contacts'),
    limit: z.number().default(10).describe('Maximum number of results per category')
  },
  async ({ query, limit }) => {
    const accessToken = await getAccessToken();
    const client = getGraphClient(accessToken);
    
    // Search emails
    const emails = await client.api('/me/messages')
      .query({
        $search: `"${query}"`,
        $top: limit,
        $select: 'id,subject,from,receivedDateTime'
      })
      .get();
    
    // Search events
    const events = await client.api('/me/events')
      .query({
        $search: `"${query}"`,
        $top: limit,
        $select: 'id,subject,start,organizer'
      })
      .get();
    
    // Search contacts
    const contacts = await client.api('/me/contacts')
      .query({
        $filter: `contains(displayName,'${query}')`,
        $top: limit,
        $select: 'id,displayName,emailAddresses'
      })
      .get();
    
    return {
      content: [{
        type: 'text',
        text: JSON.stringify({
          emails: emails.value.map(e => ({
            id: e.id,
            subject: e.subject,
            from: e.from?.emailAddress?.address,
            received: e.receivedDateTime
          })),
          events: events.value.map(e => ({
            id: e.id,
            subject: e.subject,
            start: e.start?.dateTime,
            organizer: e.organizer?.emailAddress?.address
          })),
          contacts: contacts.value.map(c => ({
            id: c.id,
            name: c.displayName,
            email: c.emailAddresses?.[0]?.address
          }))
        }, null, 2)
      }]
    };
  }
);

// Start the server
const transport = new StdioServerTransport();
await server.connect(transport);
console.error('Enhanced Outlook MCP server running with full functionality');