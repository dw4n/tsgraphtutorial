import * as readline from 'readline-sync';
import { DeviceCodeInfo } from '@azure/identity';
import { Calendar, Message, User, Event as mcgEvent } from '@microsoft/microsoft-graph-types';
import settings, { AppSettings } from './appSettings';
import * as graphHelper from './graphHelper';

async function main() {
  console.log('TypeScript Graph Tutorial');

  let choice = 0;

  // Initialize Graph
  initializeGraph(settings);

  // Greet the user by name
  await greetUserAsync();

  const choices = [
    'Display access token',
    'List my inbox',
    'Send mail',
    'List users (requires app-only)',
    'Make a Graph call',
    'listCalendarsAsync',
    'listEventsAsync'
  ];

  while (choice != -1) {
    choice = readline.keyInSelect(choices, 'Select an option', { cancel: 'Exit' });

    switch (choice) {
      case -1:
        // Exit
        console.log('Goodbye...');
        break;
      case 0:
        // Display access token
        await displayAccessTokenAsync();
        break;
      case 1:
        // List emails from user's inbox
        await listInboxAsync();
        break;
      case 2:
        // Send an email message
        await sendMailAsync();
        break;
      case 3:
        // List users
        await listUsersAsync();
        break;
      case 4:
        // Run any Graph code
        await makeGraphCallAsync();
        break;
        case 5:
        // Run any Graph code
        await listCalendarsAsync();
        break;
        case 6:
        // Run any Graph code
        await listEventsAsync();
        break;
      default:
        console.log('Invalid choice! Please try again.');
    }
  }
}

main();


function initializeGraph(settings: AppSettings) {
  graphHelper.initializeGraphForUserAuth(settings, (info: DeviceCodeInfo) => {
    // Display the device code message to
    // the user. This tells them
    // where to go to sign in and provides the
    // code to use.
    console.log(info.message);
  });
}
  
async function greetUserAsync() {
  try {
    const user = await graphHelper.getUserAsync();
    console.log(`Hello, ${user?.displayName}!`);
    // For Work/school accounts, email is in mail property
    // Personal accounts, email is in userPrincipalName
    console.log(`Email: ${user?.mail ?? user?.userPrincipalName ?? ''}`);
  } catch (err) {
    console.log(`Error getting user: ${err}`);
  }
}
  
  async function displayAccessTokenAsync() {
    try {
      const userToken = await graphHelper.getUserTokenAsync();
      console.log(`User token: ${userToken}`);
    } catch (err) {
      console.log(`Error getting user access token: ${err}`);
    }
  }
  
  async function listInboxAsync() {
    try {
      const messagePage = await graphHelper.getInboxAsync();
      const messages: Message[] = messagePage.value;
  
      // Output each message's details
      for (const message of messages) {
        console.log(`Message: ${message.subject ?? 'NO SUBJECT'}`);
        console.log(`  From: ${message.from?.emailAddress?.name ?? 'UNKNOWN'}`);
        console.log(`  Status: ${message.isRead ? 'Read' : 'Unread'}`);
        console.log(`  Received: ${message.receivedDateTime}`);
      }
  
      // If @odata.nextLink is not undefined, there are more messages
      // available on the server
      const moreAvailable = messagePage['@odata.nextLink'] != undefined;
      console.log(`\nMore messages available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting user's inbox: ${err}`);
    }
  }
  
  async function sendMailAsync() {
    // TODO
  }
  
  async function listUsersAsync() {
    try {
      const userPage = await graphHelper.getUsersAsync();
      const users: User[] = userPage.value;
  
      // Output each user's details
      for (const user of users) {
        console.log(`User: ${user.displayName ?? 'NO NAME'}`);
        console.log(`  ID: ${user.id}`);
        console.log(`  Email: ${user.mail ?? 'NO EMAIL'}`);
      }
  
      // If @odata.nextLink is not undefined, there are more users
      // available on the server
      const moreAvailable = userPage['@odata.nextLink'] != undefined;
      console.log(`\nMore users available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting users: ${err}`);
    }
  }
  
  async function makeGraphCallAsync() {
    try {
      await graphHelper.makeGraphCallAsync();
    } catch (err) {
      console.log(`Error making Graph call: ${err}`);
    }
  }


  async function listCalendarsAsync() {
    try {
      const calendarPage = await graphHelper.getCalendarsAsync();
      const calendars: Calendar[] = calendarPage.value;
  
      // Output each user's details
      for (const cal of calendars) {
        console.log(`Name: ${cal.name ?? 'NO NAME'}`);
        console.log(`  ID: ${cal.id}`);
        console.log(`  Owner: ${cal.owner ?? 'NO Owner'}`);
      }
  
      // If @odata.nextLink is not undefined, there are more users
      // available on the server
      const moreAvailable = calendarPage['@odata.nextLink'] != undefined;
      console.log(`\nMore calendars available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting calendar: ${err}`);
    }
  }

  async function listEventsAsync() {
    try {
      const EventPage = await graphHelper.getCalendarEventsAsync();
      const events: mcgEvent[] = EventPage.value;
  
      // Output each user's details
      for (const event of events) {
        console.log(`User: ${event.subject ?? 'NO NAME'}`);
        console.log(`  ID: ${event.id}`);
        console.log(`  bodyPreview: ${event.bodyPreview ?? 'NO bodyPreview'}`);
        console.log(`  start: ${event.start ?? 'NO start'}`);
        console.log(`  end: ${event.end ?? 'NO end'}`);
        console.log(`  organizer: ${event.organizer?.emailAddress?.name ?? 'NO organizer'}`);
        console.log(`  location: ${event.location?.displayName ?? 'NO location'}`);
      }
  
      // If @odata.nextLink is not undefined, there are more users
      // available on the server
      const moreAvailable = EventPage['@odata.nextLink'] != undefined;
      console.log(`\nMore users available? ${moreAvailable}`);
    } catch (err) {
      console.log(`Error getting users: ${err}`);
    }
  }