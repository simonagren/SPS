// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';
import { sp } from '@pnp/sp';
import { SPFetchClient } from "@pnp/nodejs";
import * as io from 'socket.io-client';

const UpdateCard =  require('../resources/update.json');
const CompleteCard = require('../resources/complete.json');
const StartCard = require('../resources/start.json');

// Import required bot services. // See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { BotFrameworkAdapter, ConversationState, MemoryStorage, UserState, CardFactory } from 'botbuilder';

// The bot and its main dialog.
import { ProvisionBot } from './bots/provisionBot';
// import { DialogAndWelcomeBot } from './bots/dialogAndWelcomeBot';
import { MainDialog } from './dialogs/mainDialog';

// Note: Ensure you have a .env file and include LuisAppId, LuisAPIKey and LuisAPIHostName.
const ENV_FILE = path.join(__dirname, '..', '.env');
config({ path: ENV_FILE });

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppID,
    appPassword: process.env.MicrosoftAppPassword
});

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log
    // NOTE: In production environment, you should consider logging this to Azure
    //       application insights.
    console.error(`\n [onTurnError]: ${ error }`);
    // Send a message to the user
    await context.sendActivity(`Oops. Something went wrong!`);
    // Clear out state
    await conversationState.delete(context);
};

sp.setup({
    sp: {
        fetchClientFactory: () => {
            return new SPFetchClient(process.env.SharepointTenant+'/sites/sps/', process.env.SharePointId, process.env.SharePointSecret);
        }
    }
});

const socket = io("https://spsexpress.azurewebsites.net/");

socket.on('connect', () => {
    socket.emit('room', 'bot');
});

socket.on('startProvisioning', (data) => {
    _sendProvisioningIntro(data);
});

socket.on('provisioningUpdate', (data) => {
    _sendProvisioningUpdate(data);
});

socket.on('provisioningComplete', (data) => {
    _sendProvisioningComplete(data);
});


// Define a state store for your bot. See https://aka.ms/about-bot-state to learn more about using MemoryStorage.
// A bot requires a state store to persist the dialog and user state between messages.
let conversationState: ConversationState;
let userState: UserState;

// For local development, in-memory storage is used.
// CAUTION: The Memory Storage used here is for local bot debugging only. When the bot
// is restarted, anything stored in memory will be gone.
const memoryStorage = new MemoryStorage();
conversationState = new ConversationState(memoryStorage);
userState = new UserState(memoryStorage);

const conversationReferences = {};

// Pass in a logger to the bot. For this sample, the logger is the console, but alternatives such as Application Insights and Event Hub exist for storing the logs of the bot.
const logger = console;

const dialog = new MainDialog(logger);
const bot = new ProvisionBot(conversationState, userState, dialog, logger, conversationReferences, socket);


// Create HTTP server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${ server.name } listening to ${ server.url }`);
    console.log(`\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator`);
    console.log(`\nTo test your bot, see: https://aka.ms/debug-with-emulator`);
});

// Listen for incoming activities and route them to your bot main dialog.
server.post('/api/messages', (req, res) => {
    // Route received a request to adapter for processing
    adapter.processActivity(req, res, async (turnContext) => {
        // route to bot activity handler.
        await bot.run(turnContext);
    });
});

const _sendProvisioningIntro = async (data: any) => {
    // Get specific conversation reference
    const conversationReference = conversationReferences[data.conversationId]

    // Send start card
    const startCard = CardFactory.adaptiveCard(StartCard);

    // Send Adaptive Card to user
    await adapter.continueConversation(conversationReference, async turnContext => {
        await turnContext.sendActivity({ attachments: [startCard] });
    });

    // const tenantId: string = data.tenantId;
    
    // Teams
    // await adapter.createTeamsConversation(ref, tenantId, async turnContext => {
    //     await turnContext.sendActivity(data);
    // });

    // Bot Framework
    // await adapter.continueConversation(ref, async turnContext => {
    //     await turnContext.sendActivity(data.result);
    // });
}

const _sendProvisioningUpdate = async (data: any) => {
    
    // Get specoific conversation reference
    const conversationReference = conversationReferences[data.conversationId]

    // Do a replace with the provisioning message
    const updateCard = CardFactory.adaptiveCard(JSON.parse(
        JSON.stringify(UpdateCard).replace('$Placeholder', data.result)));

    // Send Adaptive Card to user
    await adapter.continueConversation(conversationReference, async turnContext => {
        await turnContext.sendActivity({ attachments: [updateCard] });
    });

    // const tenantId: string = data.tenantId;
    
    // Teams
    // await adapter.createTeamsConversation(ref, tenantId, async turnContext => {
    //     await turnContext.sendActivity(data);
    // });

    // Bot Framework
    // await adapter.continueConversation(ref, async turnContext => {
    //     await turnContext.sendActivity(data.result);
    // });
}

const _sendProvisioningComplete = async (data: any) => {
    
    // Get specoific conversation reference
    const conversationReference = conversationReferences[data.conversationId]

    // Do a replace with the provisioning message
    const completeCard = CardFactory.adaptiveCard(JSON.parse(
        JSON.stringify(CompleteCard).replace('$Placeholder', data.result)));

    // Send Adaptive Card to user
    await adapter.continueConversation(conversationReference, async turnContext => {
        await turnContext.sendActivity({ attachments: [completeCard] });
    });

    // const tenantId: string = data.tenantId;
    
    // Teams
    // await adapter.createTeamsConversation(ref, tenantId, async turnContext => {
    //     await turnContext.sendActivity(data);
    // });

    // Bot Framework
    // await adapter.continueConversation(ref, async turnContext => {
    //     await turnContext.sendActivity(data.result);
    // });
}
