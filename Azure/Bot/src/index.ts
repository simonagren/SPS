// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';
import io from 'socket.io-client';
import * as teams from 'botbuilder-teams';
import { graph } from "@pnp/graph";
import { AdalFetchClient, SPFetchClient } from "@pnp/nodejs";

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { BotFrameworkAdapter, ConversationState, UserState, MemoryStorage, CardFactory, ConversationReference } from 'botbuilder';

// This bot's main dialog.
import { ProvisionBot } from './bots/provisionBot';
import { MainDialog } from './dialogs/mainDialog';
import { sp } from '@pnp/sp';

const ENV_FILE = path.join(__dirname, '..', '.env');
const loadFromEnv = config({ path: ENV_FILE });

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new teams.TeamsAdapter({
// const adapter = new BotFrameworkAdapter({
    appId: process.env.MicrosoftAppID,
    appPassword: process.env.MicrosoftAppPassword,
});

// Catch-all for errors.
adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    console.error(`\n [onTurnError]: ${ error }`);
    // Send a message to the user
    await context.sendActivity(`Oops. Something went wrong!`);
};

sp.setup({
    sp: {
        fetchClientFactory: () => {
            return new SPFetchClient(process.env.SharepointTenant, process.env.SharePointId, process.env.SharePointSecret);
        }
    }
});

// adapter.use(new teams.TeamsMiddleware());
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

// Create the main dialog.
const dialog = new MainDialog(logger);
const bot = new ProvisionBot(conversationState, userState, dialog, logger, conversationReferences);

// Create HTTP server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${ server.name } listening to ${ server.url }`);
    console.log(`\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator`);
    console.log(`\nSee https://aka.ms/connect-to-bot for more information`);
});

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
    adapter.processActivity(req, res, async (turnContext) => {
        // Route to bot activity handler.
        await bot.run(turnContext);
    });
});

server.post('/api/notify', async (req, res) => {
    for (let conversationReference of Object.values(conversationReferences)) {
        await adapter.continueConversation(conversationReference, async turnContext => {
            await turnContext.sendActivity('proactive hello');
            await turnContext.sendActivity({ attachments: [CardFactory.adaptiveCard(
                {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "dO STUFF"
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Http",
                            "title": "Approve",
                            "method": "POST",
                            "url": "https://actionablemessagessimon.azurewebsites.net/api/PollExample"
                        }
                    ],
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.0"
                }
            )
        ]});
        });
    }
    res.end('Ok');
});

const socket = io('http://localhost:3000');
socket.on('connect', () => {
    socket.on('updateFromServer', async(data: string) => {
        await sendMessageToConversation(data);
    });

    socket.on('doneFromServer', async(data: string) => {
        await sendMessageToConversation(data);
    });
})

const sendMessageToConversation = async (data: any) => {

    // Teams
    const ref: Partial<ConversationReference> = data.ConversationReference;
    const tenantId = data.tenantId;
    await adapter.createTeamsConversation(ref, tenantId, async turnContext => {
        await turnContext.sendActivity('');
    })
    
    // Generell
    for (const conversationReference of Object.values(conversationReferences)) {
        await adapter.continueConversation(conversationReference, async turnContext => {
            // await turnContext.sendActivity(message ? message : "default");
            const welcomeCard = CardFactory.adaptiveCard({
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "Image",
                                        "horizontalAlignment": "Center",
                                        "spacing": "None",
                                        "url": "https://vignette.wikia.nocookie.net/logopedia/images/1/11/Sharepoint365NewIcon.png/revision/latest?cb=20181202032451=1557084756765316",
                                        "size": "Small"
                                    }
                                ],
                                "width": "auto"
                            },
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "horizontalAlignment": "Right",
                                        "text": "Provisioning",
                                        "isSubtle": true
                                    },
                                    {
                                        "type": "TextBlock",
                                        "horizontalAlignment": "Right",
                                        "spacing": "None",
                                        "size": "Large",
                                        "color": "Good",
                                        "text": data.message
                                    }
                                ],
                                "width": "stretch"
                            }
                        ]
                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.0",
                "speak": "<s>Flight KL0605 to San Fransisco has been delayed.</s><s>It will not leave until 10:10 AM.</s>"
            });

            await turnContext.sendActivity({ attachments: [welcomeCard]})
        });
    }
}

