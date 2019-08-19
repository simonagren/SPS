// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { ActivityHandler, BotState, StatePropertyAccessor, ConversationState, UserState, CardFactory, TurnContext } from 'botbuilder';
import { Dialog, DialogState } from 'botbuilder-dialogs';
import { Logger } from '../logger';
import { MainDialog } from '../dialogs/mainDialog';
import WelcomeCard from '../resources/welcome.json';
import * as teams from 'botbuilder-teams';

export class ProvisionBot extends ActivityHandler {
    private conversationState: BotState;
    private userState: BotState;
    private logger: Logger;
    private dialog: Dialog;
    private dialogState: StatePropertyAccessor<DialogState>;
    private conversationReferences: any;

    /**
     *
     * @param {BotState} conversationState
     * @param {BotState} userState
     * @param {Dialog} dialog
     * @param {Logger} logger object for logging events, defaults to console if none is provided
     * @param {Test[]} conversationReferences
     */

    constructor(conversationState: BotState, userState: BotState, dialog: Dialog, logger: Logger, conversationReferences: any) {
        super();
        if (!conversationState) {
            throw new Error('[ProvisionBot]: Missing parameter. conversationState is required');
        }
        if (!userState) {
            throw new Error('[ProvisionBot]: Missing parameter. userState is required');
        }
        if (!dialog) {
            throw new Error('[ProvisionBot]: Missing parameter. dialog is required');
        }
        if (!logger) {
            logger = console as Logger;
            logger.log('[ProvisionBot]: logger not passed in, defaulting to console');
        }
        if (!conversationReferences) {
            throw new Error('[ProvisionBot]: Missing parameter. conversationReferences is required');
        }

        this.conversationState = conversationState as ConversationState;
        this.userState = userState as UserState;
        this.dialog = dialog;
        this.logger = logger;
        this.dialogState = this.conversationState.createProperty<DialogState>('DialogState');
        this.conversationReferences = conversationReferences;
        
        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                    await context.sendActivity({ attachments: [welcomeCard] });
                    // await (this.dialog as MainDialog).run(context, this.dialogState);

                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            
            // If Adaptive Card set .text to the value of .value
            if (context.activity.type === 'message') {

                const activity = context.activity;
                if (activity.text === undefined && activity.replyToId && activity.value && activity.value.isFromAdaptiveCard && activity.value.messageText) {
                    activity.text = activity.value.messageText;
                    context.sendActivity(activity);
                }
                
            }

            // Add a reference to this conversation that we could message later
            this.addConversationReference(context.activity);

            // Run the Main Dialog
            await (this.dialog as MainDialog).run(context, this.dialogState);

            // Save any state changes. The load happened during the execution of the Dialog.
            await this.conversationState.saveChanges(context, false);
            await this.userState.saveChanges(context, false);

            // By calling next() you ensure that the next BotHandler is run.
            await next();

        });

        this.onDialog(async (context, next) => {
            
            // Save any state changes. The load happened during the execution of the Dialog.
            await this.conversationState.saveChanges(context, false);
            await this.userState.saveChanges(context, false);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onTokenResponseEvent(async (context, next) => {
            console.log('Running dialog with Token Response Event Activity.');

            // Run the Dialog with the new Token Response Event Activity.
            await (this.dialog as MainDialog).run(context, this.dialogState);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    private addConversationReference(activity): void {
        const conversationReference = TurnContext.getConversationReference(activity);
        this.conversationReferences[conversationReference.conversation.id] = conversationReference;   
    }
}
