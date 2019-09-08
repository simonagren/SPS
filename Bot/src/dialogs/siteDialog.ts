// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import {
    ConfirmPrompt,
    DialogTurnResult,
    TextPrompt,
    WaterfallDialog,
    WaterfallStepContext,
} from 'botbuilder-dialogs';
import { SiteDetails } from './siteDetails';
import { CancelAndHelpDialog } from './cancelAndHelpDialog';
import { AliasResolverDialog } from './aliasResolverDialog';
import { OwnerResolverDialog } from './ownerResolverDialog';

import { CardFactory, TurnContext } from 'botbuilder';

// import * as teams from 'botbuilder-teams';

const SiteTypesCard = require('../../resources/siteTypes.json');
const GenericCard = require('../../resources/generic.json');
const SummaryCard = require('../../resources/summary.json');

const TEXT_PROMPT = 'textPrompt';
const CONFIRM_PROMPT = 'confirmPrompt';
const ALIAS_RESOLVER_DIALOG = 'aliasResolverDialog';
const OWNER_RESOLVER_DIALOG = 'ownerResolverDialog';
const WATERFALL_DIALOG = 'waterfallDialog';

import { Logger } from "../logger";

export class SiteDialog extends CancelAndHelpDialog {
    private logger: Logger;

    constructor(id: string, logger: Logger) {
        super(id || 'siteDialog');

        if (!logger) {
            logger = console as Logger;
            logger.log('[MainDialog]: logger not passed in, defaulting to console');
        }

        this.logger = logger;

        this
            .addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new AliasResolverDialog(ALIAS_RESOLVER_DIALOG))
            .addDialog(new OwnerResolverDialog(OWNER_RESOLVER_DIALOG))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.siteTypeStep.bind(this),
                this.titleStep.bind(this),
                this.descriptionStep.bind(this),
                this.ownerStep.bind(this),
                this.aliasStep.bind(this),
                this.confirmStep.bind(this),
                this.finalStep.bind(this),
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    /**
     * If a site type has not been provided, prompt for one.
     */
    private async siteTypeStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        if (!siteDetails.siteType) {

            const siteTypeCards = SiteTypesCard.cards.map(card => CardFactory.adaptiveCard(card));
            await stepContext.context.sendActivity({ attachmentLayout: "carousel", attachments: siteTypeCards });
            return await stepContext.prompt('textPrompt', '');

        } else {
            return await stepContext.next(siteDetails.siteType);
        }
    }

    /**
     * If a title has not been provided, prompt for one.
     */
    private async titleStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        siteDetails.siteType = stepContext.result;

        if (!siteDetails.title) {

            const text = `Provide a title for your ${siteDetails.siteType} site`;
            const titleCard = CardFactory.adaptiveCard(JSON.parse(
                JSON.stringify(GenericCard).replace('$Placeholder', text)));

            await stepContext.context.sendActivity({ attachments: [titleCard] });
            return await stepContext.prompt('textPrompt', '');
            // return await stepContext.prompt(TEXT_PROMPT, { prompt: `Provide a title for your ${siteDetails.siteType}` });
        } else {
            return await stepContext.next(siteDetails.title);
        }
    }

    /**
     * If a description has not been provided, prompt for one.
     */
    private async descriptionStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.title = stepContext.result;
        if (!siteDetails.description) {
            const text = `Provide a description for your ${siteDetails.siteType} site`;
            const descCard = CardFactory.adaptiveCard(JSON.parse(
                JSON.stringify(GenericCard).replace('$Placeholder', text)));

            await stepContext.context.sendActivity({ attachments: [descCard] });
            return await stepContext.prompt('textPrompt', '');
        } else {
            return await stepContext.next(siteDetails.description);
        }
    }

    /**
     * If an owner has not been provided, prompt for one.
     */
    private async ownerStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.description = stepContext.result;

        if (!siteDetails.owner) {
            return await stepContext.beginDialog(OWNER_RESOLVER_DIALOG, { siteDetails });
        } else {
            return await stepContext.next(siteDetails.owner);
        }

    }

    /**
     * If an owner has not been provided, prompt for one.
     */
    private async aliasStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.owner = stepContext.result;
        
        // Don't ask for alias if a communication site
        if (siteDetails.siteType === "Communication") {
            
            return await stepContext.next();
        
        // Otherwise ask for an alias
        } else {
            
            if (!siteDetails.alias) {
                return await stepContext.beginDialog(ALIAS_RESOLVER_DIALOG, { siteDetails });
            } else {
                return await stepContext.next(siteDetails.alias);
            }

        }

    }

    /**
     * Confirm the information the user has provided.
     */
    private async confirmStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // const teamsContext = teams.TeamsContext.from(stepContext.context);
        const ref = TurnContext.getConversationReference(stepContext.context.activity);

        // Capture the results of the previous step
        siteDetails.alias = stepContext.result;
        // siteDetails.tenantId = teamsContext ? teamsContext.tenant.id : '';
        siteDetails.conversationId = ref.conversation.id;

        const summaryCard = CardFactory.adaptiveCard(JSON.parse(
            JSON.stringify(SummaryCard)
                .replace('$Title', siteDetails.title)
                .replace('$Desc', siteDetails.description)
                .replace('$Owner', siteDetails.owner)
                .replace('$Type', siteDetails.siteType)
                .replace('$Alias', siteDetails.alias ? siteDetails.alias : "" )
                ));

        await stepContext.context.sendActivity({ attachments: [summaryCard] });

        // const msg = `A summary of your request:\n 
        // Title: ${ siteDetails.title} \n\n
        // Owner: ${ siteDetails.owner} \n\n
        // Description: ${ siteDetails.description} \n\n
        // Site type: ${ siteDetails.siteType} \n\n
        // Alias: ${ siteDetails.alias}.`;

        // Offer a YES/NO prompt.
        return await stepContext.prompt(CONFIRM_PROMPT, { prompt: '' });
    }

    /**
     * Complete the interaction and end the dialog.
     */
    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        if (stepContext.result === true) {
            const siteDetails = stepContext.options as SiteDetails;

            return await stepContext.endDialog(siteDetails);
        } else {
            return await stepContext.endDialog();
        }
    }

}
