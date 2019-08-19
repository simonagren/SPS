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

import { CardFactory } from 'botbuilder';

import SiteTypesCard from '../resources/siteTypes.json';
import { SiteDesignResolverDialog } from './siteDesignResolverDialog';

const TEXT_PROMPT = 'textPrompt';
const CONFIRM_PROMPT = 'confirmPrompt';
const ALIAS_RESOLVER_DIALOG = 'aliasResolverDialog';
const OWNER_RESOLVER_DIALOG = 'ownerResolverDialog';
const SITE_DESIGN_RESOLVER_DIALOG = 'siteDesignResolverDialog';

const WATERFALL_DIALOG = 'waterfallDialog';

import { Logger } from "../logger";
import { sp, SiteDesignInfo } from '@pnp/sp';


export class SiteDialog extends CancelAndHelpDialog {
    private logger: Logger;

    constructor(id: string, logger: Logger) {
        super(id || 'bookingDialog');

        if (!logger) {
            logger = console as Logger;
            logger.log('[MainDialog]: logger not passed in, defaulting to console');
        }

        this.logger = logger;

        this
            .addDialog(new TextPrompt(TEXT_PROMPT))
            .addDialog(new AliasResolverDialog(ALIAS_RESOLVER_DIALOG))
            .addDialog(new OwnerResolverDialog(OWNER_RESOLVER_DIALOG))
            .addDialog(new SiteDesignResolverDialog(SITE_DESIGN_RESOLVER_DIALOG))
            .addDialog(new ConfirmPrompt(CONFIRM_PROMPT))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.siteTypeStep.bind(this),
                this.siteDesignStep.bind(this),
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
            // const siteTypesCard = CardFactory.adaptiveCard(SiteTypesCard.cards.map(card => CardFactory.adaptiveCard(card)));
            const siteTypeCards = SiteTypesCard.cards.map(card => CardFactory.adaptiveCard(card));
            await stepContext.context.sendActivity({ attachmentLayout: "carousel", attachments: siteTypeCards });
            return await stepContext.prompt('textPrompt', '');

        } else {
            return await stepContext.next(siteDetails.siteType);
        }
    }

    /**
     * If a site design has not been provided, prompt for one.
     */
    private async siteDesignStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        siteDetails.siteType = stepContext.result;

        if (!siteDetails.siteDesign) {

            const siteDesigns: SiteDesignInfo[] = await sp.siteDesigns.getSiteDesigns();
            const filtSiteDesigns: SiteDesignInfo[] = siteDesigns.filter(sd => (sd.WebTemplate ===
                (siteDetails.siteType === "ModernTeamSite" ? "64" : "68")));

            const siteDesignsCard = {
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock",
                        "horizontalAlignment": "Center",
                        "size": "Medium",
                        "weight": "Bolder",
                        "text": "Site design"
                    },
                    {
                        "type": "TextBlock",
                        "text": "What site design do you want?"
                    },
                    {
                        "type": "Input.ChoiceSet",
                        "id": "messageText",
                        "value": "1",
                        "choices": [
                            {
                                "title": "test",
                                "value":"hej"
                            }                        ]
                    }
                ],
                "actions": [
                    {
                        "type": "Action.Submit",
                          "title": "OK",
                          "data": {
                              "isFromAdaptiveCard": true
                          }

                    }
                ],
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.0"
            }
            await stepContext.context.sendActivity({ attachments: [CardFactory.adaptiveCard(siteDesignsCard)] });
            return await stepContext.prompt(TEXT_PROMPT, { prompt: '' });
        } else {
            return await stepContext.next(siteDetails.siteDesign);
        }
    }

    /**
     * If a title has not been provided, prompt for one.
     */
    private async titleStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // siteDetails.siteType = stepContext.result;

        // Capture the results of the previous step
        siteDetails.siteDesign = stepContext.result;
        if (!siteDetails.title) {
            return await stepContext.prompt(TEXT_PROMPT, { prompt: `Provide a title for your ${siteDetails.siteType}` });
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
            return await stepContext.prompt(TEXT_PROMPT, { prompt: `Provide a description for your ${siteDetails.siteType}` });
        } else {
            return await stepContext.next(siteDetails.title);
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
            return await stepContext.beginDialog(OWNER_RESOLVER_DIALOG, { owner: siteDetails.owner });
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

        if (!siteDetails.alias) {
            return await stepContext.beginDialog(ALIAS_RESOLVER_DIALOG, { alias: siteDetails.alias });
        } else {
            return await stepContext.next(siteDetails.alias);
        }

    }

    /**
     * Confirm the information the user has provided.
     */
    private async confirmStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // Capture the results of the previous step
        siteDetails.alias = stepContext.result;
        const msg = `A summary of your request:\n 
        Title: ${ siteDetails.title} \n\n
        Site design: ${ siteDetails.siteDesign } \n\n
        Owner: ${ siteDetails.owner} \n\n
        Description: ${ siteDetails.description} \n\n
        Site type: ${ siteDetails.siteType} \n\n
        Alias: ${ siteDetails.alias}.`;

        // Offer a YES/NO prompt.
        return await stepContext.prompt(CONFIRM_PROMPT, { prompt: msg });
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
