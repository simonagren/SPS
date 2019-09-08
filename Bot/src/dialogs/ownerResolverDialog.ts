// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { TextPrompt, DialogTurnResult, PromptValidatorContext, WaterfallDialog, WaterfallStepContext } from 'botbuilder-dialogs';
import { CancelAndHelpDialog } from './cancelAndHelpDialog';
import { sp } from '@pnp/sp';
import { CardFactory } from 'botbuilder';
import { SiteDetails } from './siteDetails';

const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';
const GenericCard = require('../../resources/generic.json');

export class OwnerResolverDialog extends CancelAndHelpDialog {

    private static async textPromptValidator(promptContext: PromptValidatorContext<TextPrompt>): Promise<boolean> {
        if (promptContext.recognized.succeeded) {
            const owner: any = promptContext.recognized.value;
            if (!OwnerResolverDialog.validateEmail(owner)) {
                return false;
            }
            if (!await OwnerResolverDialog.userExists(owner)) {
                return false;
            }

            return true;
        } else {
            return false;
        }
    }

    private static validateEmail(email: string): boolean {
        const re = /^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))){2,6}$/i;
        return re.test(email);
    }

    private static async userExists(email: string): Promise<boolean> {
        try {
            // Get user via PnPJs
            const user = await sp.web.siteUsers.getByEmail(email).get();
            
            // If the user exists return true
            if (user) {
                return true;
            } else {
                return false;
            }
        } catch (error) {
            // If we don't get any user back, return false
            return false;
        }

    }

    constructor(id: string) {
        super(id || 'ownerResolverDialog');
        this.addDialog(new TextPrompt(TEXT_PROMPT, OwnerResolverDialog.textPromptValidator.bind(this)))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.initialStep.bind(this),
                this.finalStep.bind(this),
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = (stepContext.options as any).siteDetails;

        const promptMsg = `Provide an owner email for your ${siteDetails.siteType} site`;
        const repromptMsg = 'I\'m sorry, that email doesn\'t exists. Try again...';

        if (!siteDetails.owner) {

            const ownerCard = CardFactory.adaptiveCard(JSON.parse(
                JSON.stringify(GenericCard).replace('$Placeholder', promptMsg)));
        
            await stepContext.context.sendActivity({ attachments: [ownerCard] });
            return await stepContext.prompt(TEXT_PROMPT,
                {
                    prompt: '',
                    retryPrompt: repromptMsg,
                });

        } else {
            return await stepContext.next(siteDetails.owner);
        }


    }

    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const destination = stepContext.result;
        return await stepContext.endDialog(destination);
    }
}
