// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { TextPrompt, DialogTurnResult, PromptValidatorContext, WaterfallDialog, WaterfallStepContext } from 'botbuilder-dialogs';
import { CancelAndHelpDialog } from './cancelAndHelpDialog';
import { SPHttpClient } from '@pnp/sp';
import { CardFactory } from 'botbuilder';

const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

const GenericCard = require('../../resources/generic.json');

export class AliasResolverDialog extends CancelAndHelpDialog {

    private static async textPromptValidator(promptContext: PromptValidatorContext<TextPrompt>): Promise<boolean> {
        if (promptContext.recognized.succeeded) {
            const alias: any = promptContext.recognized.value;

            if (!await AliasResolverDialog.aliasExists(alias)) {
                return false;
            }

            return true;

        } else {
            return false;
        }
    }

    private static async aliasExists(alias: string): Promise<boolean> {
        try {
            const client = new SPHttpClient();
            const res = client.get(
                `https://simonmvp.sharepoint.com/_api/SP.Directory.DirectorySession/ValidateGroupName(displayName='${alias}',%20alias='${alias}')`)
            .then((response: Response) => {
                return response.json();
                
            }).then(result => {
                return result.IsValidName;
            });

            return res;

        } catch (error) {
            return false;
        }

    }

    constructor(id: string) {
        super(id || 'aliasResolverDialog');
        this.addDialog(new TextPrompt(TEXT_PROMPT, AliasResolverDialog.textPromptValidator.bind(this)))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.initialStep.bind(this),
                this.finalStep.bind(this),
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = (stepContext.options as any).siteDetails;

        const promptMsg = `Provide an alias for your ${siteDetails.siteType} site`;
        const repromptMsg = 'I\'m sorry, that alias already exists. Try again...';

        if (!siteDetails.alias) {

            const aliasCard = CardFactory.adaptiveCard(JSON.parse(
                JSON.stringify(GenericCard).replace('$Placeholder', promptMsg)));
        
            await stepContext.context.sendActivity({ attachments: [aliasCard] });
            
            return await stepContext.prompt(TEXT_PROMPT,
                {
                    prompt: '',
                    retryPrompt: repromptMsg,
                });

        } else {
            return await stepContext.next(siteDetails.alias);
        }

    }

    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const destination = stepContext.result;
        return await stepContext.endDialog(destination);
    }
}
