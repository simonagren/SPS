// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { TextPrompt, DialogTurnResult, PromptValidatorContext, WaterfallDialog, WaterfallStepContext } from 'botbuilder-dialogs';
import { CancelAndHelpDialog } from './cancelAndHelpDialog';
import { SiteDetails } from './siteDetails';
import { graph } from '@pnp/graph';
import { AdalFetchClient, SPFetchClient } from '@pnp/nodejs';
import { sp } from '@pnp/sp';

const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

export class SiteDesignResolverDialog extends CancelAndHelpDialog {

    private static async textPromptValidator(promptContext: PromptValidatorContext<TextPrompt>): Promise<boolean> {
        if (promptContext.recognized.succeeded) {
            
            sp.setup({
                sp: {
                    fetchClientFactory: () => {
                        return new SPFetchClient("https://simonmvp.sharepoint.com/", process.env.MicrosoftAppID, process.env.MicrosoftAppPassword);
                    }
                }
            });
            

            const test = await sp.siteDesigns.getSiteDesigns();
            const users = await graph.sites.root.get();
            if (users != null) {
                return true;
            } else {
                return false;
            }
            // return true;
        } else {
            return false;
        }
    }

    constructor(id: string) {
        super(id || 'siteDesignResolverDialog');
        this.addDialog(new TextPrompt(TEXT_PROMPT, SiteDesignResolverDialog.textPromptValidator.bind(this)))
            .addDialog(new WaterfallDialog(WATERFALL_DIALOG, [
                this.initialStep.bind(this),
                this.finalStep.bind(this),
            ]));

        this.initialDialogId = WATERFALL_DIALOG;
    }

    private async initialStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const siteDetails = stepContext.options as SiteDetails;

        // const owner = (stepContext.options as any).owner;

        const promptMsg = 'Please enter a site design';
        const repromptMsg = 'I\'m sorry, that email doesn\'t exists.';

        if (!siteDetails.siteDesign) {

            return await stepContext.prompt(TEXT_PROMPT,
                {
                    prompt: promptMsg,
                    retryPrompt: repromptMsg,
                });

        } else {
            return await stepContext.next(siteDetails.siteDesign);
        }


    }

    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const destination = stepContext.result;
        return await stepContext.endDialog(destination);
    }

    // private async getSiteDesigns: Promise<any> {
    //     Request.get('')

    // }
}
