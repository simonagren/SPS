// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { TextPrompt, DialogTurnResult, PromptValidatorContext, WaterfallDialog, WaterfallStepContext } from 'botbuilder-dialogs';
import { CancelAndHelpDialog } from './cancelAndHelpDialog';

const TEXT_PROMPT = 'textPrompt';
const WATERFALL_DIALOG = 'waterfallDialog';

export class OwnerResolverDialog extends CancelAndHelpDialog {

    private static async textPromptValidator(promptContext: PromptValidatorContext<TextPrompt>): Promise<boolean> {
        if (promptContext.recognized.succeeded) {
            return true;
        } else {
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
        const owner = (stepContext.options as any).owner;

        const promptMsg = 'Please enter an Owner email for your site';
        const repromptMsg = 'I\'m sorry, that email doesn\'t exists.';

        if (!owner) {

            return await stepContext.prompt(TEXT_PROMPT,
                {
                    prompt: promptMsg,
                    retryPrompt: repromptMsg,
                });

        } else {
            return await stepContext.next(owner);
        }


    }

    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const destination = stepContext.result;
        return await stepContext.endDialog(destination);
    }
}
