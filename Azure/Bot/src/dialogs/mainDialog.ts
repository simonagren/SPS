import {
  ComponentDialog,
  WaterfallDialog,
  DialogState,
  DialogSet,
  DialogTurnStatus,
  WaterfallStepContext,
  DialogTurnResult,
  TextPrompt,
  OAuthPrompt
} from "botbuilder-dialogs";
import { Logger } from "../logger";
import { TurnContext, StatePropertyAccessor, CardFactory } from "botbuilder";
import { SiteDetails } from "./siteDetails";
import { LuisHelper } from "./luisHelper";
import { SiteDialog } from "./siteDialog";

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const SITE_DIALOG = 'siteDialog';
const TEAMS_DIALOG = 'teamsDialog';
const OAUTH_PROMPT = 'oAuthPrompt';
const TEXT_PROMPT = "textPrompt";

export class MainDialog extends ComponentDialog {
  private logger: Logger;

  constructor(logger: Logger) {
    super('MainDialog');

    if (!logger) {
      logger = console as Logger;
      logger.log('[MainDialog]: logger not passed in, defaulting to console');
    }

    this.logger = logger;


    this
      .addDialog(new TextPrompt(TEXT_PROMPT))
      .addDialog(new SiteDialog(SITE_DIALOG, this.logger))
      .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
        this.introStep.bind(this),
        this.mainStep.bind(this),
        this.finalStep.bind(this),
      ]));

    this.initialDialogId = MAIN_WATERFALL_DIALOG;
  }

  /**
   * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
   * If no dialog is active, it will start the default dialog.
   * @param {TurnContext} context
   */
  public async run(context: TurnContext, accessor: StatePropertyAccessor<DialogState>) {

    const dialogSet = new DialogSet(accessor);
    dialogSet.add(this);

    const dialogContext = await dialogSet.createContext(context);

    const results = await dialogContext.continueDialog();
    if (results.status === DialogTurnStatus.empty) {
      await dialogContext.beginDialog(this.id);
    }
  }

  private async introStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {

      if (!process.env.LuisAppId || !process.env.LuisAPIKey || !process.env.LuisAPIHostName) {
        await stepContext.context.sendActivity('NOTE: LUIS is not configured. To enable all capabilities, add `LuisAppId`, `LuisAPIKey` and `LuisAPIHostName` to the .env file.');
        return await stepContext.next();
      }
      
      await stepContext.context.sendActivity({
        attachments: [CardFactory.adaptiveCard(
          {
            "type": "AdaptiveCard",
            "body": [
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "size": "Medium",
                            "weight": "Bolder",
                            "text": "Hi I'm Provisioning Bot"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "style": "Person",
                                            "url": "https://docs.botframework.com/static/devportal/client/images/bot-framework-default.png",
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
                                            "text": "Today I will help you create a SharePoint site or a Microsoft Team. Write something simple or a full sentence and I will extract some of the details you provided.",
                                            "wrap": true
                                        }
                                    ],
                                    "width": "stretch"
                                }
                            ]
                        }
                    ]
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "size": "Medium",
                            "weight": "Bolder",
                            "text": "Start with something like:",
                            "wrap": true
                        }
                    ]
                },
                {
                    "type": "TextBlock",
                    "text": "- 'Create a site'\n- 'Create Microsoft Teams'\n- 'Communications site' or 'modern team site'\n- 'Create a modern team site with title <title> and owner <name@mail.com>\n"
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Create a SharePoint site",
                    "data": {
                        "messageText": "Create a SharePoint site",
                        "isFromAdaptiveCard": true
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Create a Microsoft Teams team",
                    "data": {
                        "messageText": "Create a Microsoft Teams team",
                        "isFromAdaptiveCard": true
                    }
                }
            ],
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.0"
        }
        )]
      }
      );

      return await stepContext.prompt(TEXT_PROMPT, { prompt: '' });
    
  }

  /**
   * Main step in the waterall.  This will use LUIS to attempt to extract the title, site type, and owner.
   * Then, it hands off to the siteDialog child dialog to collect any remaining details.
   */
  private async mainStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    this.logger.log('[MainDialog]: Inside main step');
    let siteDetails = new SiteDetails();

    if (process.env.LuisAppId && process.env.LuisAPIKey && process.env.LuisAPIHostName) {
      // Call LUIS and gather any potential site details.
      // This will attempt to extract the title, site type and owner from the user's message
      // and will then pass those values into the site dialog
      siteDetails = await LuisHelper.executeLuisQuery(this.logger, stepContext.context);

      this.logger.log('LUIS extracted these site details:', siteDetails);
    }

    // In this sample we only have a single intent we are concerned with. However, typically a scenario
    // will have multiple different intents each corresponding to starting a different child dialog.

    // Run the BookingDialog giving it whatever details we have from the LUIS call, it will fill out the remainder.
    if (siteDetails.intent === "Create_site") {
      return await stepContext.beginDialog(SITE_DIALOG, siteDetails);
    } else if (siteDetails.intent === "Create_teams") {
      return await stepContext.beginDialog(TEAMS_DIALOG, siteDetails);
    } else {
      await stepContext.context.sendActivity('Couldn\'t establish intent, cancelling');
      return await stepContext.endDialog();
    }

  }

  /**
   * This is the final step in the main waterfall dialog.
   * It wraps up the sample "book a flight" interaction with a simple confirmation.
   */
  private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    // If the child dialog ("siteDialog") was cancelled or the user failed to confirm, the Result here will be null.
    if (stepContext.result) {
      const result = stepContext.result as SiteDetails;
      // Now we have all the site details.



      // This is where we make the calls to the Azure Function or Graåh

      // If the call to the booking service was successful tell the user.

      const msg = `I have create a site for you with title: ${result.title} and template: ${result.siteType} with owner: ${result.owner}.`;
      await stepContext.context.sendActivity(msg);

    } else {

      await stepContext.context.sendActivity('Thank you.');

    }
    return await stepContext.endDialog();

  }
}

