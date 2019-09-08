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
import { sp, ItemAddResult } from '@pnp/sp';

const IntroCard = require('../../resources/intro.json');
import { SiteDialog } from "./siteDialog";
import { SiteUserProps } from "@pnp/sp/src/siteusers";

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const SITE_DIALOG = 'siteDialog';
const TEAMS_DIALOG = 'teamsDialog';
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
        attachments: [CardFactory.adaptiveCard(IntroCard)]
      });

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
      const details = stepContext.result as SiteDetails;
      
      const result: ItemAddResult = await this.createSiteRequest(details);

      const msg = `I have created a site request for you with title: ${details.title} and template: ${details.siteType} with owner: ${details.owner}.`;
      await stepContext.context.sendActivity(msg);

    } else {

      await stepContext.context.sendActivity('Thank you.');

    }
    return await stepContext.endDialog();

  }

  private async createSiteRequest(details: SiteDetails): Promise<ItemAddResult> {

    const user: SiteUserProps  = await sp.web.siteUsers.getByEmail(details.owner).get();
    
    const addResult:ItemAddResult = await sp.web.lists.getByTitle('Sites').items.add(
      {
        Title: details.title,
        IsBot: true,
        Alias: details.alias,
        SiteType: details.siteType,
        OwnerId: user.Id,
        ConversationId: details.conversationId 
      }
    );
    
    return addResult;
  }
}

