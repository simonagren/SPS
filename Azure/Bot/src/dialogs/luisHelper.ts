import { SiteDetails } from './siteDetails';

import { RecognizerResult, TurnContext } from 'botbuilder';

import { LuisRecognizer } from 'botbuilder-ai';
import { Logger } from '../logger';

export class LuisHelper {
    /**
     * Returns an object with preformatted LUIS results for the bot's dialogs to consume.
     * @param {Logger} logger
     * @param {TurnContext} context
     */
    public static async executeLuisQuery(logger: Logger, context: TurnContext): Promise<SiteDetails> {
        const siteDetails = new SiteDetails();

        try {
            process.env["NODE_TLS_REJECT_UNAUTHORIZED"] = "0";
            const recognizer = new LuisRecognizer({
                applicationId: process.env.LuisAppId,
                endpoint: `https://${ process.env.LuisAPIHostName }`,
                endpointKey: process.env.LuisAPIKey,
            }, {}, true);

            const recognizerResult = await recognizer.recognize(context);

            const intent = LuisRecognizer.topIntent(recognizerResult);

            siteDetails.intent = intent;

            if (intent === 'Create_site') {
                // We need to get the result from the LUIS JSON which at every level returns an array

                siteDetails.title = LuisHelper.parseTitleEntity(recognizerResult);
                siteDetails.siteType = LuisHelper.parseCompositeEntity(recognizerResult, 'Site', 'SiteType');
                siteDetails.owner = LuisHelper.parseEmailEntity(recognizerResult);
            } else if (intent === 'Create_teams') {
                // We need to get the result from the LUIS JSON which at every level returns an array
                siteDetails.title = LuisHelper.parseCompositeEntity(recognizerResult, 'Title', 'Title');
                siteDetails.siteType = LuisHelper.parseCompositeEntity(recognizerResult, 'Type', 'Site');
                siteDetails.owner = LuisHelper.parseCompositeEntity(recognizerResult, 'Owner', 'Email');
            }

        } catch (err) {
            logger.warn(`LUIS Exception: ${ err } Check your LUIS configuration`);
        }
        return siteDetails;
    }

    private static parseCompositeEntity(result: RecognizerResult, compositeName: string, entityName: string): string {
        const compositeEntity = result.entities[compositeName];
        if (!compositeEntity || !compositeEntity[0]) {
            return undefined;
        }

        const entity = compositeEntity[0][entityName];
        if (!entity || !entity[0]) {
            return undefined;
        }

        const entityValue = entity[0][0];
        return entityValue;
    }

    private static parseTitleEntity(result: RecognizerResult): string {
        const titleEntity = result.entities.Title;
        if (!titleEntity || !titleEntity[0]) {
            return undefined;
        }
        return titleEntity[0];
    }

    private static parseEmailEntity(result: RecognizerResult): string {
        const emailEntity = result.entities.email;
        if (!emailEntity || !emailEntity[0]) {
            return undefined;
        }
        return emailEntity[0];
    }

}
