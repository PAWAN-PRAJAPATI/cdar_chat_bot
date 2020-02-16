// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory, CardFactory } = require('botbuilder');
const { QnAMaker } = require('botbuilder-ai');

class SuggestedActionsBot extends ActivityHandler {
    constructor() {
        super();
     
        try {
            this.qnaMaker = new QnAMaker({
                knowledgeBaseId: process.env.QnAKnowledgebaseId,
                endpointKey: process.env.QnAEndpointKey,
                host: process.env.QnAEndpointHostName
            });
        } catch (err) {
            console.warn(`QnAMaker Exception: ${ err } Check your QnAMaker configuration in .env`);
        }

        this.onMembersAdded(async (context, next) => {
            await this.sendWelcomeMessage(context);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMessage(async (context, next) => {
            const text = context.activity.text;

            // Create an array with the valid color options.
            const validColors =['PSRO Manual', 'Work Flow Manual', 'GSR'];

            // If the `text` is in the Array, a valid color was selected and send agreement.
            if (validColors.includes(text)) {

                if(validColors.includes(text)){

                    this.attachKB(context, text)
                }

                await context.sendActivity(`Knowledge base set to , ${ text } `);

            }

            else{

                const qnaResults = await this.qnaMaker.getAnswers(context);

                if(qnaResults.length>0){

                    for(var ind = 0; ind < qnaResults.length; ind++){
                        //await context.sendActivity(qnaResults[ind].questions[0]);
                        //await context.sendActivity(qnaResults[ind].answer);
                        //await context.sendActivity(qnaResults[ind].metadata[0]);
                        await context.sendActivity({ attachments: [this.createCard(qnaResults[ind].questions[0], qnaResults[ind].answer,"https:"+qnaResults[ind].metadata[0].value)] });
                    }
                }
                else {

                    await context.sendActivity('No QnA Maker answers were found.');
                }
    
                // After the bot has responded send the suggested actions.
                //await this.sendSuggestedActions(context);
    
                // By calling next() you ensure that the next BotHandler is run.
                await next();
            }
           
        });
    }

    /**
     * Send a welcome message along with suggested actions for the user to click.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendWelcomeMessage(turnContext) {
        const { activity } = turnContext;

        // Iterate over all new members added to the conversation.
        for (const idx in activity.membersAdded) {
            if (activity.membersAdded[idx].id !== activity.recipient.id) {
                const welcomeMessage = `Welcome to suggestedActionsBot ${ activity.membersAdded[idx].name }. ` +
                    'This bot will introduce you to Suggested Actions. ' +
                    'Please select an option:';
                await turnContext.sendActivity(welcomeMessage);
                await this.sendSuggestedActions(turnContext);
            }
        }
    }

    createCard(header,content, url){

        return CardFactory.thumbnailCard(
            header,
            [{ url: '' }],
            [{
                type: 'openUrl',
                title: 'Go to docs',
                value: url
            }],
            {
                subtitle: '',
                text: content
            }
        );

    }

    async attachKB(turnContext, text){

        const options = { top: 15, scoreThreshold : 0.4 };



        if(text=="PSRO"){

            try {
                this.qnaMaker = new QnAMaker({
                    knowledgeBaseId: process.env.PsroQnAKnowledgebaseId,
                    endpointKey: process.env.PsroQnAEndpointKey,
                    host: process.env.PsroQnAEndpointHostName
                },options);
            } catch (err) {
                console.warn(`QnAMaker Exception: ${ err } Check your QnAMaker configuration in .env`);
            }
        }
        else{

            try {
                this.qnaMaker = new QnAMaker({
                    knowledgeBaseId: process.env.WorkQnAKnowledgebaseId,
                    endpointKey: process.env.WorkQnAEndpointKey,
                    host: process.env.WorkQnAEndpointHostName
                },options);
            } catch (err) {
                console.warn(`QnAMaker Exception: ${ err } Check your QnAMaker configuration in .env`);
            }

        }

    }

    /**
     * Send suggested actions to the user.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendSuggestedActions(turnContext) {
        var reply = MessageFactory.suggestedActions(['PSRO Manual', 'Work Flow Manual', 'GSR']);
        await turnContext.sendActivity(reply);
    }
}

module.exports.SuggestedActionsBot = SuggestedActionsBot;
