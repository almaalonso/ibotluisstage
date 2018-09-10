/*-----------------------------------------------------------------------------
A simple Language Understanding (LUIS) bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var request = require("request");

//var botbuilder_azure = require("botbuilder-azure");

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
    console.log('%s listening to %s', server.name, server.url);
});

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

/*----------------------------------------------------------------------------------------
* Bot Storage: This is a great spot to register the private state storage for your bot. 
* We provide adapters for Azure Table, CosmosDb, SQL Azure, or you can implement your own!
* For samples and documentation, see: https://github.com/Microsoft/BotBuilder-Azure
* ---------------------------------------------------------------------------------------- */

//var tableName = 'botdata';
//var azureTableClient = new botbuilder_azure.AzureTableClient(tableName, process.env['AzureWebJobsStorage']);
//var tableStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, azureTableClient);


// Create your bot with a function to receive messages from the user
// This default message handler is invoked if the user's utterance doesn't
// match any intents handled by other dialogs.
var bot = new builder.UniversalBot(connector, function (session, args) {
    session.send('You reached the default message handler. You said \'%s\'.', session.message.text);
});

// Make sure you add code to validate these fields -- LUIS fields
//var luisAppId = process.env.LuisAppId;
//var luisAPIKey = process.env.LuisAPIKey;
var luisAppId = "0e55a2a1-8e26-4c95-89f9-bacd12bc4dac";
var luisAPIKey = "6961a5199e1f4d8090fae590525727c7";
var luisAPIHostName = process.env.LuisAPIHostName || 'westus.api.cognitive.microsoft.com';

const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v2.0/apps/' + luisAppId + '?subscription-key=' + luisAPIKey;

// Create a recognizer that gets intents from LUIS, and add it to the bot
var recognizer = new builder.LuisRecognizer(LuisModelUrl);
bot.recognizer(recognizer);

// Add a dialog for each intent that the LUIS app recognizes.
// See https://docs.microsoft.com/en-us/bot-framework/nodejs/bot-builder-nodejs-recognize-intent-luis 

//ACCESS DIALOG
bot.dialog('AccessDialog',
    function (session, args, next) {
        var intent = args.intent;
        var divisionEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Access.Division');
        var focalGroupEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Access.Focal');
        var administratorEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Access.Administrator');
        var corporateSitesEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Access.Corporate');

        if (corporateSitesEntity) {
            session.send("Ask the access to the Power User of your Division. If you don't know who are your Power Users please ask to your COF.");
        }
        else if (administratorEntity) {
            storeSessionData(session, 'SISCOR 365', 'Need administrator access to Siscor365', 'Need administrator access to Siscor365 -  sent from chatbot');

            session.send('If you need this access please create a Work Order to SisCor365 Support Team');
            session.beginDialog('requestTicketDialog');
        }
        else if (focalGroupEntity) {
            session.send("Also called COF or Division Specialists. They are in charge of granting access due their scope as administrators. Approval of the  list of division's focal group is required. If you need this access please request it to you IT Specialist");
        }
        else if (divisionEntity) {
            session.send('Please request the access to your IT Specialist of your COF');
        }
        else {

        }
    }
).triggerAction({
    matches: 'Siscor.Access'
    });


//OUTPUTS DIALOG
bot.dialog('OutputsDialog',
    function (session, args, next) {
        var intent = args.intent;
        var addEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Output.Add');
        var incompleteEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Output.Incomplete');
        var readnOnlyEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Output.ReadOnly');
        var relateEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Inputs.Relate');
        var profileDataEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Output.ProfileData');
        var userEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.PwrUser.Usr');
        var cancelEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Output.Cancel');
        var pauseEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Output.Pause');
        var editEntity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Output.Edit');

        if (editEntity && userEntity) {
            session.send('People allowed to edit outputs are: the author, the assistant, the power user or the administrator only if the output is paused.');
            session.endDialog();
        }
        else if (editEntity) {
            session.send('People allowed to edit outputs are: the author, the assistant, the power user or the administrator only if the output is paused.');
            session.endDialog();
        }
        else if (pauseEntity && userEntity) {
            session.send("Any user is allowed to pause an output assigned to them. Also it's possible to pause an output when this is generated and approved.");
            session.endDialog();
        }
        else if (pauseEntity) {
            session.send("Any user is allowed to pause an output assigned to them. Also it's possible to pause an output when this is generated and approved.");
            session.endDialog();
        }
        else if (cancelEntity && userEntity) {
            session.send("Any user is allowed to cancel an output assigned to them.");
            session.endDialog();
        }
        else if (cancelEntity) {
            session.send("Any user is allowed to cancel an output assigned to them.");
            session.endDialog();
        }
        else if (profileDataEntity) {
            session.send("When an output is generated from SisCor365, it takes the information that is entered in an input except the information from EzShare library  in case it is related to an operation");
            session.endDialog();
        }
        else if (relateEntity) {
            session.send("It's possible to relate an output with several inputs. Relate inputs will complete when the outputs are complete");
            session.endDialog();
        }
        else if (readnOnlyEntity) {
            session.send("In status Approved or waiting for approval, the document access has read only access for everyone on assigned ");
            session.endDialog();
        }
        else if (incompleteEntity) {
            storeSessionData(session, 'SISCOR 365', 'The input does not complete even though the output does', 'The input does not complete even though the output does -  sent from chatbot');

            session.send("In case an input does not complete when the output is completed, please create a ticket for SisCor365 support team ");
            session.beginDialog('requestTicketDialog');
        }
        else if (addEntity) {
            session.send("You can add more than one person to the list of approvers. But remember that you can add one person to sign and this person must be an approver");
            session.endDialog();
        }
    }
).triggerAction({
    matches: 'Siscor.Outputs'
});


//USE OF TABLE DIALOG
bot.dialog('UseOfTabletDialog',
    function (session, args, next) {
        session.send('SisCor365 works online allowing the users to work from a tablet. I do not recommend using the system in a mobile device. ');
        session.endDialog();
    }
).triggerAction({
    matches: 'Siscor.UseTablet'
});

//GREETING DIALOG
bot.dialog('GreetingDialog',
    function (session, args, next) {
        var intent = args.intent;
        var morning = builder.EntityRecognizer.findEntity(intent.entities, 'Greeting.Morning');
        var afternoon = builder.EntityRecognizer.findEntity(intent.entities, 'Greeting.Afternoon');
        var evening = builder.EntityRecognizer.findEntity(intent.entities, 'Greeting.Evening');

        if (morning) {
            session.send("Hi " + session.message.user.name)
            session.send("What's good about it (angry)?");
            session.endDialog();
        }
        else if (afternoon) {
            session.send('Hello ' + session.message.user.name + ', good afternoon. Welcome to iBot. How can I help you?');
            session.endDialog();
        }
        else if (evening) {
            session.send('Hello ' + session.message.user.name + ', good evening. Welcome to iBot. How can I help you?');
            session.endDialog();
        }
        else {
            session.send('Hello ' + session.message.user.name + ', Welcome to iBot. How can I help you?');
            session.endDialog();
        }

    }
).triggerAction({
    matches: 'Greeting'
});

bot.dialog('HelpDialog',
    (session) => {
        session.send('You reached the Help intent. You said \'%s\'.', session.message.text);
        session.endDialog();
    }
).triggerAction({
    matches: 'Help'
});

bot.dialog('CancelDialog',
    (session) => {
        session.send('You reached the Cancel intent. You said \'%s\'.', session.message.text);
        session.endDialog();
    }
).triggerAction({
    matches: 'Cancel'
});

//INPUTS DIALIOG
bot.dialog('InputsDialog',
    function (session, args, next) {
        var intent = args.intent;
        var cannotAttach = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Inputs.Cannot');
        var relateInput = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Inputs.Relate');

        if (relateInput) {
            storeSessionData(session, 'SISCOR 365', 'The input does not have an action', 'The input does not have an action -  sent from chatbot');

            session.send("You can decide later if you relate the input assigned to an operation. In case the input does not have an action, contact the SisCor365 suppor team");
            session.beginDialog('requestTicketDialog');
        }
        else if (cannotAttach) {
            storeSessionData(session, 'SISCOR 365', 'Not possible to attach a file to an input', 'Not possible to attach a file to an input -  sent from chatbot');

            session.send("In case you are not able to attach a file to an input please create a WRK for SisCor365 Team Support so they can take a look at your issue.");
            session.beginDialog('requestTicketDialog');
        }
        else {
            session.send("I couldn't understand what you wanted to say. Please try again");
            session.endDialog();
        }
    }

).triggerAction({
    matches: 'Siscor.Inputs'
});

//FRECUENT CONTACTS DIALOG
bot.dialog('FrecuentContactsDialog',
    function (session, args, next) {
        var intent = args.intent;
        var frecuentContact = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FrecCont');

        if (frecuentContact) {
            session.send('Siscor365 does not maintain your frecuent contacts with whom you work with . You must enter the contact everytime you work on SisCor365');
            session.endDialog();
        }
    }
).triggerAction({
    matches: 'Siscor.FrecuentContacts'
});

//FILE UPLOADS DIALOG
bot.dialog('FileUploadsDialog',
    function (session, args, next) {
        var intent = args.intent;
        var emailUpload = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FileUploads.Emails');
        var uploadCapacity = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FileUploads.Capacity');
        var fileTypes = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FileUploads.Types');
        var emailAttachment = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FileUploads.EmailsAttach');
        var docLess = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FileUploads.Less');
        var docGreater = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FileUploads.Greater');
        var fileError = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.FileUploads.Error');

        if (emailUpload) {
            session.send("It's possible to upload an email directly from Outlook to SisCor365. You just need to drag and drop the correspondence from Outlook. The upload will include email's attachments. For more informacion please click [here](https://idbg.sharepoint.com/sites/BID365/Documents/SISCOR365/How%20Tos/ENG/How%20to%20drag%20and%20drop%20correspondence%20from%20Outlook.pdf)");
            session.endDialog();
        }
        else if (uploadCapacity) {
            session.beginDialog('capacityAttachDialog');
        }
        else if (fileTypes) {
            session.send('Siscor365 is compatible with any document format');
        }
        else if (emailAttachment) {
            session.send('Emails that are dragged from Outlook to SisCor365 are uploaded with their attachments. Everything is uploaded ');
            session.endDialog();
        }
        else if (docLess) {
            session.send('For documets less than 100 MB you can upload them in SisCor365 ');
            session.endDialog();
        }
        else if (docGreater) {
            session.send("For documents greater than 100 MB it's advisable to directly use EzShare for your COF mailbox");
            session.endDialog();
        }
        else if (fileError) {
            session.send("Before generating an input or output please verify the file does not contain special characters like [~\"#%&*:<;>;?/{|}]");
            session.endDialog();
        }
    }
).triggerAction({
    matches: 'Siscor.FileUploads'
    });


//CAPACITY OF ATTACHMENTS DIALOG
bot.dialog('capacityAttachDialog', [
    function (session) {
        builder.Prompts.choice(session, "What is the size of your document? ", "Less than 100 MB|Greater than 100 MB", { listStyle: 2 });
    },
    function (session, results) {

        if (results.response) {
            var choice = results.response.entity;
            switch (choice) {
                case 'Less than 100 MB':
                    session.send("For documets less than 100 MB you can upload them in SisCor365 ");
                    session.endDialog();
                    break;
                case 'Greater than 100 MB':
                    session.send("For documents greater than 100 MB it's advisable to directly use EzShare for your COF mailbox");
                    session.endDialog();
                    break;
                default:
                    session.send('Please select a choice from the options given');
                    beginDialog('capacityAttachDialog');
                    break;
            }
        }
    }
]);


//COMMENTS DIALOG
bot.dialog('commentsDialog',
    function (session, args, next) {
        var intent = args.intent;
        var privateComment = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Comments.Private');
        var publicComment = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Comments.Public');

        if (privateComment) {
            session.send('Private comments are those that will only be visible to users within the department.');
            session.endDialog();
        }
        else if (publicComment) {
            session.send('Public comments are those that will be visible to all users of the Bank');
            session.endDialog();
        }
        else {
            session.send('There are two types of comments: Privates and Publics comments');
            session.endDialog();
        }        
    }
).triggerAction({
    matches: 'Siscor.Comments'
    })

//ASSISTANTS DIALOG
bot.dialog('assistantsDialog',
    function(session, args, next){
        var intent = args.intent;
        var peopleAssigned = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Assistant.Assign');
        if (peopleAssigned) {
            session.send('Currently, two people can be assigned: one person and one assistant. It is not possible to assign more than two people. Instead you can use the notifications tab to notify this assignment to others');
            session.endDialog();
        }
        else {
            session.send('In case of working with correspondence related to an operation you will have the option to list the members of the operation. Currently, two people can be assigned: one person and one assistant. It is not possible to assign more than two people. Instead you can use the notifications tab to notify this assignment to others');
            session.endDialog();
        }
    }
).triggerAction({
    matches: 'Siscor.Assistants'
    })


//ADMINISTRATORS DIALOG
bot.dialog('administratoraDialog',
    function (session, args, next) {
        var intent = args.intent;
        var powerUsersScope = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.PwrUser.Scope');
        var powerUsersAccess = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.PwrUser.Access');
        var powerUsersUsr = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.PwrUser.Usr');

        if (powerUsersScope) {
            session.send('Power Users manage all inputs and outputs');
            session.endDialog();
        }
        else if (powerUsersAccess) {
            session.send('Power users will be able to manage the permissions within SISCOR, all inputs and outputs assigment. They also will be able to manage permissions within SisCor on corporate sites.');
            session.endDialog();
        }
        else if (powerUsersUsr) {
            session.send('Power users ');
            session.endDialog();
        }
        else {
            session.send('Power users ');
            session.endDialog();
        }
    }
).triggerAction({
    matches: 'Siscor.Administrators'
});


//SIGNATURES DIALOG
bot.dialog('signaturesDialog',
    function (session, args, next) {
        var intent = args.intent;
        var signatureWork = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Signature.Work');
        var signatureError = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Signature.Error');
        var signatureRequest = builder.EntityRecognizer.findEntity(intent.entities, 'Siscor.Signature.Request');


        if (signatureWork) {
            session.send('The signatures work getting the information from convergence where they are created and stored');
            session.endDialog();
        }
        else if (signatureError) {
            session.beginDialog('signatureErrors');
        }
        else if (signatureRequest) {
            storeSessionData(session, 'Convergence Support L1', 'Signature Request for Siscor365', 'Signature Request for Siscor365 -  sent from chatbot');

            session.send('Please create a ticket in SNOW for Convergence team');
            session.beginDialog('requestTicketDialog');
        }
        else {

        }
    }
).triggerAction({
    matches: 'Siscor.Signatures'
    })


//SIGNATURE ERROR DIALOG
bot.dialog('signatureErrors', [
    function (session) {
        builder.Prompts.choice(session, "please select an option according with your situation: ", "I don't have a signature|I already have a signature|I don't know", {listStyle: 2});
    },
    function (session, results) {
        
        if (results.response) {
            var error = results.response.entity;
            switch (error) {
                case "I don't have a signature":
                    storeSessionData(session, 'Convergence Support L1', 'Signature Request for Siscor365', 'Signature Request for Siscor365 -  sent from chatbot');

                    session.send('Request a signature to convergence team in SNOW');
                    session.beginDialog('requestTicketDialog');
                    break;
                case "I already have a signature":
                    storeSessionData(session, 'SISCOR 365', 'Error in Signature. Already have a Signature', 'Error in Signature. Already have a Signature -  sent from chatbot');
                   
                    session.send('In case you are getting an 404 Error and you already have a signature  please create a ticket in SNOW for the SisCor365 support team in order to take a look at the issue');
                    session.beginDialog('requestTicketDialog');
                    break;
                case "I don't know":
                    storeSessionData(session, 'Convergence Support L1', 'Signature Error.No idea if user has a signature already', 'Signature Error.No idea if user has a signature already - sent from chatbot' );

                    session.send("In case you are getting an 404 Error and you don't know if you have a signature yet please create a ticket in SNOW for the Convergence team in order to take a look at the issue");
                    session.beginDialog('requestTicketDialog');
                    break;
                default:
                    break;
            }
        }
        else {
            session.send('I didnt recognized the error');
            session.endDialog();
        }
    }
]);

//Store data in session variable
function storeSessionData(session, team, description, shortDescription) {
    session.privateConversationData.team = team;
    session.privateConversationData.description = description;
    session.privateConversationData.shortDescription = shortDescription;
}

//REQUEST TICKET DIALOG
bot.dialog('requestTicketDialog', [
    function (session) {
        builder.Prompts.confirm(session, "Would you like to create a ticket now? ");
    },
    function (session, args) {
        if (args.response) {
            var team = session.privateConversationData.team;
            var description = session.privateConversationData.description;
            var shortDescription = session.privateConversationData.shortDescription;

            createTicket(session, team, description, shortDescription);
        }
        else {
            session.beginDialog('moreHelpDialog');
        }
    }
]);

bot.dialog('moreHelpDialog', [
    function (session) {
        builder.Prompts.confirm(session, "Alright! Is there anything else I can help you with? ");
    },
    function (session, args) {
        if (args.response) {
            session.send('Sure! How can I help you?');
            session.endDialog();
        }
        else {
            session.send('Ok! Have a nice day!');
            session.endDialog();
        }
    }
]);



//Function that creates a ticket for Service Now
function createTicket(session, team, description, shortDescription) {

        session.send('Let me create a ticket. Please wait..');
        var newTeam = team;

        var options = {
            method: 'POST',
            url: 'https://iadbdev.service-now.com/api/now/table/u_work_order',
            qs: { sysparm_fields: 'number,sys_id,description' },
            headers:
                {
                    authorization: 'Basic U29mdHRla19hdXRvbWF0aW9uOlQxZ3IzRngx',
                    'content-type': 'application/json'
                },
            body:
                {
                    u_client: 'almaa',
                    u_action_required: 'create',
                    u_category: "Bank's Business Applications",
                    u_subcategory: 'Biztalk',
                    u_component: 'Other',
                    description: description,
                    assignment_group: 'Yarvis',
                    assigned_to: 'softtek_automation',
                    short_description: shortDescription
                },
            json: true
        };

        request(options, function (error, response, body) {
            if (error) throw new Error(error);
            var ticketNum = body.result.number;
            var sys_id = body.result.sys_id;
            console.log('ticketNum is:' + ticketNum);
            //session.endDialog();
            //session.endConversation();

            //calling function to reassing the ticket to the correct team
            reassignTicket(session, sys_id, newTeam);
        });
        
    }


function reassignTicket(session, sys_id, newTeamAssignment) {
    var newTeam = newTeamAssignment;
    var options = {
        method: 'PUT',
        url: 'https://iadbdev.service-now.com/api/now/table/u_work_order/'+sys_id,
        qs: { sysparm_fields: 'number,sys_id,description' },
        headers:
            {
                authorization: 'Basic U29mdHRla19hdXRvbWF0aW9uOlQxZ3IzRngx',
                'content-type': 'application/json'
            },
        body:
            {
                assignment_group: newTeam,
                assigned_to: '',
            },
        json: true
    };


    request(options, function (error, response, body) {
        if (error) throw new Error(error);
        var ticketNum = body.result.number;
        var sys_id = body.result.sys_id;
        console.log('ticketNum is:' + ticketNum);
        session.send('Work Ticket created: ' + ticketNum);
        session.endDialog();
        //session.endConversation();
    });

}

