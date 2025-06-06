/**
 * Sends an HTML email to the nominee when a Google Form is submitted.
 */
function sendEmailToNominee(e) {
    console.log("Form submit event object (e): " + JSON.stringify(e));
    try {
        var formResponse = e.namedValues;

        if (typeof formResponse === 'undefined' && e.response && typeof e.response.getItemResponses === 'function') {
            console.log("e.namedValues is undefined. Attempting to use e.response.getItemResponses().");
            var itemResponses = e.response.getItemResponses();
            formResponse = {};
            for (var i = 0; i < itemResponses.length; i++) {
                var itemResponse = itemResponses[i];
                var questionTitle = itemResponse.getItem().getTitle();
                var answer = itemResponse.getResponse();
                if (Array.isArray(answer)) {
                    formResponse[questionTitle] = answer.length > 0 ? [String(answer[0])] : [];
                } else {
                    formResponse[questionTitle] = [String(answer)];
                }
            }
            console.log("Reconstructed formResponse: " + JSON.stringify(formResponse));
        }

        if (!formResponse) {
            console.error("Critical Error: formResponse is still undefined or null after attempting all recovery methods. Event object: " + JSON.stringify(e));
            return;
        }

        var nomineeEmailQuestionTitle = "Email of nominee";
        var reasonQuestionTitle = "Reason for nomination";
        var organizationName = "McKinnon Secondary College";

        var nomineeEmail = "";
        if (formResponse[nomineeEmailQuestionTitle] && formResponse[nomineeEmailQuestionTitle][0]) {
            nomineeEmail = formResponse[nomineeEmailQuestionTitle][0].trim();
        } else {
            console.error("Error: The question '" + nomineeEmailQuestionTitle + "' was not found in the form response or has no value. Please check your form.");
        }

        var reasonForNomination = "No specific reason was provided, but your work is appreciated!";
        if (formResponse[reasonQuestionTitle] && formResponse[reasonQuestionTitle][0]) {
            reasonForNomination = formResponse[reasonQuestionTitle][0];
        } else {
            console.log("Info: The question '" + reasonQuestionTitle + "' was not found or was empty. Using default reason.");
        }

        var nomineeFirstName = "";
        if (nomineeEmail && nomineeEmail !== "") {
            try {
                console.log("Attempting to retrieve nominee profile for: " + nomineeEmail);
                var nomineeUser = AdminDirectory.Users.get(nomineeEmail);
                if (nomineeUser && nomineeUser.name && nomineeUser.name.givenName) {
                    nomineeFirstName = nomineeUser.name.givenName;
                    console.log("Nominee first name found in directory: " + nomineeFirstName);
                } else if (nomineeUser) {
                    console.log("Nominee user found in directory, but givenName is missing. User object: " + JSON.stringify(nomineeUser.name));
                } else {
                    console.log("Nominee user not found in directory for: " + nomineeEmail);
                }
            } catch (err) {
                console.log("Could not retrieve nominee's profile or first name from directory for " + nomineeEmail + ". Error: " + err.message);
            }
        } else {
            console.log("Nominee email is empty, skipping directory lookup for first name.");
        }

        var nominatorEmail = "an anonymous nominator";
        if (e && e.response && typeof e.response.getRespondentEmail === 'function') {
            var respondentEmail = e.response.getRespondentEmail();
            if (respondentEmail) {
                nominatorEmail = respondentEmail;
            } else {
                console.log("Info: Respondent email was not collected or is empty. Using default nominator.");
            }
        } else {
            console.error("Error: Unable to retrieve respondent email. e.response or e.response.getRespondentEmail is not available. Please check form settings to ensure email collection is enabled.");
        }

        var nominatorDisplayName = nominatorEmail;
        var nominatorPhotoUrl = null;

        try {
            console.log("Attempting to retrieve nominator profile for: " + nominatorEmail);
            var nominatorUser = AdminDirectory.Users.get(nominatorEmail, { viewType: 'domain_public' });
            if (nominatorUser) {
                if (nominatorUser.name && nominatorUser.name.fullName) {
                    nominatorDisplayName = nominatorUser.name.fullName;
                } else if (nominatorUser.name && nominatorUser.name.givenName) {
                    nominatorDisplayName = nominatorUser.name.givenName;
                    if (nominatorUser.name.familyName) {
                        nominatorDisplayName += " " + nominatorUser.name.familyName;
                    }
                }
                if (nominatorUser.thumbnailPhotoUrl) {
                    nominatorPhotoUrl = nominatorUser.thumbnailPhotoUrl;
                }
            }
        } catch (err) {
            console.log("Could not retrieve nominator details for " + nominatorEmail + ". Error: " + err.message);
        }

        if (!nomineeEmail) {
            console.error("Nominee email not provided. Question title used: '" + nomineeEmailQuestionTitle + "'.");
            return;
        }

        var name = "Jazz the JoyBot";
        var subject = "🥳 JoyBot Nomination! 🥳";
        var htmlTemplate = HtmlService.createTemplateFromFile("template.html");

        htmlTemplate.nomineeFirstName = nomineeFirstName;
        htmlTemplate.nominatorEmail = nominatorEmail;
        htmlTemplate.nominatorDisplayName = nominatorDisplayName;
        htmlTemplate.nominatorPhotoUrl = nominatorPhotoUrl;
        htmlTemplate.reasonForNomination = reasonForNomination.replace(/\n/g, '<br>');
        htmlTemplate.organizationName = organizationName;

        var htmlBody = htmlTemplate.evaluate().getContent();

        MailApp.sendEmail({
            name: name,
            to: nomineeEmail,
            subject: subject,
            htmlBody: htmlBody,
            noReply: true
        });

        console.log("Nomination email sent to: " + nomineeEmail);

    } catch (error) {
        console.error("Error in sendEmailToNominee: " + error.toString());
        console.error("Error stack: " + error.stack);
    }
}

/**
 * Helper function to test the email sending functionality with sample data.
 * You can run this manually from the Apps Script editor.
 */
function testSendEmail() {
    var mockEvent = {
        namedValues: {
            "Email of nominee": ["sam.neal@mckinnonsc.vic.edu.au"],
            "Name of nominee": ["Sam Neal"],
            "Reason for nomination": ["Being wonderfully supportive and helpful."]
        },
        response: {
            getRespondentEmail: function () {
                return "jarryd.steadman@mckinnonsc.vic.edu.au";
            }
        }
    };

    sendEmailToNominee(mockEvent);
    console.log("Test email function executed. Check the recipient's inbox and Apps Script logs.");
}
