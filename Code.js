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

        var nomineeEmailQuestionTitle = "Nominee username(s)";
        var reasonQuestionTitle = "Reason for nomination";
        var organizationName = "McKinnon Secondary College";

        var nomineeUsernamesRaw = "";
        if (formResponse[nomineeEmailQuestionTitle] && formResponse[nomineeEmailQuestionTitle][0]) {
            nomineeUsernamesRaw = formResponse[nomineeEmailQuestionTitle][0].trim();
        } else {
            console.error("Error: The question '" + nomineeEmailQuestionTitle + "' was not found in the form response or has no value. Please check your form.");
            return;
        }

        var nomineeUsernames = nomineeUsernamesRaw.split(/[,\s]+/).filter(function (username) {
            return username.length > 0;
        });

        if (nomineeUsernames.length === 0) {
            console.error("No nominee usernames found in the input. Aborting.");
            return;
        }

        var potentialNomineeEmails = nomineeUsernames.map(function (username) {
            if (username.indexOf('@') === -1) {
                return username + "@mckinnonsc.vic.edu.au";
            }
            return username;
        });

        var reasonForNomination = "No specific reason was provided, but your work is appreciated!";
        if (formResponse[reasonQuestionTitle] && formResponse[reasonQuestionTitle][0]) {
            reasonForNomination = formResponse[reasonQuestionTitle][0];
        } else {
            console.log("Info: The question '" + reasonQuestionTitle + "' was not found or was empty. Using default reason.");
        }

        var validNominees = [];
        if (potentialNomineeEmails && potentialNomineeEmails.length > 0) {
            potentialNomineeEmails.forEach(function (email) {
                try {
                    console.log("Attempting to validate nominee: " + email);
                    var nomineeUser = AdminDirectory.Users.get(email, { viewType: 'domain_public' });
                    if (nomineeUser && nomineeUser.name && nomineeUser.name.givenName) {
                        validNominees.push({
                            email: email,
                            firstName: nomineeUser.name.givenName,
                            photoUrl: nomineeUser.thumbnailPhotoUrl || null
                        });
                        console.log("Successfully validated nominee: " + nomineeUser.name.givenName);
                    } else {
                        console.log("Nominee user found for " + email + ", but givenName is missing. Excluding. User object: " + JSON.stringify(nomineeUser.name));
                    }
                } catch (err) {
                    console.log("Could not retrieve nominee's profile for " + email + ". This user will be excluded. Error: " + err.message);
                }
            });
        }

        if (validNominees.length === 0) {
            console.error("No valid nominees found after checking the directory. No email will be sent.");
            return;
        }

        var nomineeEmails = validNominees.map(function (nominee) { return nominee.email; });
        var nomineeFirstNames = validNominees.map(function (nominee) { return nominee.firstName; });

        var nomineeGreeting;
        if (nomineeFirstNames.length > 1) {
            var allButLast = nomineeFirstNames.slice(0, -1);
            var last = nomineeFirstNames.slice(-1)[0];
            nomineeGreeting = allButLast.join(', ') + ' & ' + last;
        } else if (nomineeFirstNames.length === 1) {
            nomineeGreeting = nomineeFirstNames[0];
        } else {
            // This case should not be reached due to the check for validNominees.length === 0, but as a fallback.
            nomineeGreeting = 'there';
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

        if (nomineeEmails.length === 0) {
            console.error("Nominee email not provided. Question title used: '" + nomineeEmailQuestionTitle + "'.");
            return;
        }

        var name = "Jazz the JoyBot";
        var subject = "🥳 JoyBot Nomination! 🥳";
        var htmlTemplate = HtmlService.createTemplateFromFile("template.html");

        htmlTemplate.nomineeFirstName = nomineeGreeting;
        htmlTemplate.nominees = validNominees;
        htmlTemplate.nominatorEmail = nominatorEmail;
        htmlTemplate.nominatorDisplayName = nominatorDisplayName;
        htmlTemplate.nominatorPhotoUrl = nominatorPhotoUrl;
        htmlTemplate.reasonForNomination = reasonForNomination.replace(/\n/g, '<br>');
        htmlTemplate.organizationName = organizationName;

        var htmlBody = htmlTemplate.evaluate().getContent();

        MailApp.sendEmail({
            name: name,
            to: nomineeEmails.join(','),
            subject: subject,
            htmlBody: htmlBody,
            noReply: true
        });

        console.log("Nomination email sent to: " + nomineeEmails.join(','));

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
            "Nominee username(s)": ["sam.neal,jarryd.steadman"],
            "Name of nominee": ["Sam Neal"],
            "Reason for nomination": ["For being a great team and helping out with the recent project."]
        },
        response: {
            getRespondentEmail: function () {
                return "sam.neal@mckinnonsc.vic.edu.au";
            }
        }
    };

    sendEmailToNominee(mockEvent);
    console.log("Test email function executed. Check the recipient's inbox and Apps Script logs.");
}
