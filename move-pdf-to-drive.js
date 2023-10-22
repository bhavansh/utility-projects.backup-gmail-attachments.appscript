function fetchStatementsToDrive() {
    var labelList = [
        "statements-bank-hdfc",
        "statements-bank-idfc",
        "statements-bank-pnb",
        "statements-nps",
        "statements-ecas",
        "statements-creditcard-hdfc",
        "statements-creditcard-icici",
    ];

    for (var i = 0; i < labelList.length; i++) {
        var label = labelList[i];
        Logger.log("Processing label: " + label);
        processLabelAttachments(label);
    }
}

function processLabelAttachments(label, skipIfExists = true) {
    Logger.log("Searching for emails with label: " + label);

    // Calculate the date one year ago from today
    var oneYearAgo = new Date();
    oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);

    // Convert the date to the format required by Gmail's search query
    var formattedDate = Utilities.formatDate(
        oneYearAgo,
        Session.getScriptTimeZone(),
        "yyyy/MM/dd"
    );

    // Modify the Gmail search query to fetch emails from the last year
    var query = "label:" + label + " after:" + formattedDate;
    var threads = GmailApp.search(query);

    threads.reverse();
    // Create the folder structure in Google Drive
    var parentFolder = createFolderStructure(label);

    for (var j = 0; j < threads.length; j++) {
        var messages = threads[j].getMessages();

        for (var k = 0; k < messages.length; k++) {
            var message = messages[k];

            Logger.log(
                "Processing email with subject: " + message.getSubject()
            );
            var attachments = message.getAttachments();

            // Create the folder structure in Google Drive
            var folder = createFolderStructure(label);

            for (var l = 0; l < attachments.length; l++) {
                var attachment = attachments[l];
                Logger.log("Processing attachment: " + attachment.getName());

                if (isValidAttachment(label, attachment.getName())) {
                    // Check if the file already exists and skip it if it does
                    if (
                        skipIfExists &&
                        fileExistsInFolder(folder, attachment.getName())
                    ) {
                        Logger.log(
                            "File with the same name already exists. Skipping."
                        );
                        continue;
                    }

                    // Create the file inside the folder
                    var pdfFile = Utilities.newBlob(
                        attachment.getBytes(),
                        "application/pdf",
                        attachment.getName()
                    );
                    var file = folder.createFile(pdfFile);
                    Logger.log(
                        "Saved attachment to subfolder: " + file.getName()
                    );
                } else {
                    Logger.log(
                        "Invalid Attachment, Valdation Failed: " +
                            attachment.getName()
                    );
                }
            }
        }
    }
}

function isValidAttachment(label, attachment_name) {
    if (label === "statements-bank-hdfc")
        return (
            attachment_name.includes("Bhavansh") &&
            attachment_name.toLowerCase().includes(".pdf")
        );
    return attachment_name.toLowerCase().includes(".pdf");
}

function fileExistsInFolder(folder, fileName) {
    var files = folder.getFilesByName(fileName);
    return files.hasNext();
}

function createFolderStructure(label) {
    var parentFolder = DriveApp.getRootFolder();
    var labelParts = label.split("-");

    for (var i = 0; i < labelParts.length; i++) {
        var folderName = labelParts[i];
        Logger.log("Checking for folder: " + folderName);
        var existingFolder = checkFolderExists(parentFolder, folderName);

        if (existingFolder) {
            parentFolder = existingFolder;
        } else {
            parentFolder = parentFolder.createFolder(folderName);
            Logger.log("Created folder: " + folderName);
        }
    }

    return parentFolder;
}

function checkFolderExists(parentFolder, folderName) {
    var folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
        return folders.next();
    }
    return null;
}
