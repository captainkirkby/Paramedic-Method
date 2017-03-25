/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */

 function onOpen(e) {
    DocumentApp.getUi().createAddonMenu()
        .addItem('Highlight', 'runParamedicMethod')
        .addItem('Unhighlight', 'removeParamedicMethod')
        .addToUi();
    }

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
 function onInstall(e) {
    onOpen(e);
}


function runParamedicMethod() {
    var count;
    var i;
    var weakVerbs = ['is', 'are', 'were', 'was', 'will', 'be', 'been', 'being', 'shall', 'am',
     'isn\'t', 'aren\'t', 'won\'t', 'wasn\'t', 'weren\'t'];

    var prepositions = ['at', 'in', 'on', 'to', 'of', 'for', 'as', 'with', 'by', 'from',
     'into', 'onto', 'than', 'that', 'under', 'over', 'toward', 'towards', 'until',
     'up', 'upon', 'within', 'without'];

    var expletives = ['it is observed that', 'it was observed that', 'I think that',
     'we think that', 'I believe that', 'we believe that', 'respectively', 'based off'];

    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();
    var nWeakVerbs = 0;
    var nPrepositions = 0;
    var weakVerbFraction;
    var prepositionFraction;

    var wordCount = body.getText().match(/\S+/g).length;

    text.appendText('\n\n======BEGIN REPORT======\n');
    for (i = 0; i < weakVerbs.length; i++) {
        word = weakVerbs[i];
        count = highlightText(word, '#FCFC00');
        if (count > 0)  {
            text.appendText('Highlighted \'' + word + '\' ' + count + ' times.\n');
        }
        nWeakVerbs += count;
    }

    for (i = 0; i < prepositions.length; i++) {
        word = prepositions[i];
        count = highlightText(word, '#00FC00');
        if (count > 0) {
            text.appendText('Highlighted \'' + word + '\' ' + count + ' times.\n');
        }
        nPrepositions += count;
    }

    for (i = 0; i < expletives.length; i++) {
        word = expletives[i];
        count = highlightText(word, '#FC0000');
        if (count > 0) {
            text.appendText('Highlighted \'' + word + '\' ' + count + ' times.\n');
        }
    }

    weakVerbFraction = nWeakVerbs / wordCount;
    prepositionFraction = nPrepositions / wordCount;

    weakVerbFraction = Math.round(weakVerbFraction * 10000) / 100;
    prepositionFraction = Math.round(prepositionFraction * 1000) / 10;

    text.appendText('\nYellow highlights indicate weak verbs.\nGreen highlights indicate prepositions.\nRed highlights indicate expletives.\nEdit colorful sentences.\n');
    text.appendText('Weak verb fraction = ' + weakVerbFraction.toString().match(/^-?\d+(?:\.\d{0,2})?/)[0] + '%. Strive for < 0.5%.\n');
    text.appendText('Preposition fraction = ' + prepositionFraction.toString().match(/^-?\d+(?:\.\d{0,2})?/)[0] + '%. Strive for < 10%.\n');
}

function removeParamedicMethod() {
    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();

    var re = new RegExp('\n\n======BEGIN REPORT======', 'gi');
    var foundElement = re.exec(body.getText());
    var start = foundElement.index;
    var end = text.getText().length - 1;

    text.deleteText(start, end);
    text.setBackgroundColor(0, start - 1, '#FFFFFF');
}

function highlightText(findMe, color) {
    var body = DocumentApp.getActiveDocument().getBody();
    var text = body.editAsText();

    var re = new RegExp('\\s' + findMe + '\\s', 'gi');

    var foundElement;
    var count = 0;

    while (foundElement = re.exec(body.getText())) {
        count++;
        // Get the text object from the element
        var foundText = foundElement[0];

        // Where in the Element is the found text?
        var start = foundElement.index;
        var end = start + foundText.length - 1;

        // Change the background color to yellow
        text.setBackgroundColor(start, end, color);

    }
    return count;
}