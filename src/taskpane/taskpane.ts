/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-64.png";
import "../../assets/icon-80.png";

import * as transposer from "../transposer"

/* global document, Office, Word */

Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";

        // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
        if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
            console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
        }

        // Assign event handlers and other initialization logic.
        document.getElementById("run").onclick = run;
        document.getElementById("transpose-min").onclick = transposeMin;
        document.getElementById("transpose-plus").onclick = transposePlus;
    }
});

export function transposeMin() {
    var el = document.getElementById("transpose-level");
    if (el['value'] > -12) {
        el['value']--
    }
}

export function transposePlus() {
    var el = document.getElementById("transpose-level");
    if (el['value'] < 12) {
        el['value']++
    }
}


export async function run() {
    return Word.run(async context => {

        try {
            var doc = context.document;
            var originalRange = doc.getSelection();
            originalRange.load("text");

            await context.sync();

            var level = Number(document.getElementById("transpose-level")['value']);

            if (level > 0) {
                var result = transposer.transpose(originalRange.text).up(level).toString()
                originalRange.insertText(result, "Replace");
            } else if (level < 0) {
                var result = transposer.transpose(originalRange.text).down(Math.abs(level)).toString()
                originalRange.insertText(result, "Replace");
            }

            await context.sync();

        } catch (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }

        await context.sync();
    });
}
