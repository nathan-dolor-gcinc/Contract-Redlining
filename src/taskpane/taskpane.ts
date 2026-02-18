/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Assign event handlers and other initialization logic.
    document.getElementById("run-agent").onclick = () => tryCatch(runAgent);

    document.getElementById("app-body").style.display = "flex";
  }
});

async function insertParagraph() {
    await Word.run(async (context) => {

        const docBody = context.document.body;
        docBody.insertParagraph("Office has several versions, including Office 2021, Microsoft 365 subscription, and Office on the web.",
        Word.InsertLocation.start);
        await context.sync();
    });
}

async function runAgent() {
    const prompt = (document.getElementById("agent-input") as HTMLTextAreaElement).value;

    const response = await fetch("https://your-backend.com/agent", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ prompt })
    }).then(r => r.json());

    await Word.run(async (context) => {
        context.document.body.insertParagraph(
            response.text,
            Word.InsertLocation.end
        );
        await context.sync();
    });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}