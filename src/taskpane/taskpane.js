/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// ref: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("comment").onclick = comment;
    document.getElementById("selectcomment").onclick = selectaicomment;
    document.getElementById("grammar").onclick = improvegrammar;
  }
});

import { Configuration, OpenAIApi } from "openai";
import { OPENAI_API_KEY, SAPLING_API_KEY } from "../../config";

const configuration = new Configuration({
  apiKey: OPENAI_API_KEY,
});
const openai = new OpenAIApi(configuration);

//new
import { Client } from "@saplingai/sapling-js/client";

const apiKey = SAPLING_API_KEY;
const client = new Client(apiKey);
client.edits("Lets get started!");

async function generateText(message) {
  const completion = await openai.createCompletion({
    model: "text-davinci-002",
    prompt: message,
    temperature: 0.6,
  });
  return completion.data.choices[0].text;
}

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

export async function comment() {
  // Set a comment on the selected content.
  return Word.run(async (context) => {
    const text = "abcd";
    const comment = context.document.getSelection("Hello World").insertComment(text);

    // Load object for display in Script Lab console.
    comment.load();
    await context.sync();
  });
}

export async function selectaicomment() {
  return Word.run(async (context) => {
    // Get the current selection from the document
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write("Action failed. Error: " + asyncResult.error.message);
      } else {
        write(asyncResult.value);
      }
    });

    // Function that writes to a div with id='message' on the page.
    async function write(message) {
      const aitext = await generateText(message);
      const comment = context.document.getSelection("Hello World").insertComment(aitext);
      comment.load();
      await context.sync();
    }
  });
}

// NEEDS WORK
//Find grammar errors in text, and then leave comments explaining solutions and why they are errors
//>> 2 Button Options: to find errors in whole document, or selected section
//>> Another 2 Button Options: Leave comments like X-error should be Y-correction, this is because _____;; highlight text in yellow
export async function improvegrammar() {
  return Word.run(async (context) => {
    // Get the current selection from the document
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write2("Action failed. Error: " + asyncResult.error.message);
      } else {
        write2(asyncResult.value);
      }
    });

    // Function that writes to a div with id='message' on the page.
    async function write2(error_message) {
      const message = "Explain how the following text could be improved:" + error_message;
      const aitext = await generateText(message);
      const comment = context.document.getSelection("Hello World").insertComment(aitext);
      comment.load();
      await context.sync();
    }
  });
}
