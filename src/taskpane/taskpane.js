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
    document.getElementById("grammar").onclick = correctgrammar;
    document.getElementById("mySelect").onchange = givenprompt;
    document.getElementById("sendCreatedPrompt").onsubmit = createprompt;
    document.getElementById("myInput").onchange = createprompt;
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
    temperature: 0.7,
    max_tokens: 70,
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

// COULD USE WORK
//Find grammar errors in text, and then leave comments explaining solutions and why they are errors
//>> 2 Button Options: to find errors in whole document, or selected section
//>> highlight errors in yellow -- or, users just use grammarly to pinpoint errors?
//temperature: 0.7, max_tokens: 70 seems to work well
export async function correctgrammar() {
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
      const message =
        "Explain all the grammatical errors in this sentence and provide an example of the sentence grammatically corrected:" +
        error_message;
      const aitext = await generateText(message);
      const comment = context.document.getSelection("Hello World").insertComment(aitext);
      comment.load();
      await context.sync();
    }
  });
}

//for the prompt creation
/* When the user clicks on the button,
toggle between hiding and showing the dropdown content */

//NEEDS WORK
// Button Options: Write/Paste own prompt into text box
export async function givenprompt() {
  // Set a comment on the selected content, produces selected_text
  return Word.run(async (context) => {
    // Get the current selection from the document
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write4("Action failed. Error: " + asyncResult.error.message);
      } else {
        const selected_text = asyncResult.value;
        write4(selected_text);
      }
    });

    // Function that writes to a div with id='message' on the page.
    async function write4(selected_text) {
      const prompt = document.getElementById("mySelect").value;
      const aitext = await generateText(prompt + selected_text);
      const comment = context.document.getSelection(selected_text).insertComment(aitext);
      comment.load();
      await context.sync();
    }
  });
}

export async function createprompt() {
  // Set a comment on the selected content, produces selected_text
  return Word.run(async (context) => {
    // Get the current selection from the document
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write3("Action failed. Error: " + asyncResult.error.message);
      } else {
        const selected_text = asyncResult.value;
        write3(selected_text);
      }
    });

    // Function that writes to a div with id='message' on the page.
    async function write3(selected_text) {
      const prompt = document.getElementById("myInput").value;
      const aitext = await generateText(prompt + selected_text);
      const comment = context.document.getSelection(selected_text).insertComment(aitext);
      comment.load();
      await context.sync();
    }
  });
}
