/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// ref: https://raw.githubusercontent.com/OfficeDev/office-js-snippets/prod/samples/word/50-document/manage-comments.yaml

// import generateText from "./openai.js";
// import axios from "axios";
// const API_KEY = "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";
// const API_URL = "https://api.openai.com/v1/engines/davinci-codex/completions";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("comment").onclick = comment;
    document.getElementById("selectcomment").onclick = selectaicomment;
    // document.getElementById("aicomment").onclick = aicomment;
  }
});

import { Configuration, OpenAIApi } from "openai";

const configuration = new Configuration({
  apiKey: "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",
});
const openai = new OpenAIApi(configuration);

async function generateText(message) {
  const completion = await openai.createCompletion({
    model: "text-davinci-002",
    //prompt: "why is kimberly so pretty",
    prompt: message,
    temperature: 0.6,
  });
  return completion.data.choices[0].text;
  // return "hello";
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
    // const text = $("#comment-text").val().toString();
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

// export async function aicomment() {
//   // Set a comment on the selected content.
//   return Word.run(async (context) => {
//     const text = await generateText();
//     const comment = context.document.getSelection("Hello World").insertComment(text);

//     // Load object for display in Script Lab console.
//     comment.load();
//     await context.sync();
//   });
// }
