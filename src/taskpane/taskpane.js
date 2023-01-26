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
  }
});

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
  // export async function callGPT3() {
  //   // send to gpt 3 the paragraph

  //   //get the result and append to the comment section
  // }
}

// export async function comment(text) {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     text = context.document.body.Comments.Add(this.body.Paragraphs[1].Range, "Hello World");

//     await context.sync(text);
//   });
// }

//object text = "Add a comment to the first paragraph.";
//this.Application.ActiveDocument.Comments.Add(
// this.Application.ActiveDocument.Paragraphs[1].Range, ref text)

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
