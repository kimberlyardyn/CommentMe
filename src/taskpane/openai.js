import axios from "axios";

// const API_KEY = "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";
// const API_URL = "https://api.openai.com/v1/engines/davinci-codex/completions";

export default async function generateText() {
  return "hello";
//   try {
//     const response = await axios.post(
//       API_URL,
//       {
//         prompt: "Hello World",
//         temperature: 0.5,
//         max_tokens: 100,
//       },
//       {
//         headers: {
//           "Content-Type": "application/json",
//           Authorization: `Bearer ${API_KEY}`,
//         },
//       }
//     );
//     //return response.data.choices[0].text;
//     return "What is the meaning";
//   } catch (error) {
//     //console.error(error);
//   }
};

//console.log(generateText("What is the meaning of life?"));
