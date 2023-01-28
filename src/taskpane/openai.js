import axios from "axios";

export default async function generateText() {
  return "hello";
  try {
    const response = await axios.post(
      API_URL,
      {
        prompt: "Hello World",
        temperature: 0.5,
        max_tokens: 100,
      },
      {
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${API_KEY}`,
        },
      }
    );
    //return response.data.choices[0].text;
    return "What is the meaning";
  } catch (error) {
    //console.error(error);
  }
};

//console.log(generateText("What is the meaning of life?"));
