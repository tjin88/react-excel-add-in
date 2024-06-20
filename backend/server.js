const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const https = require("https");
const fs = require("fs");
const path = require("path");
const { OpenAI } = require("openai");
require("dotenv").config();

// Path to SSL certificate and key files
const keyPath = path.join(__dirname, "..", "server.key");
const certPath = path.join(__dirname, "..", "server.crt");

// Read the SSL certificate and key
let privateKey;
let certificate;

try {
  privateKey = fs.readFileSync(keyPath, "utf8");
  certificate = fs.readFileSync(certPath, "utf8");
} catch (error) {
  console.error("Error reading SSL certificate and key files:", error.message);
  process.exit(1); // Exit process with an error code
}

const credentials = { key: privateKey, cert: certificate };

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY_EXCEL_ADD_IN });

const app = express();
app.use(cors());
app.use(bodyParser.json());

const port = process.env.PORT || 3001;

app.get("/", (req, res) => {
  res.send("Server up and running!");
});

// Make the API call to your custom assistant
async function callCustomAssistant(prompt) {
  try {
    // set up the Assistant - assistants are persistent, only create once!
    let assistant = null;

    // retrieve or create the assistant
    let assistants = await openai.beta.assistants.list();
    assistant = assistants.data.find(assistant => assistant.name == process.env.OPENAI_API_ASSISTANT_NAME);

    // If there is no assistant, create one with the specified name and instructions
    // And add the desired files to the vector store
    if (assistant == null) {
      assistant = await openai.beta.assistants.create({
        name: process.env.OPENAI_API_ASSISTANT_NAME,
        instructions: process.env.OPENAI_API_ASSISTANT_INSTRUCTIONS,
        model: "gpt-4o",
        tools: [{ type: "file_search" }],
      });

      const r_units = await openai.files.create({
        file: fs.createReadStream("files/om_r_units.pdf"),
        purpose: "assistants",
      });

      const afg_units = await openai.files.create({
        file: fs.createReadStream("files/om_afg_units.pdf"),
        purpose: "assistants",
      });

      // Create a vector store including our two files.
      let vectorStore = await openai.beta.vectorStores.create({
        name: "Offering Memorandums",
        file_ids: [r_units.id, afg_units.id],
      });

      console.log("Created vector store:", vectorStore.id);

      // Ensure files are passed correctly
      // await openai.beta.vectorStores.fileBatches.uploadAndPoll(vectorStore.id, fileStreams);

      await openai.beta.assistants.update(assistant.id, {
        tool_resources: { file_search: { vector_store_ids: [vectorStore.id] } },
      });
    }

    console.log("Using the following assistant:", assistant.id, assistant.name, assistant.model, assistant.tools);

    const thread = await openai.beta.threads.create({
      messages: [
        {
          role: "user",
          content: prompt,
        },
      ],
    });

    console.log("Thread: ", thread);

    const run = await openai.beta.threads.runs.createAndPoll(thread.id, {
      assistant_id: assistant.id,
    });

    const messages = await openai.beta.threads.messages.list(thread.id, {
      run_id: run.id,
    });

    const message = messages.data.pop();
    if (message.content[0].type === "text") {
      const { text } = message.content[0];
      const { annotations } = text;
      let citations = [];
      let index = 0;

      // TODO: Adjust later (if we want or don't want citations)
      for (let annotation of annotations) {
        text.value = text.value.replace(annotation.text, "[" + index + "]");
        const { file_citation } = annotation;
        if (file_citation) {
          console.log("File citation:", file_citation)
          const citedFile = await openai.files.retrieve(file_citation.file_id);
          console.log("Cited file:", citedFile.filename);
          citations.push("[" + index + "]" + citedFile.filename);
        }
        index++;
      }

      console.log("Text.value: ", text.value);
      console.log("Citations: ", citations.join("\n"));

      const full_response = text.value + "\n\n" + citations.join("\n");

      return { text: full_response };
    }
  } catch (error) {
    console.error("Error in callCustomAssistant:", error);
    throw error;
  }
}

app.post("/gpt-api", async (req, res) => {
  console.log("Received request to /gpt-api");
  const { prompt } = req.body;

  console.log("Prompt:", prompt);

  try {
    const answer = await callCustomAssistant(prompt);
    res.json(answer);
  } catch (error) {
    console.error("Error occurred while calling OpenAI API:", error.response ? error.response.data : error.message);
    res.status(500).send(error.toString());
  }
});

// Create HTTPS server
const httpsServer = https.createServer(credentials, app);

httpsServer.listen(port, () => {
  console.log(`Server is running on https://localhost:${port}`);
});
