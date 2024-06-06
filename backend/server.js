const express = require("express");
const axios = require("axios");
const cors = require("cors");
const bodyParser = require("body-parser");
const https = require("https");
const fs = require("fs");
const path = require("path");
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

const app = express();
app.use(cors());
app.use(bodyParser.json());

const port = process.env.PORT || 3001;

app.get("/", (req, res) => {
  res.send("Server up and running!");
});

app.post("/gpt-api", async (req, res) => {
  /*
    Example messages object:
      "messages": [
          { "role": "user", "content": "Can you explain the difference between revenue and profit?" },
          { "role": "assistant", "content": "Revenue is the total amount of money generated by the sale of goods or services, while profit is the amount of money that remains after all expenses are subtracted from revenue." },
          { "role": "user", "content": "Can you give an example?" }
      ]
  */

  console.log("Received request to /gpt-api");
  const { messages } = req.body;

  try {
    const response = await axios.post(
      "https://api.openai.com/v1/chat/completions",
      {
        model: "gpt-4",
        messages: [
          {
            role: "system",
            content:
              "You are a helpful assistant knowledgeable in financial accounting. Please only respond with numbers and very short sentences (less than 5 words).",
          },
          ...messages,
        ],
        max_tokens: 150,
      },
      {
        headers: {
          Authorization: `Bearer ${process.env.OPENAI_API_KEY_EXCEL_ADD_IN}`,
          "Content-Type": "application/json",
        },
      }
    );
    res.json(response.data);
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
