/* eslint-disable no-undef */
import React, { useState } from "react";
import axios from "axios";
import insertText from "../office-document";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import "./TaskPaneChild.css";

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
  subtitle: {
    margin: 0,
  },
});

const TaskPane = () => {
  const [prompt, setPrompt] = useState("");
  const [response, setResponse] = useState("");
  const [messages, setMessages] = useState([]);
  const styles = useStyles();

  const callGPT = async () => {
    try {
      console.log("Calling GPT with prompt: " + prompt);

      const updatedMessages = [...messages, { role: "user", content: prompt }];

      const res = await axios.post("https://localhost:3001/gpt-api", {
        messages: updatedMessages,
      });

      const assistantResponse = res.data.choices[0].message.content;

      setResponse(assistantResponse);

      // Insert the assistant's response into the document (currently set to cell A1)
      await insertText(assistantResponse, "A1");

      setMessages([...updatedMessages, { role: "assistant", content: assistantResponse }]);
    } catch (error) {
      console.error("Error occurred while calling the server:", error);
      setResponse("Error: " + error.toString());
    }
  };

  return (
    <div className={styles.textPromptAndInsertion}>
      {/* <div className="gpt-integration"> */}
      <h1>GPT Integration</h1>
      <h3>Enter a prompt below and click the button to send it to GPT-4:</h3>
      <textarea
        className={styles.textAreaField}
        // className="prompt-textarea"
        rows={4}
        cols={50}
        value={prompt}
        onChange={(e) => setPrompt(e.target.value)}
      />
      <br />
      <button onClick={callGPT}>Send to GPT</button>
      <div>{response}</div>
    </div>
  );
};

export default TaskPane;
