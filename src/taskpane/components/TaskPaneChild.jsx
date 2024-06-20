/* eslint-disable no-undef */
import React, { useState } from "react";
import axios from "axios";
// import insertText from "../office-document";
import { tokens, makeStyles } from "@fluentui/react-components";

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

  const callTextboxGPT = async () => {
    try {
      let assistantResponse;

      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["address", "values"]);
        await context.sync();

        document.getElementById("message").innerText += `Calling GPT with prompt: ${prompt}.\n`;

        const updatedMessages = [...messages, { role: "user", content: prompt }];
        const res = await axios.post("https://localhost:3001/gpt-api", {
          // messages: updatedMessages,
          prompt: prompt,
        });
        assistantResponse = res.data.text;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.protection.unprotect();
        const rangeToUpdate = sheet.getRange(range.address).getCell(0, 0).getOffsetRange(0, 1);
        rangeToUpdate.values = [[assistantResponse]];
        rangeToUpdate.format.font.color = "black";
        rangeToUpdate.format.autofitColumns();

        setResponse(assistantResponse);
        setMessages([...updatedMessages, { role: "assistant", content: assistantResponse }]);

        sheet.protection.protect();
        await context.sync();
      });
    } catch (error) {
      console.error("Error occurred while calling the server:", error);
      setResponse("Error: " + error.toString());
    }
  };

  const callActiveCellGPT = async () => {
    try {
      let assistantResponse;

      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load(["address", "values"]);
        await context.sync();

        const prompt = range.values[0][0];

        document.getElementById("message").innerText += `Calling GPT with prompt: ${prompt}.\n`;

        const updatedMessages = [...messages, { role: "user", content: prompt }];
        const res = await axios.post("https://localhost:3001/gpt-api", {
          // messages: updatedMessages,
          prompt: prompt,
        });
        // assistantResponse = res.data.choices[0].message.content;
        assistantResponse = res.data.text;

        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.protection.unprotect();
        const rangeToUpdate = sheet.getRange(range.address).getCell(0, 0).getOffsetRange(0, 1);
        rangeToUpdate.values = [[assistantResponse]];
        rangeToUpdate.format.font.color = "black";
        rangeToUpdate.format.autofitColumns();

        setResponse(assistantResponse);
        setMessages([...updatedMessages, { role: "assistant", content: assistantResponse }]);

        sheet.protection.protect();
        await context.sync();
      });
    } catch (error) {
      console.error("Error occurred while calling the server:", error);
      setResponse("Error: " + error.toString());
    }
  };

  return (
    <div className={styles.textPromptAndInsertion}>
      <h1>GPT Integration</h1>
      <h3>Enter a prompt below and click the button to send it to GPT-4:</h3>
      <textarea
        className={styles.textAreaField}
        rows={4}
        cols={50}
        value={prompt}
        onChange={(e) => setPrompt(e.target.value)}
      />
      <br />
      <div>
        <button onClick={callTextboxGPT}>Send to prompt from text box to GPT</button>
        <button onClick={callActiveCellGPT}>Send to prompt from active cell to GPT</button>
      </div>
      <div>{response}</div>
    </div>
  );
};

export default TaskPane;
