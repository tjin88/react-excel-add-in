// eslint-disable-next-line no-redeclare
/* global Office */

import React, { useState, useEffect } from "react";
import Button from "./components/Button";
import Message from "./components/Message";
import {
  // backgroundWhite,
  validate,
  // hideRows,
  unhideRows,
  clearMessage,
  questionaire,
  deleteQuestionaireSheet,
} from "./helpers/functions";
import { initializeOffice } from "./helpers/eventHandlers";
// import TextInsertion from "./components/TextInsertion";
import "./App.css";

const App = () => {
  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        initializeOffice();
      }
    });
  }, []);

  /**
   * Will be changed to an API call to get the questions and slugs from FlowPoint.
   * Until then, the setQuestions funciton willl be unused
   *
   * Format: [Order, Question, Method, Answer, Hidden/Visible, Slug]
   * Order: Number. Can have one or more follow-up questions
   * Question: String. Will need to get through an API call from FlowPoint
   * Method: ["Num", "Num & capped", "String", "Bool & Hide No", "Bool & Hide Yes"]
   * Answer: THIS WILL ALWAYS BE EMPTY. It's added to the array to ensure the table border is created successfully
   * Hidden/Visible: ["Hidden", "Visible"]. Think of it as a bool
   * Slug: String (assumed). Will need to get through an API call from FlowPoint
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [questions, setQuestions] = useState([
    ["1", "How much exposure do you have to the US and Canada (%)", "Num & capped", "", "", "exp-us-canada"],
    ["2", "Have you had any board meetings in the past month (Yes/No)", "Bool & Hide No", "", "", "board"],
    ["2.1", "What was the purpose of the board meeting?", "String", "", "Hidden", "board-reason"],
    ["3", "Are you following all legal obligations (Yes/No)", "Bool & Hide Yes", "", "", "legal-obl"],
    ["3.1", "Why not?", "String", "", "Hidden", "legal-obl-reason"],
    ["4", "What was your growth in your MRR (%)", "Num", "", "", "mrr-growth"],
    ["5", "How much exposure do you have to the US and Canada (%)", "Num & capped", "", "", "exp-us-canada"],
    ["6", "Have you had any board meetings in the past month (Yes/No)", "Bool & Hide No", "", "", "board"],
    ["6.1", "What was the purpose of the board meeting?", "String", "", "Hidden", "board-reason"],
    ["7", "Are you following all legal obligations (Yes/No)", "Bool & Hide Yes", "", "", "legal-obl"],
    ["7.1", "Why not?", "String", "", "Hidden", "legal-obl-reason"],
    ["8", "What was your growth in your MRR (%)", "Num", "", "", "mrr-growth"],
  ]);

  return (
    <div className="ms-Fabric ms-font-m ms-welcome">
      <header className="ms-welcome__header ms-bgColor-neutralLighter">
        <img height="50" src="../../assets/FlowPoint-Logo.svg" alt="FlowPoint" title="FlowPoint" />
      </header>
      <main id="app-body" className="ms-welcome__main">
        {/* <Button id="backgroundWhite" label="Background White" onClick={backgroundWhite} /> */}
        <Button id="clearMessage" label="Clear Message" onClick={clearMessage} />
        <Button id="validate" label="Validate" onClick={validate} />
        {/* <Button id="hideRows" label="Hide Rows" onClick={hideRows} /> */}
        <Button id="unhideRows" label="Unhide Rows" onClick={unhideRows} />
        <Button id="questionaire" label="Questionaire" onClick={() => questionaire(questions)} />
        <Button id="deleteQuestionaire" label="Delete Questionaire Sheet" onClick={deleteQuestionaireSheet} />
        <Message />
        {/* <TextInsertion /> */}
      </main>
    </div>
  );
};

export default App;
