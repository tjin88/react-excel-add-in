/* eslint-disable prettier/prettier */
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
  questionaire_v2,
  deleteSheet,
  report,
} from "./helpers/functions";
import { initializeOffice } from "./helpers/eventHandlers";
// import TextInsertion from "./components/TextInsertion";
import TaskPane from "./components/TaskPaneChild";
import "./App.css";

const App = () => {
  useEffect(() => {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        initializeOffice();
      }
    });
  }, []);

  useEffect(() => {
    // some API call to get the questions and slugs from FlowPoint
    // setFields(data.fields);
  }, []);

  /**
   * Old model (v1)
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

  /**
   * Currently updating this (v2)
   * 
   * Format: key = slug, value = [Question num, Question, Answer, Validation]
   * Question num: Number. Can have one or more follow-up questions
   * Question: Just a string
   * Answer: Flowpoint will either provide a pre-filled answer OR an empty string
   * Validation: JSON string containing FlowPoint's DSL
   */
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [flowpointQuestions, setFlowpointQuestions] = useState({
    "slug1": ["1", "How much exposure do you have to the US and Canada (%)", "", "Some DSL Here"],
    "slug2": ["2", "Have you had any board meetings in the past month (Yes/No)", "", "Some DSL Here"],
    "slug3": ["2.1", "What was the purpose of the board meeting?", "", "Some DSL Here"],
    "slug4": ["3", "Are you following all legal obligations (Yes/No)", "", "Some DSL Here"],
    "slug5": ["3.1", "Why not?", "", "Some DSL Here"],
    "slug6": ["4", "What was your growth in your MRR (%)", "", "Some DSL Here"],
    "slug7": ["5", "How much exposure do you have to the US and Canada (%)", "50", "Some DSL Here"],
    "slug8": ["6", "Have you had any board meetings in the past month (Yes/No)", "Yes", "Some DSL Here"],
    "slug9": ["6.1", "What was the purpose of the board meeting?", "Some reason", "Some DSL Here"],
    "slug10": ["7", "Are you following all legal obligations (Yes/No)", "Yes", "Some DSL Here"],
    "slug11": ["7.1", "Why not?", "N/A", "Some DSL Here"],
    "slug12": ["8", "What was your growth in your MRR (%)", "12345", "Some DSL Here"],
  });

  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [fields, setFields] = useState([
    ["Net Sales", "$30000"],
    ["Total Cost of Goods Sold", "$5000"],
    ["Gross Profit", "$25000"],
    ["Total Operating Expenses", "$10000"],
    ["Operating Profit (Loss)", "$15000"],
    ["Interest Income", "-"],
    ["Other Income", "$1000"],
    ["Profit (Loss) Before Taxes", "$16000"],
    ["Income Tax Expense", "$4800"],
    ["Net Profit (Loss)", "$11200"],
    ["How much exposure do you have to the US and Canada (%)", "25%"],
    ["Have you had any board meetings in the past month (Yes/No)", "Yes"],
    ["What was the purpose of the board meeting?", "Reason"],
    ["Are you following all legal obligations (Yes/No)", "Yes"],
    ["Why not?", "-"],
    ["What was your growth in your MRR (%)", "0%"],
  ]);

  return (
    <div className="ms-Fabric ms-font-m ms-welcome">
      <header className="ms-welcome__header ms-bgColor-neutralLighter">
        <img height="50" src="../../assets/FlowPoint-Logo.svg" alt="FlowPoint" title="FlowPoint" />
      </header>
      <main id="app-body" className="ms-welcome__main">
        {/* <Button id="backgroundWhite" label="Background White" onClick={backgroundWhite} /> */}
        <TaskPane />
        <Button id="clearMessage" label="Clear Message" onClick={clearMessage} />
        <Button id="validate" label="Validate" onClick={validate} />
        {/* <Button id="hideRows" label="Hide Rows" onClick={hideRows} /> */}
        <Button id="unhideRows" label="Unhide Rows" onClick={unhideRows} />
        <Button id="questionaire" label="Questionaire" onClick={() => questionaire(questions)} />
        <Button id="deleteQuestionaire" label="Delete Questionaire Sheet" onClick={() => deleteSheet("Questionaire")} />
        <Button id="questionaire" label="Questionaire v2" onClick={() => questionaire_v2(questions)} />
        <Button id="deleteQuestionaire" label="Delete Questionaire v2 Sheet" onClick={() => deleteSheet("Questionaire_v2")} />
        <Button id="report" label="Report" onClick={() => report(fields)} />
        <Button id="deleteReport" label="Delete Report Sheet" onClick={() => deleteSheet("Report")} />
        <Message />
        {/* <TextInsertion /> */}
      </main>
    </div>
  );
};

export default App;
