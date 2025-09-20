// api/generate.js
// Vercel Serverless (CommonJS) + docx-templates
// CORS fix für Credentials + echtes DOCX (Binary), kein JSON-Wrapper.

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

function setCors(req, res) {
  const origin = req.headers["origin"] || "*";

  // Wichtig: Bei Credentials KEIN '*', sondern exakten Origin spiegeln
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");

  // Credentials erlauben
  res.setHeader("Access-Control-Allow-Credentials", "true");

  // Preflight-Ankündigungen spiegeln
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHeaders = req.headers["access-control-request-headers"] || "Content-Type,Authorization";

  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", reqHeaders);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-
