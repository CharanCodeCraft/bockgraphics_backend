const express = require("express");
const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const bodyParser = require("body-parser");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(bodyParser.json());

const COMPANY_CONFIG = {
  Automotive: {
    template: "Automotive.docx",
    outputName: "Automotive_Offer_Letter.docx",
  },
  Foods: {
    template: "Foods.docx",
    outputName: "Foods_Offer_Letter.docx",
  },
  Space: {
    template: "Space.docx",
    outputName: "Space_Offer_Letter.docx",
  },
  Ai: {
    template: "Ai.docx",
    outputName: "AI_Offer_Letter.docx",
  },
  Health: {
    template: "health.docx",
    outputName: "Health_Offer_Letter.docx",
  },
  Chain: {
    template: "Chain.docx",
    outputName: "Chain_Offer_Letter.docx",
  },
  Institutes: {
    template: "Institutes.docx",
    outputName: "Institutes_Offer_Letter.docx",
  },
  Force: {
    template: "Force.docx",
    outputName: "Force_Offer_Letter.docx",
  },
  Parent: {
    template: "Bock.docx",
    outputName: "Bock_Offer_Letter.docx",
  },
};


app.post("/generate-offer", (req, res) => {
  try {
    const {
      name,
      program,
      issueDate,
      startDate,
      endDate,
      months,
      mode,
      field, // e.g. "Automotive", "AI", "Bock"
    } = req.body;

    const config = COMPANY_CONFIG[field];

    if (!config) {
      return res.status(400).json({
        error: "Invalid company selected",
      });
    }

    const templatePath = path.join(
      __dirname,
      "templates/offer_letter",
      config.template
    );

    const content = fs.readFileSync(templatePath, "binary");

    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{", end: "}" },
    });

    doc.render({
      XLAMAX: name,
      XVARIXBLEX: program,
      XDATEX: issueDate,
      XOX: months,
      XNOXNOXXXXOXX: startDate,
      YNOXNOXXXXOXY: endDate,
      XMODEX: mode,
    });

    const buffer = doc.getZip().generate({
      type: "nodebuffer",
      compression: "DEFLATE",
    });

    res.set({
      "Content-Disposition": `attachment; filename=${config.outputName}`,
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({
      error: "Offer letter generation failed",
      details: err.message,
    });
  }
});


app.listen(3000, () =>
  console.log("Server running on http://localhost:3000")
);
