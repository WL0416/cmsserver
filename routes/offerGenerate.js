const express = require("express");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const compressing = require("compressing");

const router = express.Router();

router.use(cors());

// The error object contains additional information when logged with JSON.stringify (it contains a properties object containing all suberrors).
function replaceErrors(key, value) {
  if (value instanceof Error) {
    return Object.getOwnPropertyNames(value).reduce(function (error, key) {
      error[key] = value[key];
      return error;
    }, {});
  }
  return value;
}

function errorHandler(error) {
  console.log(JSON.stringify({ error: error }, replaceErrors));

  if (error.properties && error.properties.errors instanceof Array) {
    const errorMessages = error.properties.errors
      .map(function (error) {
        return error.properties.explanation;
      })
      .join("\n");
    console.log("errorMessages", errorMessages);
    // errorMessages is a humanly readable message looking like this :
    // 'The tag beginning with "foobar" is unopened'
  }
  throw error;
}

const convertDate = (date) => {
  const timeArray = date.split("-");
  return timeArray[2] + "/" + timeArray[1] + "/" + timeArray[0];
};

const deleteFile = (file) => {
  fs.unlink(file, (err) => {
    if (err) {
      console.error(err.toString());
    } else {
      console.warn(file + " deleted");
    }
  });
};

router.get("/:passport.:code", (req, res) => {
  const { passport, code } = req.params;
  const file = path.resolve(__dirname, `../templates/${passport}@${code}.docx`);
  res.send(file);
});

router.post("/", (req, res) => {
  let data = req.body;
  const birthday = convertDate(data.birthday);

  data = {
    ...data,
    birthday,
  };

  //Load the docx file as a binary
  const content = fs.readFileSync(
    path.resolve(__dirname, "../templates/input1.docx"),
    "binary"
  );

  const zip = new PizZip(content);
  let doc;

  try {
    doc = new Docxtemplater(zip);
  } catch (error) {
    // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
    errorHandler(error);
  }

  //set the templateVariables
  doc.setData(data);

  try {
    // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
    doc.render();
  } catch (error) {
    // Catch rendering errors (errors relating to the rendering of the template : angularParser throws an error)
    errorHandler(error);
  }

  var buf = doc.getZip().generate({ type: "nodebuffer" });

  const file = path.resolve(
    __dirname,
    `../templates/${data.firstName}${data.lastName}@${data.code}.docx`
  );

  // buf is a nodejs buffer, you can either write it to a file or do anything else with it.
  fs.writeFileSync(file, buf);

  setTimeout(() => deleteFile(file), 60000);

  res.send(file);
});

module.exports = router;
