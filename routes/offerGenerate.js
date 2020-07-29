const express = require("express");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const { Storage } = require("@google-cloud/storage");
const storage = new Storage();
const srcBucket = "csmserver.appspot.com";

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

const deleteRemotefile = async (filename) => {
  // Deletes the file from the bucket
  await storage.bucket(srcBucket).file(`outputs/${filename}`).delete();

  console.log(`gs://${srcBucket}/outputs/${filename} deleted.`);
};

async function listBuckets() {
  try {
    const results = await storage.getBuckets();

    const [buckets] = results;

    console.log("Buckets:");
    buckets.forEach((bucket) => {
      console.log(bucket.name);
    });
  } catch (err) {
    console.error("ERROR:", err);
  }
}

router.get("/", (req, res) => {
  const file = path.resolve(__dirname, "../templates/input1.docx");
  const remoteFile = storage.bucket(srcBucket).file("outputs/output.docx");
  fs.createReadStream(file)
    .pipe(remoteFile.createWriteStream())
    .on("error", (err) => {
      console.log(err);
    })
    .on("finish", () => {
      console.log("finish!");
    });
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

  const buf = doc.getZip().generate({ type: "nodebuffer" });

  const fileName = `${data.firstName}${data.lastName}@${data.code}.docx`;

  const remoteFile = storage.bucket(srcBucket).file(`outputs/${fileName}`);

  fs.writeFileSync(`/tmp/${fileName}`, buf);

  fs.createReadStream(`/tmp/${fileName}`)
    .pipe(remoteFile.createWriteStream())
    .on("error", (err) => {
      console.log(err);
      res.send(err);
    })
    .on("finish", () => {
      console.log("finish!");
      res.send(
        `https://storage.googleapis.com/${srcBucket}/outputs/${fileName}`
      );
      setTimeout(() => {
        deleteFile(`/tmp/${fileName}`);
        deleteRemotefile(fileName);
      }, 30000);
    });
});

module.exports = router;
