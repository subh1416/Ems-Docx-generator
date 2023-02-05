const express = require("express");
const moment = require("moment");
const nodemailer = require("nodemailer");
const cors = require("cors");
require("dotenv").config();
const app = express();
const port = 3000;
app.use(cors([]));
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
const docx = require("docx");
var sizeOf = require("image-size");
const fs = require("fs");
const mongoose = require("mongoose");
const Logs = require("./logs");
request = require("request");
const {
  Document,
  ImageRun,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Run,
  Table,
  TableRow,
  TableCell,
  WidthType,
} = docx;
mongoose.set("strictQuery", false);
mongoose
  .connect(process.env.mongo_uri, {
    useUnifiedTopology: true,
    useNewUrlParser: true,
  })
  .then(() => console.log("connected to database"))
  .catch((error) => console.log(error));

const transport = {
  host: "smtp.gmail.com",
  secure: true,
  auth: {
    user: process.env.USER,
    pass: process.env.PWD,
  },
};

const transporter = nodemailer.createTransport(transport);

const resetPasswordMail = async (req, res) => {
  console.log(req.body);
  const mailOptions = {
    from: "EMS-KJSIEIT <kjsieit.ems@somaiya.edu>",
    to: `${req.body.email}`,
    subject: `${req.body.subject}`,
    html: `${req.body.body}`,
    replyTo: `${req.body.reply_to}`,
  };
  try {
    const info = await transporter.sendMail(mailOptions);
    console.log(info.response);
    res.status(200).json({ done: "ok" });
  } catch (err) {
    console.log(err);
    res.status(500).json({ error: err });
  }
};

var download = function (urls) {
  var y = 0;
  urls.map((item) =>
    request.head(item.src, function (err, res, body) {
      request(item.src)
        .pipe(fs.createWriteStream(item.name))
        .on("close", () => console.log("done"));
    })
  );
  return urls.length;
};

var split = (data, bool, val) => {
  let ret = [];
  data.split("\n").map((line) => {
    ret.push(new TextRun({ text: line, size: 24 }));
    if (bool) {
      ret.push(
        new Run({
          break: val,
        })
      );
    }
  });
  return ret;
};

var images = (arr) => {
  let ret = [];
  arr.map((item) => {
    if (item.alt !== undefined && item.alt !== null) {
      ret.push(
        new ImageRun({
          data: fs.readFileSync(`./${item.name}`),
          transformation: {
            width: 548.5714285714286,
            height: 308.57142857142856,
          },
        })
      );
      ret.push(new TextRun({ break: 1, text: item.alt, size: 24 }));
      ret.push(
        new Run({
          break: 2,
        })
      );
    }
  });
  return ret;
};

app.post("/mail", resetPasswordMail);

app.post("/", async (req, res) => {
  const body = req.body;
  var imgs = [];
  var banner = body.data.banner.replace("../../../report/", "");
  if (banner !== "") {
    imgs.push({ src: `https://ems.kjsieit.in/report/${banner}`, name: banner });
  }
  var glimpse = body.data.content.glimpse.data.map((item) => {
    var name = item.src.replace("../../../report/", "");
    if (name !== "") {
      imgs.push({
        src: `https://ems.kjsieit.in/report/${name}`,
        name,
        alt: item.alt,
      });
    }
    return {
      src: name,
      alt: item.alt,
    };
  });
  res.send("sucess");
  download(imgs);
  await new Promise((r) => setTimeout(r, 8000));
  imgs.map(({ name }, index) => {
    var dimensions = sizeOf(name);
    imgs[index].dimensions = dimensions;
  });
  console.log("Going");
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 500,
            },
          },
        },
        children: [
          new Paragraph({
            children: [
              new ImageRun({
                data: fs.readFileSync("./letterhead1.png"),
                transformation: {
                  width: 600,
                  height: 100,
                },
              }),
              new Run({
                break: 1,
              }),
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `Report of ${body?.data?.type} on`,
                bold: true,
                underline: true,
                size: 26,
              }),
              new Run({
                break: 1,
              }),
            ],

            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `${body?.data?.title}`,
                bold: true,
                underline: true,
                size: 26,
              }),
              new Run({
                break: 1,
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
          banner !== "" &&
            new Paragraph({
              children: [
                new ImageRun({
                  data: fs.readFileSync(`./${banner}`),
                  transformation: {
                    width: 488,
                    height: 469.3333333333333,
                  },
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
          new Paragraph({
            children: [],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `The ${
                  body?.data?.org
                } of K. J. Somaiya Institute of Engineering and Information Technology (KJSIEIT) organized a ${
                  body?.data?.type
                } on "${body?.data?.title}" on ${moment(
                  body?.data?.date
                ).format("MMMM DD YYYY")} at ${body?.data?.time}`,
                size: 22,
              }),
            ],
            alignment: AlignmentType.JUSTIFIED,
          }),
          new Paragraph({
            children: [],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `${body.data.content.objective.label}`,
                bold: true,
                size: 24,
              }),
              new Run({
                break: 1,
              }),
              ...split(body.data.content.objective.data, true, 1),
              ,
            ],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `${body.data.content.ip.label}: `,
                bold: true,
                size: 24,
              }),
              new TextRun({
                text: `${body.data.content.ip.data}`,
                size: 24,
              }),
              new Run({
                break: 2,
              }),
              new TextRun({
                text: `${body.data.content.ep.label}: `,
                bold: true,
                size: 24,
              }),
              new TextRun({
                text: `${body.data.content.ep.data}`,
                size: 24,
              }),
              new Run({
                break: 2,
              }),
              new TextRun({
                text: `${body.data.content.venue.label}: `,
                bold: true,
                size: 24,
              }),
              new TextRun({
                text: `${body.data.content.venue.data}`,
                size: 24,
              }),
              new Run({
                break: 2,
              }),
              new TextRun({
                text: `${body.data.content.rp.label}: `,
                bold: true,
                size: 24,
              }),
            ],
          }),
          ...[...split(body.data.content.rp.data, false, 0)].map(
            (item) =>
              new Paragraph({
                children: [item],
                alignment: AlignmentType.JUSTIFIED,
              })
          ),
          new Paragraph({
            children: [
              new Run({
                break: 2,
              }),
              new TextRun({
                text: `${body.data.content.kp.label}: `,
                bold: true,
                size: 24,
              }),
            ],
          }),
          ...[...split(body.data.content.kp.data, false, 0)].map(
            (item) =>
              new Paragraph({
                children: [item],
                alignment: AlignmentType.JUSTIFIED,
              })
          ),
          new Paragraph({
            children: [
              new Run({
                break: 1,
              }),
              new TextRun({
                text: `${body.data.content.outcomes.label}: `,
                bold: true,
                size: 24,
              }),
              new Run({
                break: 1,
              }),
              ...split(body.data.content.outcomes.data, true, 1),
              new Run({
                break: 2,
              }),
              new TextRun({
                text: `${body.data.content.pos.label}: `,
                bold: true,
                size: 24,
              }),
              new Run({
                break: 1,
              }),
            ],
          }),

          new Table({
            rows: [
              ...body.data.content.pos.data.data.map(
                (it) =>
                  new TableRow({
                    children: [
                      ...it.map(
                        (item) =>
                          new TableCell({
                            width: {
                              size: 100 / it.length,
                              type: WidthType.PERCENTAGE,
                            },
                            height: {
                              size: 100,
                            },
                            children: [new Paragraph({ text: item })],
                          })
                      ),
                    ],
                  })
              ),
            ],
          }),

          new Paragraph({
            children: [
              new Run({
                break: 1,
              }),
              new TextRun({
                text: `${body.data.content.ec.label}: `,
                bold: true,
                size: 24,
              }),
              new Run({
                break: 1,
              }),
              new TextRun({
                text: `${body.data.content.ec.data}`,
                size: 24,
              }),
              new Run({
                break: 1,
              }),
              new TextRun({
                text: `${body.data.content.glimpse.label}`,
                bold: true,
                size: 24,
                break: 2,
              }),
            ],
          }),
          new Paragraph({
            children: [...images(imgs)],
            alignment: AlignmentType.CENTER,
          }),
        ],
      },
    ],
  });
  const file =
    body.data.title.replace("/", "-").trim() +
    "_" +
    body.data.org.trim() +
    ".docx";

  fs.exists(`./documents/${file}`, function (exists) {
    if (exists) {
      //Show in green
      console.log("eeee");
      fs.unlink(`./documents/${file}`, (err) => {
        console.log(err);
      });
    } else {
      console.log(exists);
    }
  });

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(`./documents/${file}`, buffer);
  });
  imgs.map((item) => {
    fs.unlink(`./${item.name}`, () => {
      console.log("Deleted");
    });
  });
});

// DB save
app.post("/savelog", async (req, res) => {
  try {
    let log = await Logs.create({
      useremail: req.body.useremail,
      ip: req.body.ip,
      uri:
        req.body.uri.lastIndexOf("/") != -1
          ? req.body.uri
              .substring(req.body.uri.lastIndexOf("/") + 1, req.body.uri.length)
              .replaceAll(".php", " ")
              .trim()
          : req.body.uri.replaceAll(".php", " ").trim(),
      urioriginal: req.body.uri,
      agent: req.body.agent,
      referer: req.body.referer,
    });
    res.status(200).json({ message: "OK", result: log });
  } catch (err) {
    res.status(422).send(err);
    console.log(err);
  }
});

app.get("/getlogs", async (_, res) => {
  try {
    let logs = await Logs.find(
      {},
      { ip: 0, agent: 0, referer: 0, urioriginal: 0 }
    );
    res.status(200).json({ message: "OK", result: logs });
  } catch (err) {
    res.status(422).send(err);
    console.log(err);
  }
});

app.listen(port, () => {
  console.log(`Example app listening on port ${port}`);
});
