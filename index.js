const fs = require("fs");
const docx = require("docx");
const { Paragraph, ExternalHyperlink, Packer, TextRun } = require("docx");
const he = require("he");

async function main() {
  const args = process.argv.slice(2);
  const path = args.slice(0, 1)[0];
  await exportHighlights({ path });
}

const exportHighlights = async ({ path }) => {
  fs.readFile(path, "utf8", (err, data) => {
    if (err) {
      console.error(err);
      return;
    }

    const filtered = filterTrash(data);
    writeToDocx(filtered);

    console.info("Done!");
  });
};

/**
 * @param {string} data
 * @returns {[{note: '', link: ''}]} array of objects
 */
const filterTrash = (data) => {
  const regexp =
    /<td\s+class="\w+"\s+colspan="\d+"\s+rowspan="\d+"><p\s+class="\w+"><span\s+class="\w+">([\w\W][^<>]+?)<\/span><\/p><p\s+class="[\w\s]+"><span\s+class="\w+"><\/span><\/p><p\s+class="\w+"><span\s+class="\w+">.*?<\/span><\/p><\/td><td.*?href="(.*?)"/gm;
  let m;
  let returnVal = [];

  const matches = data.matchAll(regexp);

  for (const match of matches)
    returnVal.push({ note: he.decode(match[1]), link: match[2] });

  return returnVal;
};

const writeToDocx = (obj) => {
  const children = [];

  for (const note of obj) {
    const paragraph = new Paragraph({
      text: note.note,
    });
    const link = new Paragraph({
      children: [
        new ExternalHyperlink({
          children: [
            new TextRun({
              text: "Link",
              style: "Hyperlink",
            }),
          ],
          link: note.link,
        }),
      ],
    });

    children.push(
      paragraph,
      link,
      new Paragraph({ text: "" }),
      new Paragraph({ text: "" })
    );
  }

  const doc = new docx.Document({
    sections: [
      {
        children,
      },
    ],
  });

  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("highlights.docx", buffer);
  });
};

main();
