const express = require('express');
const cors = require('cors');
const {Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell, WidthType} = require('docx');
const htmlparser2 = require('htmlparser2');

const app = express();
const port = process.env.PORT || 8000;

app.use(cors());
app.use(express.json({limit: '1mb'}));

function createFormattedDocx(text) {
  function parseInlineFormatting(str) {
    const textRuns = [];
    const regex = /\*\*(.+?)\*\*|{red}(.+?){\/red}|\[(.+?)\]\((.+?)\)/g;
    let lastIndex = 0;
    for (const match of str.matchAll(regex)) {
      const [full, bold, red, linkText, linkUrl] = match;
      const start = match.index;
      if (start > lastIndex) {
        textRuns.push(
          new TextRun({
            text: str.slice(lastIndex, start),
            font: 'Calibri',
            size: 19,
            color: '000000',
          })
        );
      }
      if (bold) {
        textRuns.push(
          new TextRun({
            text: bold,
            bold: true,
            font: 'Calibri',
            size: 19,
            color: '000000',
          })
        );
      } else if (red) {
        textRuns.push(
          new TextRun({
            text: red,
            font: 'Calibri',
            size: 19,
            color: 'FF0000',
          })
        );
      } else if (linkText && linkUrl) {
        textRuns.push(
          new ExternalHyperlink({
            link: linkUrl,
            children: [
              new TextRun({
                text: linkText,
                style: 'Hyperlink',
                font: 'Calibri',
                size: 19,
              }),
            ],
          })
        );
      }
      lastIndex = match.index + full.length;
    }
    if (lastIndex < str.length) {
      textRuns.push(
        new TextRun({
          text: str.slice(lastIndex),
          font: 'Calibri',
          size: 19,
          color: '000000',
        })
      );
    }
    return textRuns;
  }

  const children = [];
  const dom = htmlparser2.parseDocument(text);
  function walk(nodes, parentParagraph, listContext = []) {
    for (const node of nodes) {
      if (node.type === 'tag') {
        if (node.name === 'p') {
          const paraRuns = [];
          walk(node.children || [], paraRuns, listContext);
          if (paraRuns.length > 0) {
            children.push(
              new Paragraph({
                children: paraRuns,
                spacing: {before: 0, after: 100, line: 276},
              })
            );
          }
        } else if (node.name === 'a' && node.attribs && node.attribs.href) {
          const linkText = htmlparser2.DomUtils.getText(node);
          const linkUrl = node.attribs.href;
          const linkRun = new ExternalHyperlink({
            link: linkUrl,
            children: [
              new TextRun({
                text: linkText,
                style: 'Hyperlink',
                font: 'Calibri',
                size: 19,
              }),
            ],
          });
          if (parentParagraph) {
            parentParagraph.push(linkRun);
          } else {
            children.push(new Paragraph({children: [linkRun]}));
          }
        } else if (node.name === 'table') {
          const rows = [];
          const trNodes = node.children.filter((n) => n.type === 'tag' && n.name === 'tbody')[0]?.children || node.children;
          for (const tr of trNodes) {
            if (tr.type === 'tag' && tr.name === 'tr') {
              const cells = [];
              for (const td of tr.children) {
                if (td.type === 'tag' && (td.name === 'td' || td.name === 'th')) {
                  const cellRuns = [];
                  walk(td.children || [], cellRuns, listContext);
                  cells.push(
                    new TableCell({
                      children: [new Paragraph({children: cellRuns})],
                      width: {size: 50, type: WidthType.PERCENTAGE},
                    })
                  );
                }
              }
              if (cells.length > 0) {
                rows.push(new TableRow({children: cells}));
              }
            }
          }
          if (rows.length > 0) {
            children.push(
              new Table({
                rows,
                width: {size: 100, type: WidthType.PERCENTAGE},
              })
            );
          }
        } else if (node.name === 'ol' || node.name === 'ul') {
          // Ordered or unordered list
          walk(node.children || [], parentParagraph, [...listContext, node.name]);
        } else if (node.name === 'li') {
          // List item
          const liRuns = [];
          walk(node.children || [], liRuns, listContext);
          // Determine bullet or number
          let bullet = '';
          let numbering = undefined;
          let indent = listContext.length > 0 ? listContext.length - 1 : 0;
          if (listContext[listContext.length - 1] === 'ol') {
            numbering = {reference: 'numbered-list', level: indent};
          } else {
            bullet = '\u2022';
          }
          if (liRuns.length > 0) {
            children.push(
              new Paragraph({
                children: numbering ? liRuns : [new TextRun({text: bullet + ' ', font: 'Calibri', size: 19}), ...liRuns],
                numbering: numbering,
                indent: {left: 720 * indent},
                spacing: {before: 0, after: 100, line: 276},
              })
            );
          }
        } else {
          walk(node.children || [], parentParagraph, listContext);
        }
      } else if (node.type === 'text') {
        const trimmed = node.data;
        if (trimmed && parentParagraph) {
          parentParagraph.push(...parseInlineFormatting(trimmed));
        } else if (trimmed) {
          children.push(
            new Paragraph({
              children: parseInlineFormatting(trimmed),
              spacing: {before: 0, after: 100, line: 276},
            })
          );
        }
      }
    }
  }
  walk(dom.children);

  const doc = new Document({
    numbering: {
      config: [
        {
          reference: 'numbered-list',
          levels: [
            {
              level: 0,
              format: 'decimal',
              text: '%1.',
              alignment: 'left',
              style: {
                paragraph: {
                  indent: {left: 0, hanging: 360},
                },
              },
            },
            {
              level: 1,
              format: 'decimal',
              text: '%2.',
              alignment: 'left',
              style: {
                paragraph: {
                  indent: {left: 720, hanging: 360},
                },
              },
            },
            {
              level: 2,
              format: 'decimal',
              text: '%3.',
              alignment: 'left',
              style: {
                paragraph: {
                  indent: {left: 1440, hanging: 360},
                },
              },
            },
          ],
        },
      ],
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 640,
              bottom: 640,
              left: 640,
              right: 640,
            },
          },
        },
        children,
      },
    ],
  });
  return Packer.toBuffer(doc);
}

app.post('/generate-docx', async (req, res) => {
  try {
    const {text} = req.body || {};
    const download = req.query.download === 'true';

    if (typeof text !== 'string') {
      return res.status(400).json({error: 'Input "text" must be a string.'});
    }

    const buffer = await createFormattedDocx(text);

    if (download) {
      res.setHeader('Content-Disposition', 'attachment; filename="output.docx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      return res.end(buffer);
    }

    const base64 = buffer.toString('base64');
    res.json({file: base64});
  } catch (err) {
    res.status(500).json({error: 'Failed to generate docx.'});
  }
});

app.listen(port);
