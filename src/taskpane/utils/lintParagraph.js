const commonFont = { bold: true, name: "Arial", color: "#000000" };
const correctFormats = {
  "Heading 1": { font: { ...commonFont, size: 12, italic: false }, isListItem: true },
  "Heading 2": { font: { ...commonFont, size: 12, italic: false }, isListItem: true },
  "Heading 3": { font: { ...commonFont, size: 10, italic: false }, isListItem: true },
  "Heading 4": { font: { ...commonFont, size: 10, italic: true }, isListItem: true },
  // Normal: { font: { size: 12, name: "Times New Roman" } },
  // isListItem, spaceAfter, etc.
};

const getErrors = (correctFormat, actualFormat, paragraph) => {
  const correctProperties = Object.keys(correctFormat);
  let errors = [];
  correctProperties.forEach((property) => {
    const actual = actualFormat[property];
    const correct = correctFormat[property];
    if (typeof actual !== "object") {
      if (actual !== correct) {
        const error = { property, actual, correct, paragraph };
        errors.push(error);
      }
    } else {
      const subErrors = getErrors(correctFormat[property], actualFormat[property], paragraph);
      errors = [...errors, ...subErrors];
    }
  });
  return errors;
};

const lintParagraph = (paragraph) => {
  const actualFormat = paragraph.toJSON();
  const correctFormat = correctFormats[paragraph.style];
  let errors = [];
  if (correctFormat !== undefined && paragraph.text !== "") {
    errors = getErrors(correctFormat, actualFormat, paragraph);
  }
  return errors;
};

export default lintParagraph;
