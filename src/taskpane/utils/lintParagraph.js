const commonFont = { bold: true, name: "Arial", color: "#000000" };
const correctFormats = {
  "Heading 1": { font: { ...commonFont, size: 12 }, isListItem: true },
  "Heading 2": { font: { ...commonFont, size: 12 }, isListItem: true },
  "Heading 3": { font: { ...commonFont, size: 10 }, isListItem: true },
  "Heading 4": { font: { ...commonFont, size: 10, italic: true }, isListItem: true },
  // isListItem, spaceAfter, etc.
  // Normal: { bold: false, size: 12, name: "Times New Roman" },
};

const checkForErrors = (correctFormat, actualFormat, paragraph, setErrors) => {
  const correctProperties = Object.keys(correctFormat);
  correctProperties.forEach((property) => {
    const actual = actualFormat[property];
    const correct = correctFormat[property];
    if (typeof actual !== "object") {
      if (actual !== correct) {
        const error = { property, actual, correct, paragraph };
        setErrors((errors) => [...errors, error]);
      }
    } else {
      checkForErrors(correctFormat[property], actualFormat[property], paragraph, setErrors);
    }
  });
};

const lintParagraph = (paragraph, setErrors) => {
  const actualFormat = paragraph.toJSON();
  const correctFormat = correctFormats[paragraph.style];
  if (correctFormat !== undefined && paragraph.text !== "") {
    checkForErrors(correctFormat, actualFormat, paragraph, setErrors);
  }
};

export default lintParagraph;
