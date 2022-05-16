// TODO: check more things:
// also check figure names and table names: Arial, bold 10pt and 12pt
// also check table text: Arial 10pt and 12pt
// also check headers and footers
// also check bolded lists (Arial and Times New Roman mixed)
// also handle text inside tables (smaller I think?)
// also check for floating (blank) section titles

const boldBlackArial = { bold: true, name: "Arial", color: "#000000" };

const formattingRules = [
  {
    condition: (p) => p.style === "Heading 1",
    format: { font: { ...boldBlackArial, size: 12, italic: false }, isListItem: true },
  },
  {
    condition: (p) => p.style === "Heading 2",
    format: { font: { ...boldBlackArial, size: 12, italic: false }, isListItem: true },
  },
  {
    condition: (p) => p.style === "Heading 3",
    format: { font: { ...boldBlackArial, size: 10, italic: false }, isListItem: true },
  },
  {
    condition: (p) => p.style === "Heading 4",
    format: { font: { ...boldBlackArial, size: 10, italic: true }, isListItem: true },
  },
  {
    condition: (p) => p.text.match(/Figure [0-9]/) && p.alignment === "Centered",
    format: { font: { ...boldBlackArial, size: 10 } },
  },
  {
    condition: (p) => p.text.match(/Table [0-9]/) && p.alignment === "Centered",
    format: { font: { ...boldBlackArial, size: 10 }, tableNestingLevel: 1 },
  },
];

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
  const actualFormat = paragraph;
  const applicableRules = formattingRules.filter((rule) => rule.condition(paragraph));
  let errors = [];
  if (applicableRules.length > 0 && paragraph.text !== "") {
    applicableRules.forEach((rule) => {
      const newErrors = getErrors(rule.format, actualFormat, paragraph);
      errors = [...errors, ...newErrors];
    });
  }
  return errors;
};

export default lintParagraph;
