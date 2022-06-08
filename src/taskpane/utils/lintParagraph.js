// TODO: check more things:
// also check headers and footers
// also check bolded lists (Arial and Times New Roman mixed)
// also handle text inside tables (smaller I think?)
// also check for floating (blank) section titles

import formattingRules from "../formattingRules";

const getErrors = (correctFormat, actualFormat, paragraph, rule) => {
  const correctProperties = Object.keys(correctFormat);
  let errors = [];
  correctProperties.forEach((property) => {
    const actual = actualFormat[property];
    const correct = correctFormat[property];
    if (typeof actual !== "object") {
      if (actual !== correct) {
        const error = { property, actual, correct, paragraph, ruleName: rule.name };
        errors.push(error);
      }
    } else {
      const subErrors = getErrors(correctFormat[property], actualFormat[property], paragraph, rule);
      errors = [...errors, ...subErrors];
    }
  });
  return errors;
};

const lintParagraph = (paragraph, prevParagraph) => {
  const actualFormat = paragraph;
  const applicableRules = formattingRules.filter((rule) => rule.condition(paragraph, prevParagraph));
  let errors = [];
  if (applicableRules.length > 0) {
    applicableRules.forEach((rule) => {
      if (rule.debug) {
        console.log("rule:", rule.name);
        console.log("paragraph:", paragraph);
      }
      const newErrors = getErrors(rule.format, actualFormat, paragraph, rule);
      errors = [...errors, ...newErrors];
    });
  }
  return errors;
};

export default lintParagraph;
