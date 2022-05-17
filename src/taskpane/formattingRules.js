const boldBlackArial = { bold: true, name: "Arial", color: "#000000" };

const formattingRules = [
  {
    name: "Heading 1",
    condition: (p) => p.style === "Heading 1",
    format: { font: { ...boldBlackArial, size: 12, italic: false }, isListItem: true },
    debug: true,
  },
  {
    name: "Heading 2",
    condition: (p) => p.style === "Heading 2",
    format: { font: { ...boldBlackArial, size: 12, italic: false }, isListItem: true },
  },
  {
    name: "Heading 3",
    condition: (p) => p.style === "Heading 3",
    format: { font: { ...boldBlackArial, size: 10, italic: false }, isListItem: true },
  },
  {
    name: "Heading 4",
    condition: (p) => p.style === "Heading 4",
    format: { font: { ...boldBlackArial, size: 10, italic: true }, isListItem: true },
  },
  {
    name: "Figure Numbers",
    condition: (p) => p.text.match(/Figure [0-9]/) && p.alignment === "Centered",
    format: { font: { ...boldBlackArial, size: 10 } },
  },
  {
    name: "Table Numbers",
    condition: (p) => p.text.match(/Table [0-9]/) && p.alignment === "Centered",
    format: { font: { ...boldBlackArial, size: 10 }, tableNestingLevel: 1 },
  },
  {
    name: "Figure & Table Captions",
    condition: (p, pp) => pp && pp.text.match(/Table [0-9]|Figure [0-9]/) && pp.alignment === "Centered",
    format: { font: { ...boldBlackArial, size: 12 }, alignment: "Centered" },
  },
  // {
  //   name: "Footers",
  //   condition: (p) => p.type === "Footer",
  //   // format: { text: "Tenco\tPage 1\tMarch 18 2022" },
  //   formatCheckers: [
  //     {
  //       errorMessage: "Footer should be centered",
  //       check: (p) => p.alignment === "Centered",
  //     },
  //     {
  //       errorMessage: "Footer should be in Arial",
  //       check: (p) => p.font.name === "Arial",
  //     },
  //   ],
  //   debug: true,
  // },
];

export default formattingRules;
