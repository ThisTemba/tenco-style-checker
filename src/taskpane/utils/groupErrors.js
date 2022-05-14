import _ from "lodash";

const truncateText = (text, len) => {
  if (text.length > len) {
    return text.substring(0, len) + "...";
  } else return text;
};

function getl1Groups(l2Groups) {
  let l1Groups = [];
  const groupedGroups = _.groupBy(l2Groups, (group) => group.data.style);

  _.mapValues(groupedGroups, (subGroups, style) => {
    const key = style;
    const name = style;
    const count = _.sumBy(subGroups, (group) => group.count);
    const startIndex = 0;
    const children = subGroups;
    const group = { count, key, startIndex, name, children, level: 0, isCollapsed: true };
    l1Groups.push(group);
    console.log("group", group);
    console.log("groups", l1Groups);
  });
  console.log("groups", l1Groups);
  return l1Groups;
}

function getl2Groups(errors) {
  const groupedErrors = _.groupBy(errors, (error) => error.paragraph._Id);
  let currentStartIndex = 0;
  let subGroups = [];
  _.mapValues(groupedErrors, (errors, i) => {
    const first = errors[0];

    const count = errors.length;
    const key = first.paragraph._Id;
    const name = truncateText(first.paragraph.text, 25);
    const startIndex = currentStartIndex;
    const data = { style: first.paragraph.style };
    const group = { count, key, name, startIndex, data, level: 1 };
    currentStartIndex += count;
    subGroups.push(group);
  });
  return subGroups;
}

function getGroups(errors) {
  let l2Groups = getl2Groups(errors);
  let l1Groups = getl1Groups(l2Groups);
  return l1Groups;
}

export default getGroups;
