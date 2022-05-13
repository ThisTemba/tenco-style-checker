import * as React from "react";
import { GroupedList } from "@fluentui/react/lib/GroupedList";
import { DetailsRow } from "@fluentui/react/lib/DetailsList";
import { SelectionMode } from "@fluentui/react/lib/Selection";
import _ from "lodash";

const toggleStyles = { root: { marginBottom: "20px" } };
const groupCount = 3;
const groupDepth = 1;

// const groups = createGroups(groupCount, groupDepth, 0, groupCount);

export const ErrorMessage = ({ error }) => {
  const { property, actual, correct, paragraph } = error;
  return (
    <span>
      Font <b>{property}</b> should be <b>{String(correct)}</b> but is <b>{String(actual)}</b>
    </span>
  );
};

async function jumpToParagraph(paragraphId) {
  await Word.run(async (context) => {
    // Select can be at the start or end of a range; this by definition moves the insertion point without selecting the range.
    context.document.body.paragraphs.load("items");
    await context.sync();
    const paragraphs = context.document.body.paragraphs.items;
    const paragraph = paragraphs.find((p) => p._Id === paragraphId);
    paragraph.select();

    await context.sync();
  });
}

export const GroupedListBasicExample = ({ errors }) => {
  let groups = getGroups(errors);

  const columns = [{ fieldName: "errorMessage" }];

  const onRenderCell = (nestingDepth, item, itemIndex, group) => {
    item = { ...item, errorMessage: <ErrorMessage error={item} /> };

    return item && typeof itemIndex === "number" && itemIndex > -1 ? (
      <DetailsRow
        columns={columns}
        groupNestingDepth={nestingDepth}
        item={item}
        itemIndex={itemIndex}
        selectionMode={SelectionMode.none}
        compact={true}
        group={group}
      />
    ) : // <ErrorMessage error={item} />
    null;
  };

  return (
    <div>
      <GroupedList
        items={errors}
        // eslint-disable-next-line react/jsx-no-bind
        onRenderCell={onRenderCell}
        selectionMode={SelectionMode.none}
        groups={groups}
        compact={true}
        groupProps={{
          headerProps: {
            onGroupHeaderClick: (group) => {
              jumpToParagraph(group.key);
            },
          },
        }}
      />
    </div>
  );
};

const getShortName = (text, numChars) => {
  if (text.length > numChars) {
    return text.substring(0, numChars) + "...";
  } else return text;
};

function getGroups(errors) {
  const groupedErrors = _.groupBy(errors, (error) => error.paragraph._Id);
  let currentStartIndex = 0;
  let subGroups = [];
  _.mapValues(groupedErrors, (errors, i) => {
    const first = errors[0];
    const count = errors.length;
    const key = first.paragraph._Id;
    const name = getShortName(first.paragraph.text, 25);
    const startIndex = currentStartIndex;
    const data = { style: first.paragraph.style };
    const group = { count, key, name, startIndex, data, level: 1 };
    currentStartIndex += count;
    subGroups.push(group);
  });

  let groups = [];
  const groupedGroups = _.groupBy(subGroups, (group) => group.data.style);

  _.mapValues(groupedGroups, (subGroups, style) => {
    const key = style;
    const name = style;
    const count = _.sumBy(subGroups, (group) => group.count);
    const startIndex = 0;
    const children = subGroups;
    const group = { count, key, startIndex, name, children, level: 0, isCollapsed: true };
    groups.push(group);
    console.log("group", group);
    console.log("groups", groups);
  });
  console.log("groups", groups);

  return groups;
}
