import * as React from "react";
import { GroupedList, DetailsRow, SelectionMode } from "@fluentui/react";
import _ from "lodash";

import getGroups from "../utils/groupErrors";
import catcher from "../utils/catchWordError";

async function jumpToParagraph(paragraphId) {
  await Word.run(async (context) => {
    context.document.body.paragraphs.load("items"); // load all paragraphs
    await context.sync(); // wait for load to complete
    const paragraphs = context.document.body.paragraphs.items; // get all paragraphs
    const paragraph = paragraphs.find((p) => p._Id === paragraphId); // find paragraph with id
    paragraph.select(); // select paragraph
    await context.sync(); // wait for selection to complete
  }).catch(catcher);
}

export const ErrorMessage = ({ error }) => {
  const { property, actual, correct } = error;
  return (
    <span>
      Font <b>{property}</b> should be <b>{String(correct)}</b> but is <b>{String(actual)}</b>
    </span>
  );
};

export const ErrorList = ({ errors }) => {
  let groups = getGroups(errors);

  const renderErrorRow = (nestingDepth, item, itemIndex, group) => {
    const columns = [{ fieldName: "errorMessage" }];
    const updatedItem = { ...item, errorMessage: <ErrorMessage error={item} /> };
    return item && typeof itemIndex === "number" && itemIndex > -1 ? (
      <DetailsRow
        columns={columns}
        groupNestingDepth={nestingDepth}
        item={updatedItem}
        itemIndex={itemIndex}
        selectionMode={SelectionMode.none}
        compact={true}
        group={group}
      />
    ) : null;
  };

  const groupProps = {
    headerProps: {
      onGroupHeaderClick: (group) => {
        jumpToParagraph(group.key);
      },
    },
  };

  return (
    <div>
      <GroupedList
        items={errors}
        onRenderCell={renderErrorRow}
        selectionMode={SelectionMode.none}
        groups={groups}
        compact={true}
        groupProps={groupProps}
      />
    </div>
  );
};
