import React from "react";
import { Link } from "@fluentui/react";

const ErrorMessage = ({ error }) => {
  const { property, actual, correct, paragraph } = error;

  async function jumpToParagraph(paragraph) {
    await Word.run(async (context) => {
      // Select can be at the start or end of a range; this by definition moves the insertion point without selecting the range.
      context.document.body.paragraphs.load("items");
      await context.sync();
      const paragraphs = context.document.body.paragraphs.items;
      const paragraphOfInterest = paragraphs.find((p) => p._Id === paragraph._Id);
      paragraphOfInterest.select();

      await context.sync();
    });
  }

  return (
    <span>
      Font <b>{property}</b> should be <b>{String(correct)}</b> but is <b>{String(actual)}</b> <br></br>
      Location:{" "}
      <Link href="#" onClick={() => jumpToParagraph(paragraph)}>
        {paragraph.text}
      </Link>
      <br></br>
    </span>
  );
};

const ErrorList = ({ errors }) => {
  if (errors.length === 0) {
    return null;
  }
  return (
    <div>
      {errors.map((error, index) => (
        <>
          <br></br>
          <ErrorMessage key={index} error={error} />
        </>
      ))}
    </div>
  );
};

export default ErrorList;
