import React, { useState } from "react";
import PropTypes from "prop-types";
import { DefaultButton, Link } from "@fluentui/react";
import Progress from "./Progress";
import lintParagraph from "../utils/lintParagraph";

/* global Word, require */

// TODO: group errors by paragraph (heading)
// TODO: display errors grouped (in simple way, don't waste time)
// TODO: test running "checkParagraphs" after every change
// TODO: consider creating a gmail add-in

const ErrorMessage = ({ error }) => {
  const { property, actual, correct, paragraph } = error;

  return (
    <span>
      Font <b>{property}</b> should be <b>{String(correct)}</b> but is <b>{String(actual)}</b> <br></br>
      Location:{" "}
      <Link href="#" onClick={() => jumpToParagraph(paragraph)}>
        {paragraph.text}
      </Link>
    </span>
  );
};

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

const App = (props) => {
  const { title, isOfficeInitialized } = props;
  const [errors, setErrors] = useState([]);

  const checkParagraphs = async () => {
    return Word.run(async (context) => {
      setErrors([]);
      const paragraphs = context.document.body.paragraphs.load("items");

      await context.sync();

      paragraphs.items.forEach((paragraph) => {
        paragraph.load("font");
      });

      await context.sync();

      paragraphs.items.forEach((paragraph) => {
        lintParagraph(paragraph, setErrors);
      });
      console.log("done!");
      await context.sync();
    });
  };

  if (!isOfficeInitialized) {
    return (
      <Progress
        title={title}
        logo={require("./../../../assets/logo-filled.png")}
        message="Please sideload your addin to see app body."
      />
    );
  }

  return (
    <>
      <DefaultButton className="ms-welcome__action" onClick={checkParagraphs}>
        Check Paragraphs
      </DefaultButton>
      {errors.length > 0 && (
        <div>
          <h2>Errors</h2>
          <ul>
            {errors.map((error) => (
              <>
                <li>
                  <ErrorMessage error={error} />
                </li>
                <br></br>
              </>
            ))}
          </ul>
        </div>
      )}
    </>
  );
};

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App;
