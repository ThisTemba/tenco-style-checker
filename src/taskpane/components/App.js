import React, { useState } from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
import lintParagraph from "../utils/lintParagraph";
import ErrorList from "./ErrorList";

/* global Word, require */

// TODO: group errors by paragraph (heading)
// TODO: display errors grouped (in simple way, don't waste time)
// TODO: test running "checkParagraphs" after every change
// TODO: consider creating a gmail add-in

const App = (props) => {
  const { title, isOfficeInitialized } = props;
  const [errors, setErrors] = useState([]);

  const checkParagraphs = async () => {
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs.load("items");

      await context.sync();

      paragraphs.items.forEach((paragraph) => {
        paragraph.load("font");
      });

      await context.sync();

      setErrors([]);
      paragraphs.items.forEach((paragraph) => {
        const newErrors = lintParagraph(paragraph);
        setErrors((errors) => [...errors, ...newErrors]);
      });

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
      <ErrorList errors={errors} />
    </>
  );
};

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App;
