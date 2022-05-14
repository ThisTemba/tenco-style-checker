import React, { useState } from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
import lintParagraph from "../utils/lintParagraph";
import { GroupedListBasicExample } from "./ErrorList";
import { MessageBar, MessageBarButton, MessageBarType } from "@fluentui/react";

/* global Word, require */

// TODO: test running "checkParagraphs" after every change
// TODO: consider creating a gmail add-in
const CustomMessageBar = ({ numErrors, checkStyles }) => {
  const message = numErrors === 0 ? "No errors found!" : `${numErrors} error${numErrors === 1 ? "" : "s"} found.`;
  return (
    <MessageBar
      actions={
        <div>
          <MessageBarButton onClick={checkStyles}>Check again</MessageBarButton>
        </div>
      }
      messageBarType={numErrors > 0 ? MessageBarType.warning : MessageBarType.success}
      isMultiline={false}
    >
      {message}
    </MessageBar>
  );
};

const App = (props) => {
  const { title, isOfficeInitialized } = props;
  const [errors, setErrors] = useState([]);
  const [lastChecked, setLastChecked] = useState(null);
  const [isRunning, setIsRunning] = useState(false);

  const checkStyles = async () => {
    return Word.run(async (context) => {
      setIsRunning(true);
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
      setIsRunning(false);
      setLastChecked(new Date());

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
      {isRunning && <Progress message="Checking styles..." />}
      {!isRunning && (
        <>
          {lastChecked && <CustomMessageBar numErrors={errors.length} checkStyles={checkStyles} />}
          {lastChecked && <p>Last checked: {lastChecked.toLocaleTimeString()}</p>}
          {!lastChecked && (
            <DefaultButton className="ms-welcome__action" onClick={checkStyles}>
              Check Styles
            </DefaultButton>
          )}
          {errors.length !== 0 && <GroupedListBasicExample errors={errors} />}
        </>
      )}
    </>
  );
};

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};

export default App;
