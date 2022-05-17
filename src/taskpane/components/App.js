import React, { useState, useEffect } from "react";
import { DefaultButton, MessageBar, MessageBarButton, MessageBarType, Text } from "@fluentui/react";
import PropTypes from "prop-types";
import Progress from "./Progress";
import { ErrorList } from "./ErrorList";
import lintParagraph from "../utils/lintParagraph";
import catcher from "../utils/catchWordError";
import dayjs from "dayjs";
import relativeTime from "dayjs/plugin/relativeTime";

dayjs.extend(relativeTime);

/* global Word, require */

// TODO: test running "checkParagraphs" after every change
// TODO: consider creating a gmail add-in
const CustomMessageBar = ({ numErrors, checkStyles }) => {
  const message =
    numErrors === 0 ? (
      "No errors found!"
    ) : (
      <span>
        Found <b>{numErrors}</b> error{numErrors === 1 ? "" : "s"}.
      </span>
    );
  return (
    <MessageBar
      actions={
        <div>
          <MessageBarButton onClick={checkStyles}>Check again</MessageBarButton>
        </div>
      }
      messageBarType={numErrors > 0 ? MessageBarType.error : MessageBarType.success}
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
  const [running, setRunning] = useState(false);
  const [currentTime, setCurrentTime] = useState(Date.now());

  useEffect(() => {
    const interval = setInterval(() => setCurrentTime(Date.now()), 60 * 1000); // FIXME: reloading the add-in causes the grouped items to collapse
    return () => clearInterval(interval);
  }, []);

  const getParagraphs = async () => {
    return await Word.run(async (context) => {
      context.document.body.paragraphs.load("items");
      await context.sync();
      const paragraphs = context.document.body.paragraphs.items;
      paragraphs.forEach((paragraph) => {
        paragraph.load("font");
        paragraph.load("parentTableOrNullObject");
      });
      await context.sync();
      return paragraphs;
    }).catch(catcher);
  };

  const checkStyles = async () => {
    return Word.run(async (context) => {
      setRunning(true);
      setErrors([]);

      await context.sync();
      const paragraphs = await getParagraphs();
      paragraphs.forEach((paragraph, i) => {
        const prevParagraph = i > 0 ? paragraphs[i - 1] : null;
        const newErrors = lintParagraph(paragraph, prevParagraph);
        setErrors((errors) => [...errors, ...newErrors]);
      });

      setRunning(false);
      setLastChecked(new Date());
    }).catch(catcher);
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

  // update the last checked to say "checked X minutes ago" instead of "last checked at time X"
  return (
    <>
      {running && <Progress message="Checking styles..." />}
      {!running && (
        <>
          {lastChecked && <CustomMessageBar numErrors={errors.length} checkStyles={checkStyles} />}
          {lastChecked && (
            <Text variant="small">
              Last checked <u>{dayjs(lastChecked).fromNow()}</u>
            </Text>
          )}

          {!lastChecked && (
            <DefaultButton className="ms-welcome__action" onClick={checkStyles}>
              Check Styles
            </DefaultButton>
          )}
          {errors.length !== 0 && <ErrorList errors={errors} />}
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
