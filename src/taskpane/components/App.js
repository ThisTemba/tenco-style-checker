import React, { useState } from "react";
import PropTypes from "prop-types";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";

/* global Word, require */
const correctFormats = {
  "Heading 1": { font: { bold: true, size: 12, name: "Arial" }, isListItem: true },
  "Heading 2": { font: { bold: true, size: 12, name: "Arial" }, isListItem: true },
  "Heading 3": { font: { bold: true, size: 10, name: "Arial" }, isListItem: true },
  "Heading 4": { font: { bold: true, size: 10, name: "Arial", italic: true }, isListItem: true },
  // TODO: make this check not just font props but also paragraph props such as
  // isListItem, spaceAfter, etc.
  // Normal: { bold: false, size: 12, name: "Times New Roman" },
};

const getErrorMessage = ({ property, correct, actual, location }) => {
  return `Font ${property} should be ${correct} but is ${actual} at: ${getBreadcrumb(location)}`;
};

const ErrorMessage = ({ error }) => {
  const { property, actual, correct, location } = error;

  return (
    <span>
      Font <b>{property}</b> should be <b>{String(correct)}</b> but is <b>{String(actual)}</b> at:{" "}
      <i>{getBreadcrumb(location)}</i>
    </span>
  );
};

const getBreadcrumb = (list) => {
  return list.filter((item) => item).join(" > ");
};

const updateLocation = (currentLocation, paragraph) => {
  let headings = [...currentLocation];
  const style = paragraph.style;
  if (style === "Heading 1") {
    headings = [paragraph.text];
  } else if (style === "Heading 2") {
    headings = [...headings.slice(0, 1), paragraph.text];
  } else if (style === "Heading 3") {
    headings = [...headings.slice(0, 2), paragraph.text];
  } else if (style === "Heading 4") {
    headings = [...headings.slice(0, 3), paragraph.text];
  }
  return headings;
};

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

      let currentLocation = [];
      paragraphs.items.forEach((paragraph) => {
        const actualFormat = paragraph.toJSON();
        const correctFormat = correctFormats[paragraph.style];
        currentLocation = updateLocation(currentLocation, paragraph);

        if (correctFormat) {
          console.log(actualFormat, "Hello World", correctFormat);
          Object.keys(correctFormat).forEach((property) => {
            const actual = actualFormat[property];
            const correct = correctFormat[property];
            if (typeof actual !== "object") {
              console.log(actual, correct);
              if (actual !== correct) {
                const error = { property, actual, correct, location: currentLocation };
                setErrors((errors) => [...errors, error]);
              }
            } else {
              console.log("object");
              console.log(property);
            }
          });
        }
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

export default App;

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
