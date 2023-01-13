import { ITextFieldStyles, Stack, TextField } from "@fluentui/react";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import * as React from "react";

const textFieldStyles: Partial<ITextFieldStyles> = { fieldGroup: { marginBottom: 20 } };

export const ButtonPrimaryExample: React.FunctionComponent = () => {
  const [firstName, setFirstName] = React.useState("");
  const [lastName, setLastName] = React.useState("");

  const onFirstNameChange = React.useCallback(
    (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
      setFirstName(newValue || "");
    },
    []
  );

  const onSecondNameChange = React.useCallback(
    (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
      setLastName(newValue || "");
    },
    []
  );

  const insertText = async () => {
    // In the click event, write text to the document.
    // eslint-disable-next-line no-undef
    await Word.run(async (context) => {
      let body = context.document.body;
      // eslint-disable-next-line no-undef
      body.insertParagraph(`Welcome ${firstName} ${lastName}!!!`, Word.InsertLocation.end);
      await context.sync();
    });
  };

  return (
    <div>
      <Stack>
        <TextField
          label="Enter your first name"
          value={firstName}
          onChange={onFirstNameChange}
          styles={textFieldStyles}
        />

        <TextField
          label="Enter your last name"
          value={lastName}
          onChange={onSecondNameChange}
          styles={textFieldStyles}
        />
        <PrimaryButton data-automation-id="test" text="Submit" onClick={insertText} />
      </Stack>
    </div>
  );
};
