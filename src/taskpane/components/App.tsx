import { DefaultButton, PrimaryButton, TextField } from "@fluentui/react";
import OpenAI from "openai";
import React, { FC, FormEvent, useRef, useState } from "react";
import Progress from "./Progress";

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global require, Office, console */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export const App: FC<AppProps> = ({ title, isOfficeInitialized }) => {
  const openAi = useRef<OpenAI>(new OpenAI({ dangerouslyAllowBrowser: true, apiKey: "" }));
  const [token, setToken] = useState<string | undefined>(undefined);
  const [errorText, setErrorText] = useState<string>("");
  const [mailContent, setMailContent] = useState<string>("");
  const [completion, setCompletion] = useState<string | null>(null);

  Office.context.mailbox.item?.body.getAsync("text", (result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      setErrorText("There was an error getting the content of the email");
    } else {
      setMailContent(result.value);
    }
  });

  const initialize = () => {
    if (!token) return setErrorText("Provide a token");

    openAi.current.apiKey = token;
  };

  const onTokenChange = (_: FormEvent, token: string | undefined) => {
    setErrorText("");
    setToken(token);
  };

  const generate = async () => {
    try {
      const completion = await openAi.current.chat.completions.create({
        model: "gpt-3.5-turbo",
        messages: [
          {
            role: "system",
            content:
              "Du skal komme med et svar på en mail-korrespondance. Du bliver givet en hel mailtråd, og skal forsøge at komme med det bedste svar",
          },
          {
            role: "user",
            content: mailContent,
          },
        ],
      });

      setCompletion(completion.choices[0].message.content);
    } catch (err: unknown) {
      if (err instanceof OpenAI.APIError) {
        setErrorText(err.message);
      } else {
        setErrorText("An unknown error occured");
      }
    }
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
    <div className="container">
      <TextField
        label="OpenAPI token"
        autoComplete="off"
        spellCheck="false"
        onChange={onTokenChange}
        errorMessage={errorText}
      ></TextField>

      <DefaultButton text="Initialize" allowDisabledFocus onClick={initialize} />

      <PrimaryButton className="generate-button" text="Generate response" allowDisabledFocus onClick={generate} />

      {completion}
    </div>
  );
};
