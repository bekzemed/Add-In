import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import { ButtonPrimaryExample } from "./Button";

/* global Word, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export const App: React.FunctionComponent<AppProps> = ({ title }: AppProps) => {
  const [listItems, setListItems] = React.useState([]);

  React.useEffect(() => {
    setListItems([
      {
        icon: "Ribbon",
        primaryText: "Achieve more with Office integration",
      },
      {
        icon: "Unlock",
        primaryText: "Unlock features and functionality",
      },
      {
        icon: "Design",
        primaryText: "Create and visualize like a pro",
      },
    ]);
  }, []);

  const click = async () => {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  };

  return (
    <div className="ms-welcome">
      <Header logo="assets/logo-filled.png" title={title} message="Welcome" />
      <HeroList message="Discover what this add-in can do for you today!" items={listItems}>
        <ButtonPrimaryExample />
      </HeroList>
    </div>
  );
};
