import * as React from 'react';

import Configuration from "../core/Configuration";
import TokenContainer from "./TokenContainer";
import ContributionContainer from "./ContributionContainer";

export default class App extends React.Component<{}, {}> {
  public render(): JSX.Element {
    if (!Configuration.officejsHasBeenInitialized || Configuration.isInBackgroundMode) {
      return <div />;
    }

    return (
      <div>
        <ContributionContainer />
        <TokenContainer/>
      </div>
    );
  }
}
