import * as React from 'react';

import Header from './Header';

export default class ContributionContainer extends React.Component<{}, {}> {
  public render() {
    return (
      <div className='ms-welcome'>
        <Header logo='assets/logo-filled.png' title="Contributions!" message='Welcome' />
      </div>
    );
  }
}
