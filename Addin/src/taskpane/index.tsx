import 'office-ui-fabric-react/dist/css/fabric.min.css';
import App from './components/App';
import Configuration from "./core/Configuration";
import EffectiveCommunicationActionCreator from "./actioncreator/EffectiveComunicationActionCreator";

import { AppContainer } from 'react-hot-loader';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

initializeIcons();

const title = 'Contoso Task Pane Add-in';

const render = (Component) => {
    ReactDOM.render(
        <AppContainer>
            <Component title={title} isOfficeInitialized={Configuration.officejsHasBeenInitialized} />
        </AppContainer>,
        document.getElementById('container')
    );
};

/* Render application after Office initializes */
Office.initialize = () => {
    Configuration.officejsHasBeenInitialized = true;
    EffectiveCommunicationActionCreator.Load();
    render(App);
};
