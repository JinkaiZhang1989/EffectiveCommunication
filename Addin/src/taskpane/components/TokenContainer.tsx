import * as React from "react";

import { Button, ButtonType } from 'office-ui-fabric-react';
import { TextField, ITextField } from 'office-ui-fabric-react/lib/TextField';

import EffectiveCommunicationActionCreator from "../actioncreator/EffectiveComunicationActionCreator";

export enum TokenStatus {
    Initialize,
    Completed
}

export interface IState {
    status: TokenStatus;
}

export default class TokenContainer extends React.Component<{}, IState> {
    private textInput = React.createRef<ITextField>();

    constructor(props: any) {
        super(props);

        this.state = {
            status: TokenStatus.Completed
        };
    }

    public render(): JSX.Element {
        return (
            <div>
                <Button 
                    onClick={this.switch.bind(this)}
                    className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>
                    {this.title}
                </Button>
                {this.buildTokenInput()}
            </div>
        );
    }

    private buildTokenInput(): JSX.Element {
        if (this.state.status === TokenStatus.Completed) {
            return null;
        }

        return (
            <div>
                <TextField
                    placeholder="Paste token here"
                    componentRef={this.textInput}
                    defaultValue={this.token}
                    multiline />
                <Button
                    className='ms-welcome__action'
                    buttonType={ButtonType.hero}
                    iconProps={{ iconName: 'ChevronRight' }}
                    onClick={this.click.bind(this)}>
                    Set
                </Button>
            </div>
        );
    }

    private get title(): string {
        return this.state.status === TokenStatus.Completed ? "Token Initialized" : "Token";
    }

    private get token(): string {
        return EffectiveCommunicationActionCreator.getToken();
    }

    private switch(): void {
        console.log(`Hello world`);
        this.setState({status: TokenStatus.Initialize});
    }

    private click(): void {
        EffectiveCommunicationActionCreator.setToken(this.textInput.current.value);
        this.setState({status: TokenStatus.Completed});
    }
}
