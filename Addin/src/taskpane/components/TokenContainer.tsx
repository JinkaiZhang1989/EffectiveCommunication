import * as React from "react";
import * as Flux from "flux/utils";

export enum TokenStatus {
    Initialize,
    Completed
}

export interface IState {
    status: TokenStatus;
}

export default class TokenContainer extends React.Component<{}, IState> {
    constructor(props: any) {
        super(props);

        this.state = {
            status: TokenStatus.Completed
        };
    }

    public render(): JSX.Element {
        return (
            <div>
                <h1 className='ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary'>{this.title}</h1>
                {this.buildTokenInput()}
            </div>
        );
    }

    private buildTokenInput(): JSX.Element {
        if (this.state.status === TokenStatus.Completed) {
            return null;
        }

        
    }

    private get title(): string {
        return this.state.status === TokenStatus.Completed ? "Token Initialized" : "Token";
    }
}
