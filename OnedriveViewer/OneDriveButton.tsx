import React = require("react");


interface OneDriveViewerButtonProps {
    title: string;
    name: string;
    onClick(): void;
    onMount(obj: OneDriveViewerButton) : void;
}


const buttonCss = {
    'width': '185px',
    'boxShadow': 'inset 0px 1px 0px 0px #ffffff',
    'background': 'linear-gradient(to bottom, #ffffff 5%, #f6f6f6 100%)',
    'backgroundColor': '#ffffff',
    'borderRadius': '6px',
    'border': '1px solid #dcdcdc',
    'display': 'inline-block',
    'cursor': 'pointer',
    'color': '#666666',
    'fontFamily': 'Arial',
    'fontSize': '13px',
    'fontWeight': 'bold',
    'padding': '6px 24px',
    'textDecoration': 'none',
    'textShadow': '0px 1px 0px #ffffff',

    "&:hover": {
        'backgroundColor': "#f6f6f6",
        'background': 'linear - gradient(to bottom, #f6f6f6 5 %, #ffffff 100 %)'
    },
    "&:active": {
        'position': 'relative',
        'top': '1px'

    }
} as React.CSSProperties;

export interface OneDriveViewerButtonState {
    display: string;
    disabled: boolean;
}


export class OneDriveViewerButton extends React.Component<OneDriveViewerButtonProps, OneDriveViewerButtonState> {

    _props: OneDriveViewerButtonProps;

    constructor(props: OneDriveViewerButtonProps, state: OneDriveViewerButtonState ) {
        super(props);
        this.state = state; 
        this._props = props;
    }

    componentDidMount() {
        this._props.onMount(this);
    }


    render() {

        const divCss = {
            'display': this.state.display
        } as React.CSSProperties;

        return (
            <div style={divCss} >
                <button style={buttonCss} onClick={() => this._props.onClick()} disabled={this.state.disabled} >
                {this.props.title}
                </button>
            </div>
        );
    }
}

