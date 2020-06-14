import React = require("react");
import graph = require("@microsoft/microsoft-graph-client");
import { GraphFileBrowser } from '@microsoft/file-browser';




export interface IOneDriveViewerFileSelectorProps {
    endpoint: string;
    accesToken: string;
    onMount(obj: OneDriveViewerSelector): void;
    onCancel(): void;
    onItemSelected(id: string): void;
    onGetAccessToken(): Promise<string>;
}

export interface IOneDriveViewerSelectorState {
    display: string;
}

export class OneDriveViewerSelector extends React.Component<IOneDriveViewerFileSelectorProps, IOneDriveViewerSelectorState> {

    private _props: IOneDriveViewerFileSelectorProps;

    constructor(props: IOneDriveViewerFileSelectorProps, state: IOneDriveViewerSelectorState) {
        super(props);
        this.state = state;
        //display to none by default
        this.state = {
            'display' : 'none'
        };        

        this._props = props;        
    }

    componentDidMount() {
        this._props.onMount(this);
    }

    public render(): React.ReactElement<IOneDriveViewerFileSelectorProps> {

        const styles = {
            containerStyle: {
                display: this.state.display,
            }
        };
        const { containerStyle } = styles;


        return (
            <div id="fraFileSelect" style={containerStyle} >
            <GraphFileBrowser
                    getAuthenticationToken={this._props.onGetAccessToken}
                    onSuccess={(keys: any) => { this._props.onItemSelected(keys[0]["driveItem_203"][2]) }}
                    onCancel={this._props.onCancel}
                    endpoint={this._props.endpoint}
                    selectionMode="single"
                    itemMode="files"
                />
            </div>
        );
    }
}

