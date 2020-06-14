import React = require("react");
import { debug } from "util";


interface OneDriveViewerFrameProps {
    onMount(obj: OneDrivePreviewFrame): void;
}

export interface OneDriveViewerFrameState {
    display: string;
    src: string;
}



export class OneDrivePreviewFrame extends React.Component<OneDriveViewerFrameProps, OneDriveViewerFrameState> {

    _props: OneDriveViewerFrameProps;

    constructor(props: OneDriveViewerFrameProps, state: OneDriveViewerFrameState) {
        super(props);
        this.state = state;    
        this._props = props;
    }    

    componentDidMount() {
        this._props.onMount(this);
    }

    render() {
        
        const styles = {
            containerStyle: {
                display: this.state.display,
                'width': '100%',
                'height': '800px',
            }
        };
        const { containerStyle } = styles;

        return (
            <div id="fraPreview" style={containerStyle}>
                <iframe id="preview" width="100%" height="100%" src={this.state.src} />
            </div>
        );
    }
}

