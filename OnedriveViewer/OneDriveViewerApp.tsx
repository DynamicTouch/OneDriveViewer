import { IInputs } from "./generated/ManifestTypes";
import { OneDriveViewerButton } from "./OneDriveButton";
import { OneDrivePreviewFrame } from "./OneDrivePreviewFrame";
import { OneDriveViewerSelector } from "./OneDriveSelector";
import React = require("react");
import graph = require("@microsoft/microsoft-graph-client");


interface IOnedriveViewerAppState {
    showSelectFrame: boolean;
    enableButtonSelect: boolean;
    showPreview: boolean;
    showButtonOpen: boolean;
    showButtonDownload: boolean;
    webUrl: string;
    previewUrl: string;
    downloadUrl: string;
    accessToken: string;
    fileId: string;
}


//global variables
let _appProperties: IOnedriveViewerAppProperties;
let _state: IOnedriveViewerAppState;


export interface IOnedriveViewerAppProperties {
    context: ComponentFramework.Context<IInputs>;
    node: HTMLDivElement;
    accessToken: string;    
    endpoint: string;
    path: string;
    fileId: string;    
    enableSelect: boolean;
    onFileSelected(previewUrl: string, webUrl: string, id: string): void;
}


export const getAuthenticatedClient = (accessToken: string) => {
    return graph.Client.init({
        authProvider: (done: any) => {
            done(null, accessToken);
        }
    });
}

export const getFilePreviewUrl = async (accessToken: string, id: string) => {
    const client = getAuthenticatedClient(accessToken);
    const itemPreviewInfo = {};
    //const res = await client.api("/me/drive/items/" + id + "/preview")
    const res = await client.api("/" + _appProperties.path + "/drive/items/" + id + "/preview")
        .post(itemPreviewInfo);
    return res;
}

export const getFileDetails = async (accessToken: string, id: string) => {
    const client = getAuthenticatedClient(accessToken);
    const itemInfo = {};
    //const res = await client.api("/me/drive/items/" + id)
    const res = await client.api("/" + _appProperties.path + "/drive/items/" + id)
        .post(itemInfo);
    return res;
}



async function onGetAccessToken(): Promise<string> {
    return Promise.resolve(_state.accessToken);
}


export const OneDriveViewerApp: React.FC<IOnedriveViewerAppProperties> = (fileViewProps: IOnedriveViewerAppProperties) => {

    let buttonOpen: OneDriveViewerButton;
    let buttonDownload: OneDriveViewerButton;
    let buttonSelect: OneDriveViewerButton;
    let preView: OneDrivePreviewFrame;
    let selector: OneDriveViewerSelector;


    _appProperties = fileViewProps;
    console.log(_appProperties);


    //initial state
    _state = {
        enableButtonSelect: fileViewProps.enableSelect,
        showButtonDownload: false,
        showButtonOpen: false,
        showPreview: false,
        showSelectFrame: false,        
        fileId : fileViewProps.fileId,
        accessToken : fileViewProps.accessToken,
        downloadUrl : '',
        previewUrl : '',
        webUrl : ''
    };
    console.log(_state);


    const fraButtonsCss = {
        'borderSpacing': '10px'
    } as React.CSSProperties;


    function resizeWindow() {

        window.dispatchEvent(new Event('resize'));
        /*
        const evt = document.createEvent('UIEvents');
        evt.initEvent('resize', true, false);
        window.dispatchEvent(evt);*/
    }

    function updateComponentStates() {

        if (_state.showSelectFrame ) {
            _state.showButtonDownload = false;
            _state.showButtonOpen = false;
            _state.showPreview = false;
        }
        else { 
            _state.showButtonDownload = (_state.downloadUrl !== '');
            _state.showButtonOpen = (_state.webUrl !== '');
            _state.showPreview = (_state.previewUrl !== '');            
        }

        buttonDownload?.setState(
            {
                'display': _state.showButtonDownload ? "block" : "none",
                'disabled': false
            }
        );

        buttonOpen?.setState(
            {
                'display': _state.showButtonOpen ? "block" : "none",
                'disabled': false
            }
        );
        buttonSelect?.setState(
            {
                'display': _state.showButtonOpen ? "block" : "none",
                'disabled': !_state.enableButtonSelect
            }
        );

        selector?.setState(
            {
                "display": _state.showSelectFrame ? "block" : "none",
            }
        );

        preView?.setState(
            {
                "display": _state.showPreview ? "block" : "none",
                "src": _state.previewUrl
            }
        );

        resizeWindow();
    }

    function onCancelFileSelected() {

        _state.showSelectFrame = false;

        updateComponentStates();
    }

    function OnShowItem(previewUrl: string,
                        webUrl: string,
                        downloadUrl: string) {


        _state.webUrl = webUrl;
        _state.downloadUrl = downloadUrl;
        _state.previewUrl = previewUrl;
        _state.showSelectFrame = false;

        //update button states
        updateComponentStates();

        //callback
        _appProperties.onFileSelected(_state.previewUrl, _state.webUrl, _state.fileId);

    }

    function loadItem() {
        if (_state.fileId !== '') {
            getFileDetails(_state.accessToken, _state.fileId).then(details => {
                getFilePreviewUrl(_state.accessToken, _state.fileId).then(
                    preView => {
                        OnShowItem(preView.getUrl,
                            details.webUrl, details["@microsoft.graph.downloadUrl"]);
                    });
            }
            );
        };
    }

    function onItemSelected(id: string) {
        _state.fileId = id;

        loadItem();
    }

    function onButtonSelectFile() {
        _state.showSelectFrame = !_state.showSelectFrame;
        updateComponentStates();
    }

    function onButtonOpenFile() {
        window.open(_state.webUrl);
    }

    function onButtonDownloadFile() {
        window.open(_state.downloadUrl);
    }

    

    function onMountButton(obj: OneDriveViewerButton) {
        switch (obj.props.name) {
            case "buttonSelect":
                buttonSelect = obj;
                break;
            case "buttonOpen":
                buttonOpen = obj;
                break;
            case "buttonDownload":
                buttonDownload = obj;
                break;
        }
    }

    function onMountFileSelector(obj: OneDriveViewerSelector) {
        selector = obj;
    }

    function onMountPreview(obj: OneDrivePreviewFrame) {
        preView = obj;
    }

    //load item
    setTimeout(loadItem, 500);
  
    
    return (
        <div id="ctrl">
            <table style={fraButtonsCss}>
                <tbody>
                    <tr>
                        <td>
                            <OneDriveViewerButton title="Select file" onClick={() => onButtonSelectFile()}
                                name="buttonSelect"
                                onMount={function (obj: OneDriveViewerButton) { onMountButton(obj) }}
                                 />
                        </td>
                        <td>
                            <OneDriveViewerButton title="Open in OneDrive" onClick={() => onButtonOpenFile()}
                                name="buttonOpen"
                                onMount={function (obj: OneDriveViewerButton) { onMountButton(obj) }}
                                />
                        </td>
                        <td>
                            <OneDriveViewerButton title="Download file" onClick={() => onButtonDownloadFile()}
                                name="buttonDownload"
                                onMount={function (obj: OneDriveViewerButton) { onMountButton(obj) }}
                                />
                        </td>
                    </tr>
                </tbody>
            </table>
            <br />
            
            <OneDriveViewerSelector
                onMount={function (obj: OneDriveViewerSelector) { onMountFileSelector(obj) }}
                    accesToken={fileViewProps.accessToken}
                    //driveId={fileViewProps.driveId}
                
                    endpoint={fileViewProps.endpoint + fileViewProps.path}
                    onItemSelected={(id: string) => onItemSelected(id)} 
                    onGetAccessToken={() => { return onGetAccessToken() } }
                    onCancel={() => onCancelFileSelected()}
                />

            <OneDrivePreviewFrame
                onMount={function (obj: OneDrivePreviewFrame) { onMountPreview(obj) }}
            />

        </div>
    );


}
