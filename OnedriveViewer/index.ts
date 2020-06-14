import {IInputs, IOutputs} from "./generated/ManifestTypes";
import { IAquireTokenProperties, MSALConnection } from "./msal";
import { IOnedriveViewerAppProperties, OneDriveViewerApp   } from "./OneDriveViewerApp";
import ReactDOM = require("react-dom");
import React = require("react");

export class OnedriveViewer implements ComponentFramework.StandardControl<IInputs, IOutputs> {


	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	private _notifyOutputChanged: () => void;
	public _fileId: string;
	
	private fileViewProps: IOnedriveViewerAppProperties = {
		context: this._context,
		path : "",
		fileId: "",	
		node: this._container,
		onFileSelected: (previewUrl: string, webUrl: string, id: string) => { },
		endpoint: "",
		accessToken: "",
		enableSelect: false,
	};

	private aquireTokenProperties: IAquireTokenProperties =
		{
			authority: "",
			validateAuthority: false,
			clientId: "",
			redirectUri: "",
			scopes: ["user.read", "Files.Read.All"],

		};


	/**
	 * Empty constructor.
	 */
	constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		console.log('init called');

		// Add control initialization code
		this._container = container;
		this._context = context;
		this._notifyOutputChanged = notifyOutputChanged;

	}


	private renderComponent(): void {

		this.initParameters();

		const auth: MSALConnection = new MSALConnection();
		auth.AquireToken(this.aquireTokenProperties).then(token => {
			this.fileViewProps.accessToken = token;
			if (this.fileViewProps.accessToken !== '' && this.fileViewProps.accessToken !== null) {
				ReactDOM.render(
						React.createElement(
							OneDriveViewerApp,
							this.fileViewProps
						),
						this._container
				);				
			}
			else {
				//try again
				auth.AquireToken(this.aquireTokenProperties).then(token => {
					this.fileViewProps.accessToken = token;
					if (this.fileViewProps.accessToken !== '' && this.fileViewProps.accessToken !== null) {
						ReactDOM.render(
							React.createElement(
								OneDriveViewerApp,
								this.fileViewProps
							),
							this._container
						);
					}
					else {
						const notifyLabelElement = document.createElement("label")
						notifyLabelElement.innerText = "Invalid authentication. Please refresh and re-authenticate";
						this._container.appendChild(notifyLabelElement);
					}
				});
			}
		});
	}



	private initParameters(): void {
		this.fileViewProps = {
			context: this._context,
			node: this._container,
			enableSelect: this._context.parameters.EnableSelect.raw === 1 ? true : false,
			endpoint: this._context.parameters.Endpoint.raw as string,
			accessToken: '',
			path: this._context.parameters.Path.raw as string,
			fileId: this._context.parameters.FileId.raw === null ? "" : this._context.parameters.FileId.raw as string,
			onFileSelected: (previewUrl: string, webUrl: string, id: string) => {
				if (this._fileId !== id) {
					this._fileId = id;
					this._notifyOutputChanged();
				}
			},
		};

		
		if (this.fileViewProps.path === "val")
			this.fileViewProps.path = "";
		if (this.fileViewProps.fileId === "val")
			this.fileViewProps.fileId = "";
		if (this.fileViewProps.endpoint === "val")
			this.fileViewProps.endpoint = "";


		this.aquireTokenProperties =
		{
			authority: this._context.parameters.Authority.raw === null ? "" : this._context.parameters.Authority.raw as string,
			clientId: this._context.parameters.ClientId.raw as string,
			redirectUri: this._context.parameters.RedirectUri.raw === null ? "" : this._context.parameters.RedirectUri.raw as string,
			scopes: ["user.read", "Files.Read.All"],
			validateAuthority: true 
		}
	}


	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
		
		console.log('updateview called');
		this._context = context;

		this.renderComponent();
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs
	{
		return {
			FileId: this._fileId
		};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
		ReactDOM.unmountComponentAtNode(this._container);
	}
}