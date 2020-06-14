/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    FileId: ComponentFramework.PropertyTypes.StringProperty;
    FrameHeight: ComponentFramework.PropertyTypes.StringProperty;
    EnableSelect: ComponentFramework.PropertyTypes.WholeNumberProperty;
    ClientId: ComponentFramework.PropertyTypes.StringProperty;
    Authority: ComponentFramework.PropertyTypes.StringProperty;
    RedirectUri: ComponentFramework.PropertyTypes.StringProperty;
    Endpoint: ComponentFramework.PropertyTypes.StringProperty;
    Path: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    FileId?: string;
}
