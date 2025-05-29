import * as React from "react";
import * as ReactDOM from "react-dom";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ActivityDatasetControl, IActivityDatasetControlProps } from "./Components/EmailDatasetControl";

export class EmailPCFControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _container: HTMLDivElement;
  private _context: ComponentFramework.Context<IInputs>;
  private _notifyOutputChanged: () => void;
  private _controlProps: IActivityDatasetControlProps;

  constructor() {}

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._container = container;
    this._context = context;
    this._notifyOutputChanged = notifyOutputChanged;

    this._controlProps = {
      context: this._context,
      notifyOutputChanged: this._notifyOutputChanged,
    };

    ReactDOM.render(React.createElement(ActivityDatasetControl, this._controlProps), this._container);
    console.log("Control initialized");
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
    this._controlProps.context = this._context;
    ReactDOM.render(React.createElement(ActivityDatasetControl, this._controlProps), this._container);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    ReactDOM.unmountComponentAtNode(this._container);
  }
}