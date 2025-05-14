import * as React from "react";
import * as ReactDom from "react-dom";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IEmployeeWebPartProps } from "./components/IEmployeeWebPartProps";
import EmployeeList from "./components/EmployeeList";

export interface IEmployeeDetailsWebPartProps {
  description: string;
}

export default class EmployeeDetailsWebPart extends BaseClientSideWebPart<IEmployeeWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IEmployeeWebPartProps> =
      React.createElement(EmployeeList, {
        description: this.properties.description,
        context: this.context,
      });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}
