import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DemoPnpSpfxPropertyFieldPeoplePickerWebPartStrings';
import DemoPnpSpfxPropertyFieldPeoplePicker from './components/DemoPnpSpfxPropertyFieldPeoplePicker';
import { IDemoPnpSpfxPropertyFieldPeoplePickerProps } from './components/IDemoPnpSpfxPropertyFieldPeoplePickerProps';

// PNP Property Panel
import {
  IPropertyFieldGroupOrPerson,
  PropertyFieldPeoplePicker,
  PrincipalType
} from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker'
// PNP Property Panel


export interface IDemoPnpSpfxPropertyFieldPeoplePickerWebPartProps {
  description: string;
  people: IPropertyFieldGroupOrPerson[]
}

export default class DemoPnpSpfxPropertyFieldPeoplePickerWebPart extends BaseClientSideWebPart<IDemoPnpSpfxPropertyFieldPeoplePickerWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDemoPnpSpfxPropertyFieldPeoplePickerProps> = React.createElement(
      DemoPnpSpfxPropertyFieldPeoplePicker,
      {
        description: this.properties.description,
        people: this.properties.people
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),


// PNP Property Panel
                PropertyFieldPeoplePicker('people',{
                  label: 'People Picker',
                  initialData : this.properties.people,
                  allowDuplicate: false,
                  principalType : [PrincipalType.Users],
                  onPropertyChange : this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties : this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                })
// PNP Property Panel


              ]
            }
          ]
        }
      ]
    };
  }
}
