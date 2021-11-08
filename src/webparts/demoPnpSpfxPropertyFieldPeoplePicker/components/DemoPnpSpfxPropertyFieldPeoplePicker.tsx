import * as React from 'react';
import styles from './DemoPnpSpfxPropertyFieldPeoplePicker.module.scss';
import { IDemoPnpSpfxPropertyFieldPeoplePickerProps } from './IDemoPnpSpfxPropertyFieldPeoplePickerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class DemoPnpSpfxPropertyFieldPeoplePicker extends React.Component<IDemoPnpSpfxPropertyFieldPeoplePickerProps, {}> {
  public render(): React.ReactElement<IDemoPnpSpfxPropertyFieldPeoplePickerProps> {
    return (
      <div className={ styles.demoPnpSpfxPropertyFieldPeoplePicker }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>

              
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <p className={ styles.description }>{JSON.stringify(this.props.people)}</p>

              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
