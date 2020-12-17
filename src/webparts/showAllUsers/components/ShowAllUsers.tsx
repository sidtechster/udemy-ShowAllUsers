import * as React from 'react';
import styles from './ShowAllUsers.module.scss';
import { IShowAllUsersProps } from './IShowAllUsersProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IUser } from './IUser';
import { IShowAllUsersState } from './IShowAllUsersState';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import {
  TextField,
  autobind,
  PrimaryButton,
  DetailsList,
  DetailsListLayoutMode,
  CheckboxVisibility,
  SelectionMode
} from 'office-ui-fabric-react';

import * as strings from 'ShowAllUsersWebPartStrings';

let _usersListColumn = [
  {
    key: 'displayName',
    name: 'Display Name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true
  },
  {
    key: 'mail',
    name: 'Email',
    fieldName: 'mail',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true
  }
]

export default class ShowAllUsers extends React.Component<IShowAllUsersProps, IShowAllUsersState> {

  constructor(props: IShowAllUsersProps, state: IShowAllUsersState) {
    super(props);

    this.state = {
      users: [],
      searchFor: "Siddhartha"
    };
  }

  public componentDidMount(): void {
    this.fetchUserDetails();
  }

  @autobind
  public _search(): void {
    this.fetchUserDetails();
  }

  @autobind
  private _onSearchForChanged(newValue: string): void {
    this.setState({
      searchFor: newValue
    });
  }

  @autobind
  private _getSearchForErrorMessage(value: string): string {
    return (value == null || value.length == 0 || value.indexOf(" ") < 0)
      ? ''
      : `${strings.SearchForValidationErrorMessage}`;
  }

  public fetchUserDetails(): void {
    this.props.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
      client
        .api('users')
        .version("v1.0")
        .select("*")
        .filter(`startswith(givenname, '${escape(this.state.searchFor)}')`)
        .get((error: any, response, rawResponse?: any) => {
          if(error) {
            console.error("Message is : " + error);
            return;
          }

          //Prepare the output array
          var allUsers: Array<IUser> = new Array<IUser>();

          //Map the json response to the output array
          response.value.map((item: IUser) => {
            allUsers.push({
              displayName: item.displayName,
              mail: item.mail
            });
          });

          this.setState({users: allUsers});
        });
    });
  }

  public render(): React.ReactElement<IShowAllUsersProps> {
    return (
      <div className = { styles.showAllUsers }>
        <TextField 
          label = { strings.SearchFor }
          required = { true }
          value = { this.state.searchFor }
          onChanged = { this._onSearchForChanged }
          onGetErrorMessage = { this._getSearchForErrorMessage }
        />

          <p className = { styles.title }>
            <PrimaryButton 
              text='Search'
              title='Search'
              onClick={this._search}
            />
          </p>

          {
            (this.state.users != null && this.state.users.length > 0) ?
              <p className={styles.row}>
                <DetailsList 
                  items={this.state.users}
                  columns={_usersListColumn}
                  setKey='set'
                  checkboxVisibility={CheckboxVisibility.onHover}
                  selectionMode={SelectionMode.single}
                  layoutMode={DetailsListLayoutMode.fixedColumns}
                  compact={true}
                />
              </p>
              : null
          }

      </div>
    );
  }
}
