import * as React from "react";
import { Fragment } from "react";
import styles from "./HelloWorld.module.scss";
import { IHelloWorldProps } from "./IHelloWorldProps";
import UserCard from "./UserCard";
import { Store, StoreProvider } from "../context/user-context";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import UserForm from "./UserForm";
import { ClientMode } from "./ClientMode";
import { IHelloWorldState } from "./IHelloWorldState";
import { IUserItem } from "./IUserItem";
import {
  Dropdown,
  DropdownMenuItemType,
  IDropdownOption,
  IDropdownProps,
} from "@fluentui/react/lib/Dropdown";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
} from "office-ui-fabric-react/lib/DetailsList";
import { TextField } from "@fluentui/react/lib/TextField";
import { Button, BaseButton, PrimaryButton } from "office-ui-fabric-react/lib";
import * as strings from "HelloWorldWebPartStrings";
import { MSGraphClientV3, AadHttpClient } from "@microsoft/sp-http";
import TableList from "./TableList";
import EditModal from "./EditModal";

// import { escape } from "@microsoft/sp-lodash-subset";

interface HelloWorldState {
  users: any[];
  searchFor: string;
}

class HelloWorld extends React.Component<
  IHelloWorldProps,
  HelloWorldState,
  {}
> {
  static contextType: any = Store;

  constructor(props: any, state: IHelloWorldState) {
    super(props);

    // Initialize the initial state
    this.state = {
      users: [],
      searchFor: "",
    };
  }

  // Graph
  // private _onSearchForChanged = (
  //   event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
  //   newValue?: string
  // ): void => {
  //   // Update the component state accordingly to the current user's input
  //   this.setState({
  //     searchFor: newValue,
  //   });
  // };

  // private _getSearchForErrorMessage = (value: string): string => {
  //   // The search for text cannot contain spaces
  //   return value == null || value.length == 0 || value.indexOf(" ") < 0
  //     ? ""
  //     : `${strings.SearchForValidationErrorMessage}`;
  // };

  // private _search = (
  //   event: React.MouseEvent<
  //     | HTMLAnchorElement
  //     | HTMLButtonElement
  //     | HTMLDivElement
  //     | BaseButton
  //     | Button,
  //     MouseEvent
  //   >
  // ): void => {
  //   console.log(this.props.clientMode);

  //   // Based on the clientMode value search users
  //   switch (this.props.clientMode) {
  //     // case ClientMode.aad:
  //     //   this._searchWithAad();
  //     //   break;
  //     case ClientMode.graph:
  //       this._searchWithGraph();
  //       break;
  //   }
  // };

  private _searchWithGraph = (): void => {
    // Log the current operation

    this.props.ctx.msGraphClientFactory
      .getClient("3")
      .then((client: MSGraphClientV3) => {
        // From https://github.com/microsoftgraph/msgraph-sdk-javascript sample
        client
          .api("/sites")
          .version("v1.0")
          // .select("displayName,mail,userPrincipalName")
          // .filter(
          //   `(givenName eq '${escape(
          //     this.state.searchFor
          //   )}') or (surname eq '${escape(
          //     this.state.searchFor
          //   )}') or (displayName eq '${escape(this.state.searchFor)}')`
          // )
          .get((err: any, res: any) => {
            if (err) {
              console.log(err);
              return;
            }

            console.log("VALUE: ", res);

            // Prepare the output array
            let users: Array<IUserItem> = new Array<IUserItem>();

            // Map the JSON response to the output array
            res.value.map((item: any) => {
              users.push({
                displayName: item.displayName,
                mail: item.mail,
                userPrincipalName: item.userPrincipalName,
              });
            });

            // Update the component state accordingly to the result
            this.setState({
              users: users,
            });
          })
          .then((obj) => console.log(obj))
          .catch((err) => console.log(err));
      });
  };

  private _searchWithAad = async (
    link: string,
    saveParam: string,
    filterFunction?: Function,
    method: "get" | "post" = "get",
    options?: any
  ): Promise<any> => {
    try {
      const client: AadHttpClient =
        await this.props.ctx.aadHttpClientFactory.getClient(
          "https://graph.microsoft.com"
        );

      const response = await client[method](
        link,
        AadHttpClient.configurations.v1,
        options
      );
      const json = await response.json();

      let filteredData: Array<any> = [];

      console.log(json);

      if (filterFunction) {
        filteredData = filterFunction(json.value);
      } else {
        filteredData = json.value;
      }

      if (saveParam) {
        this.context.dispatch({ type: saveParam, payload: filteredData });
      }

      return json;
    } catch (error) {
      console.error(error);
    }
  };

  // private async _searchWithSP(
  //   link: string,
  //   saveParam: string,
  //   filterFunction?: Function,
  //   method: "get" | "post" = "get",
  //   options?: any
  // ): Promise<any[]> {
  //   console.log(method);
  //   const spHttpClient: SPHttpClient = this.props.ctx.spHttpClient;
  //   const response: SPHttpClientResponse = await spHttpClient[method](
  //     link,
  //     SPHttpClient.configurations.v1,
  //     options
  //   );

  //   if (response.ok) {
  //     const data = await response.json();
  //     let filteredData: Array<any> = new Array<any>();

  //     if (filterFunction) {
  //       filteredData = filterFunction(data?.value || data);
  //     } else {
  //       filteredData = data?.value || data;
  //     }

  //     console.log("SP Value: ", filteredData);

  //     if (saveParam)
  //       this.context.dispatch({ type: saveParam, payload: filteredData });

  //     return filteredData;
  //   } else {
  //     throw new Error(`Error fetching data from list: ${response.statusText}`);
  //   }
  // }

  private async _searchWithSP(
    link: string,
    saveParam: string,
    filterFunction?: Function,
    method: "get" | "post" = "get",
    options?: any
  ): Promise<any[]> {
    try {
      console.log(method);
      const spHttpClient: SPHttpClient = this.props.ctx.spHttpClient;
      const response: SPHttpClientResponse = await spHttpClient[method](
        link,
        SPHttpClient.configurations.v1,
        options
      );

      if (response.ok) {
        const data = await response.json();
        let filteredData: Array<any> = [];

        if (filterFunction) {
          filteredData = filterFunction(data?.value || data);
        } else {
          filteredData = data?.value || data;
        }

        console.log("SP Value: ", filteredData);

        if (saveParam) {
          this.context.dispatch({ type: saveParam, payload: filteredData });
        }

        return filteredData;
      } else {
        throw new Error(
          `Error fetching data from list: ${response.statusText}`
        );
      }
    } catch (error) {
      console.error(error);
    }
  }

  // private async fetchUsersFromList(): Promise<any[]> {
  //   const apiUrl =
  //     "https://sxmyf.sharepoint.com/sites/ArkitektzTest/_api/web/lists('e3dd271b-6d15-4d07-9c5d-520cb4c4fff4')/items";
  //   // "https://graph.microsoft.com/v1.0/sites/ArkitektzTest/lists/e3dd271b-6d15-4d07-9c5d-520cb4c4fff4/items";

  //   const spHttpClient: SPHttpClient = this.props.ctx.spHttpClient;
  //   const response: SPHttpClientResponse = await spHttpClient.get(
  //     apiUrl,
  //     SPHttpClient.configurations.v1
  //   );

  //   if (response.ok) {
  //     const data = await response.json();
  //     this.setState({ users: data.value });
  //     return data.value;
  //   } else {
  //     throw new Error(`Error fetching data from list: ${response.statusText}`);
  //   }
  // }

  // private createUserInList(formData: any): any {
  //   const apiUrl =
  //     "https://sxmyf.sharepoint.com/sites/ArkitektzTest/_api/web/lists('e3dd271b-6d15-4d07-9c5d-520cb4c4fff4')/items";
  //   const body = JSON.stringify({ Title: formData.name, Role: formData.role });

  //   if (formData.role && formData.name) {
  //     const spHttpClient: SPHttpClient = this.props.ctx.spHttpClient;
  //     spHttpClient
  //       .post(apiUrl, SPHttpClient.configurations.v1, { body })
  //       .then(() => this.fetchUsersFromList())
  //       .catch((err) => console.log(err));
  //   }
  // }

  // private deleteUserFromList(id: string, etag: string): any {
  //   const apiUrl = `https://sxmyf.sharepoint.com/sites/ArkitektzTest/_api/web/lists('e3dd271b-6d15-4d07-9c5d-520cb4c4fff4')/items(${id})`;

  //   if (id) {
  //     console.log(etag);
  //     const spHttpClient: SPHttpClient = this.props.ctx.spHttpClient;
  //     spHttpClient
  //       .post(apiUrl, SPHttpClient.configurations.v1, {
  //         headers: { "X-HTTP-Method": "DELETE", "If-Match": etag },
  //       })
  //       .then(() => this.fetchUsersFromList())
  //       .catch((err) => console.log(err));
  //   }
  // }

  componentDidMount(): void {
    // this.fetchUsersFromList()
    //   .then(() => null)
    //   .catch((err) => console.log(err));

    // const currentUserLoginName = this.props.ctx.pageContext.user.loginName;

    // fetch(
    //   // "https://sxmyf.sharepoint.com/_api/site/rootweb?$select=Title,ServerRelativeUrl",
    //   // `https://sxmyf.sharepoint.com/_api/search/query?querytext='contentclass:STS_Site'`,
    //   // `https://sxmyf.sharepoint.com/_api/search/query?querytext='contentclass:STS_Site Path:"https://sxmyf.sharepoint.com/*"'`,
    //   "https://sxmyf.sharepoint.com/_api/search/query?querytext='contentclass:STS_Site OR contentclass:STS_Web'&select='Title,SPWebId",
    //   {
    //     headers: {
    //       Accept: "application/json;odata=minimalmetadata;charset=utf-8",
    //     },
    //   }
    // )
    //   .then((data) => data.json())
    //   .then((respData) => console.log(respData))
    //   .catch((err) => console.log(err));

    this._searchWithAad(
      "https://graph.microsoft.com/v1.0/sites?search=",
      "SET_SITES"
    );
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    // const {
    // hasTeamsContext,
    // description,
    // isDarkTheme,
    // environmentMessage,
    // userDisplayName,
    // } = this.props;

    console.log(this.props.ctx.pageContext.web.absoluteUrl);

    if (!this?.context?.state?.sites?.length) return <p>Loading</p>;

    return (
      <Fragment>
        {/* <Dropdown
          placeholder="Select a Site"
          label="Site"
          ariaLabel="Site Selection"
          onRenderOption={(option: any) => option.displayName}
          options={this?.context?.state?.sites}
          onChange={(_, item) => {
            this.context.dispatch({ type: "SET_SELECTED_SITE", payload: item });
            this._searchWithAad(
              `https://graph.microsoft.com/v1.0/sites/${item.id}/lists`,
              "SET_LISTS"
            );
          }}
          selectedKey={this?.context?.state?.selectedSite?.id || undefined}
        />
        <br />
        <Dropdown
          placeholder="Select a List"
          label="List"
          ariaLabel="List Selection"
          onRenderOption={(option: any) => option.displayName}
          options={this?.context?.state?.lists}
          onChange={(_, item) => {
            this.context.dispatch({ type: "SET_SELECTED_LIST", payload: item });
            // this._searchWithAad(
            //   // `https://graph.microsoft.com/v1.0/sites/${this?.context?.state?.selectedSite?.id}/lists/${item.id}/columns`,
            //   `https://sxmyf.sharepoint.com/sites/ArkitektzTest/_api/web/lists('${item.id}')/fields?$filter=Hidden eq false and ReadOnlyField eq false`,
            //   "SET_COLUMNS",
            //   (columns: any) =>
            //     columns.filter((column: any) => !column.readOnly)
            // );

            this._searchWithSP(
              `${this?.context?.state?.selectedSite?.webUrl}/_api/web/lists('${item.id}')/fields?$filter=Hidden eq false and ReadOnlyField eq false and FromBaseType eq false`,
              "SET_COLUMNS"
            );

            this._searchWithAad(
              `https://graph.microsoft.com/v1.0/sites/${this?.context?.state?.selectedSite?.id}/lists/${item.id}/items?$expand=fields`,
              "SET_TABLE_LIST"
            );
          }}
        /> */}

        <TableList
          selectedSite={this.props.selectedSite}
          selectedList={this.props.selectedList}
          request={this._searchWithAad}
        />

        <UserForm
          selectedSite={this.props.selectedSite}
          selectedList={this.props.selectedList}
          request={this._searchWithAad}
          requestSP={(
            link: string,
            param: string,
            func: Function,
            method: "get" | "post",
            options: any
          ) => this._searchWithSP(link, param, func, method, options)}
        />

        <EditModal
          request={this._searchWithAad}
          requestSP={(
            link: string,
            param: string,
            func: Function,
            method: "get" | "post",
            options: any
          ) => this._searchWithSP(link, param, func, method, options)}
        />
      </Fragment>
    );
  }

  //   public render(): React.ReactElement<IHelloWorldProps> {
  //     return (
  //       <div className={styles.helloWorld}>
  //         <div className={styles.container}>
  //           <div className={styles.row}>
  //             <div className={styles.column}>
  //               <span className={styles.title}>Search for a user!</span>
  //               <p className={styles.form}>
  //                 <TextField
  //                   label={strings.SearchFor}
  //                   required={true}
  //                   onChange={this._onSearchForChanged}
  //                   onGetErrorMessage={this._getSearchForErrorMessage}
  //                   value={this.state.searchFor}
  //                 />
  //               </p>
  //               <p className={styles.form}>
  //                 <PrimaryButton
  //                   text="Search"
  //                   title="Search"
  //                   onClick={this._search}
  //                 />
  //               </p>
  //               {this.state.users != null && this.state.users.length > 0 ? (
  //                 <p className={styles.form}>
  //                   <DetailsList
  //                     items={this.state.users}
  //                     columns={_usersListColumns}
  //                     setKey="set"
  //                     // checkboxVisibility={CheckboxVisibility.hidden}
  //                     // selectionMode={SelectionMode.none}
  //                     layoutMode={DetailsListLayoutMode.fixedColumns}
  //                     compact={true}
  //                   />
  //                 </p>
  //               ) : null}
  //             </div>
  //           </div>
  //         </div>
  //       </div>
  //     );
  //   }
}

export default class Wrapper extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    const {
      hasTeamsContext,
      description,
      isDarkTheme,
      environmentMessage,
      userDisplayName,
      ctx,
      clientMode,
      selectedSite,
      selectedList,
    } = this.props;

    return (
      <StoreProvider>
        <HelloWorld
          description={description}
          isDarkTheme={isDarkTheme}
          environmentMessage={environmentMessage}
          hasTeamsContext={hasTeamsContext}
          userDisplayName={userDisplayName}
          clientMode={clientMode}
          ctx={ctx}
          selectedList={selectedList}
          selectedSite={selectedSite}
        />
      </StoreProvider>
    );
  }
}

// <div className={styles.welcome}>
//   <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
//   <h2>Well done, {escape(userDisplayName)}!</h2>
//   <div>{environmentMessage}</div>
//   <div>Web part property value: <strong>{escape(description)}</strong></div>
// </div>
// <div>
//   <h3>Welcome to SharePoint Framework!</h3>
//   <p>
//     The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
//   </p>
//   <h4>Learn more about SPFx development:</h4>
//   <ul className={styles.links}>
//     <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
//     <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
//     <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
//     <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
//     <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
//     <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
//     <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
//   </ul>
// </div>
