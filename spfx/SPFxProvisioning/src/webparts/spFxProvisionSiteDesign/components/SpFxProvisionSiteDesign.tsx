import * as React from 'react';
import styles from './SpFxProvisionSiteDesign.module.scss';
import { ISpFxProvisionSiteDesignProps, ISpFxProvisionSiteDesignState } from '.';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  autobind,
  Dropdown,
  IDropdownOption,
  TextField,
  MessageBar,
  MessageBarType,
  Label,
  PrimaryButton,
  Spinner
} from 'office-ui-fabric-react';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';


export default class SpFxProvisionSiteDesign extends React.Component<ISpFxProvisionSiteDesignProps, ISpFxProvisionSiteDesignState> {

  constructor(props: ISpFxProvisionSiteDesignProps) {
    super(props);

    this.state = {
      selectedItem: null,
      siteDesigns: [],
      groupAliasAvailable: null,
      siteUrl: "",
      emailDisabled: true,
      siteName: "",
      isLoading: false,
      siteCreationResult: null,
    };

  }

  public async componentDidMount(): Promise<void> {
    // Get all site designs and add to state
    // const siteDesigns = await this._getSiteDesigns();
    // this.setState({ siteDesigns: siteDesigns });
    this._loadSiteDesigns();
  }


  public render(): React.ReactElement<ISpFxProvisionSiteDesignProps> {
    const { siteCreationResult, isLoading, siteDesigns, siteUrl, selectedItem, groupAliasAvailable, emailDisabled, siteName } = this.state;

    return (
      <div>
        <h1>{"Create a modern site with a site design"}</h1>

        {/* Show spinner while creating site */}
        {
          isLoading &&
          <Spinner>{"Creating Site"}</Spinner>
        }

        {/* Show message after site creation */}
        {
          siteCreationResult &&
          <MessageBar messageBarType={MessageBarType.success}>
            {`Successfully created the site: ${siteCreationResult.SiteUrl}`}
          </MessageBar>
        }

        {/* Populate dropdown with sitedesigns, and add webtemplate property as well */}
        <Dropdown
          label={"Select site template"}
          options={siteDesigns ? siteDesigns.map(s => ({ webTemplate: s.WebTemplate, key: s.Id, text: s.Title })) : []}
          selectedKey={selectedItem ? selectedItem.key : null}
          onChanged={this._dropDownSelected}

        />
        <br />

        {/* Don't show any these if we havn't selected any site design */}
        {
          selectedItem &&
          <div>
            <TextField
              label="Site name"
              value={siteName}
              resizable={false}
              onChanged={this._textChanged}
              disabled={!selectedItem}
            />

            {/* Show only if we're creating a Team Site */}
            {selectedItem.webTemplate === "64" &&
              <div>

                {/* Show only if team site and if the alias is available  */}
                {groupAliasAvailable && siteUrl && siteName.length > 0 &&
                  <MessageBar
                    messageBarType={MessageBarType.success}
                  >
                    {"The site name is available"}
                  </MessageBar>
                }
                <br />

                <TextField
                  label="Group Email Adress"
                  value={siteName}
                  resizable={false}
                  disabled={emailDisabled}

                />
                {groupAliasAvailable != null && siteName.length > 0 &&
                  <MessageBar
                    messageBarType={groupAliasAvailable ? MessageBarType.success : MessageBarType.error}
                    isMultiline={true}>
                    {groupAliasAvailable ? "Group alias is available" : "Another group with the same alias already exists"}
                  </MessageBar>
                }
              </div>
            }
            <br />

            <Label>{"Site address"}</Label>

            {/* Only show the textfield if communication site  */}
            {selectedItem.webTemplate === "68" &&
              <TextField
                value={siteName}
                resizable={false}
                onChanged={this._textChanged}
                disabled={emailDisabled}
              />
            }
            {/* Always show the site url we get back */}
            {groupAliasAvailable != false && siteName.length > 0 &&
              <MessageBar
                messageBarType={MessageBarType.success}
              >{siteUrl}
              </MessageBar>
            }
            <br />
            {/* Only show if we should be able to create a site */}
            <PrimaryButton disabled={!siteName || !siteUrl || groupAliasAvailable == false} onClick={this._createSite}>Create site</PrimaryButton>
          </div>
        }
      </div >
    );
  }

  // When selecting a site design, reset the inputs
  @autobind
  private _dropDownSelected(option: IDropdownOption) {
    this.setState({
      selectedItem: option,
      siteName: "",
      siteUrl: "",
      groupAliasAvailable: null,
    });
  }

  // When the text is changed in the site name input
  @autobind
  private async _textChanged(text: string) {

    // Reset
    this.setState({
      siteName: text,
      siteUrl: "",
      groupAliasAvailable: null,
    });

    if (text.length > 0) {
      // If modern team site
      if (this.state.selectedItem.webTemplate === "64") {

        // Validate group name from the SharePoint api. Returns true or false
        const isValidName = await this._validateGroupName(text);

        this.setState({
          groupAliasAvailable: isValidName
        });

        // If valid alias
        if (isValidName) {

          // Get the valid site url
          const siteUrl = await this._getValidSiteUrl(text);

          this.setState({
            siteUrl: siteUrl
          });
        }

      }
      // If communication site 
      else if (this.state.selectedItem.webTemplate === "68") {

        // Just get the valid site url
        const siteUrl = await this._getValidSiteUrl(text);

        this.setState({
          siteUrl: siteUrl
        });
      }
    }
  }

  // Create the site
  @autobind
  private async _createSite() {
    const { selectedItem, siteUrl, siteName } = this.state;

    // Activate spinner
    this.setState({
      isLoading: true,
    });

    let siteCreated;

    if (selectedItem.webTemplate === "64") {
      // Create the team site with the specific site design
      siteCreated = await this.createTeamSite(siteName, siteName, true, 1033, "", "", null, selectedItem.key);
    }
    else if (selectedItem.webTemplate === "68") {
      // Create communication site with the specific site design
      siteCreated = await this.createCommunicationSite(siteName, 1033, siteUrl, selectedItem.key);
    }

    // disable spinner and show the messagebar with the creation result
    this.setState({
      isLoading: false,
      siteCreationResult: siteCreated
    });

    // wait three seconds and change to site url
    setTimeout(() => window.location.href = siteCreated.SiteUrl, 3000);

  }

  // Get all available site designs
  private async _loadSiteDesigns(): Promise<any> {

    return this.props.client.post(
      `https://simonmvp.sharepoint.com/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
      SPHttpClient.configurations.v1, null)
      .then((response: SPHttpClientResponse) => {
        debugger;
        return response.json();
      }).then(result => this.setState({
        siteDesigns: result.value
      }));

  }

  // Validate the alias to make sure the group does not exist
  private async _validateGroupName(groupName: string): Promise<boolean> {
    return this.props.client.get(
      `https://simonmvp.sharepoint.com/_api/SP.Directory.DirectorySession/ValidateGroupName(displayName='${groupName}',%20alias='${groupName}')`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then(result => {
        return result.IsValidName;
      });
  }

  // Get a valid site url. Returns suggestion if it already exists. Like in the UI
  private async _getValidSiteUrl(siteName: string): Promise<string> {
    return this.props.client.get(
      `https://simonmvp.sharepoint.com/_api/GroupSiteManager/GetValidSiteUrlFromAlias?alias='${siteName}'&isTeamSite=true`,
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then(result => {
        return result.value;
      });
  }

  // Create modern team time
  public async createTeamSite(displayName: string, alias: string, isPublic = true, lcid = 1033, description = "",
    classification = "", owners?: string[], siteDesignId?: string): Promise<any> {

    const creationOptions = [`SPSiteLanguage:${lcid}`];
    if (siteDesignId) {
      creationOptions.push(`implicit_formula_292aa8a00786498a87a5ca52d9f4214a_${siteDesignId}`);
    }

    const postBody = {
      alias: alias,
      displayName: displayName,
      isPublic: isPublic,
      optionalParams: {
        Classification: classification,
        CreationOptions: creationOptions,
        Description: description,
        Owners: owners ? owners : [],
      }
    };

    const opt: ISPHttpClientOptions = {};
    opt.headers = {
      'Accept': 'application/json;odata.metadata=minimal',
    };
    opt.body = JSON.stringify(postBody);

    return this.props.client.post(`https://simonmvp.sharepoint.com/_api/GroupSiteManager/CreateGroupEx`, SPHttpClient.configurations.v1, opt)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then(result => {
        return result;
      });

  }

  // Create communication site
  public createCommunicationSite(title: string, lcid: number, url: string, webTemplateExtensionId?: string, description?: string, classification?: string, shareByEmailEnabled?: boolean,
    siteDesignId?: string): Promise<any> {

    const postBody =
    {
      "request": {
        WebTemplate: "SITEPAGEPUBLISHING#0",
        Title: title,
        Url: url,
        Description: description ? description : "",
        Classification: classification ? classification : "",
        SiteDesignId: siteDesignId ? siteDesignId : "00000000-0000-0000-0000-000000000000",
        Lcid: lcid ? lcid : 1033,
        ShareByEmailEnabled: shareByEmailEnabled ? shareByEmailEnabled : false,
        WebTemplateExtensionId: webTemplateExtensionId ? webTemplateExtensionId : "00000000-0000-0000-0000-000000000000",
        HubSiteId: "00000000-0000-0000-0000-000000000000"
      }
    };

    const opt: ISPHttpClientOptions = {};
    opt.headers = { "Accept": "application/json;odata.metadata=minimal" };
    opt.body = JSON.stringify(postBody);

    return this.props.client.post(`https://simonmvp.sharepoint.com/_api/SPSiteManager/Create`, SPHttpClient.configurations.v1, opt)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then(result => {
        return result;
      });
  }

}