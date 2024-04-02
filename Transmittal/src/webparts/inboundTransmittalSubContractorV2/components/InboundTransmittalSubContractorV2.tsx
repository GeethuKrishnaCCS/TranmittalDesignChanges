import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, DatePicker, Dialog, DialogFooter, DialogType, Dropdown, IDropdownOption, MessageBar, PrimaryButton, ProgressIndicator, SearchBox } from 'office-ui-fabric-react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { SPFI, IList, Web } from "@pnp/sp/presets/all";
import * as _ from 'lodash';
import Select from 'react-select-plus';
import 'react-select-plus/dist/react-select-plus.css';
import * as moment from 'moment';
import SimpleReactValidator from 'simple-react-validator';
import { MSGraphClient, HttpClient, SPHttpClient, HttpClientConfiguration, HttpClientResponse, ODataVersion, IHttpClientConfiguration, IHttpClientOptions, ISPHttpClientOptions } from '@microsoft/sp-http';
import { InboundTransmittalSubContractorV2Props, InboundTransmittalSubContractorV2State } from './InboundTransmittalSubContractorV2Props';




// For mobile view  


export default class EmecInboundSubContractor extends React.Component<InboundTransmittalSubContractorV2Props, InboundTransmittalSubContractorV2State, {}> {
  private validator: SimpleReactValidator;
  private reqWeb;
  private siteUrl;
  private currentEmail;
  private currentUserTitle;
  private currentUserId;
  private addDocument = [];
  private addExternalDocument = [];
  private today;
  private documentNameExtension;
  private sourceDocumentID;
  private documentIndexID;
  private revisionHistoryUrl;
  private revokeUrl;
  private additionalDocumentId;
  private transmittalID;
  private redirectUrl;
  private dataSaved;
  private status;
  private typeForDelete;
  private postUrl;
  private keyForDelete;
  private myfile;
  private myfileadditional;
  private permissionpostUrl;
  private docIndexId;
  public constructor(props: InboundTransmittalSubContractorV2Props) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      dcc: "",
      dccId: null,
      owner: "",
      recievedDate: null,
      subContractorNumber: "",
      poNumber: "",
      comments: "",
      projectName: "",
      projectNumber: "",
      revisionSettingsArray: [],
      transmittalSettingsArray: [],
      subContractorArray: [],
      subContractorID: null,
      subContractor: "",
      purchaseOrderArray: [],
      multidealer: false,
      transmittalOutlookDocumentArray: [],
      documentIndexArray: [],
      documentIndexID: "",
      documentIndexTitle: "",
      revisionCodingId: null,
      isIncrement: false,
      transmittalOutlookId: "",
      transmittalSettingsId: null,
      ownerId: "",
      showGrid: true,
      gridDocument: [],
      externalDate: null,
      externalComments: "",
      gridExternalDocument: [],
      showExternalGrid: true,
      incrementSequenceNumber: "",
      transmittalSubContractorId: "",
      transmittalNo: "none",
      searchDocuments: [],
      items: [],
      searchDiv: "none",
      documentSelectedDiv: true,
      documentSelect: "",
      noSubContactor: "none",
      inboundTransmittalHeaderId: null,
      sourceDocumentId: null,
      additionalDocumentId: null,
      transmittalInboundID: "",
      transmittalDate: null,
      currentInboundDetailItem: [],
      currentInboundAdditionalItem: [],
      myfile: null,
      cancelConfirmMsg: "none",
      confirmDialog: true,
      statusKey: "",
      validComment: "none",
      validDocument: "none",
      disableOutlook: false,
      validAdditionalComment: "none",
      deleteConfirmMsg: "none",
      tempDocIndexIDForDelete: null,
      submitDisable: false,
      viewDocument: "",
      validDocumentIndex: "none",
      confirmDeleteDialog: true,
      accessDeniedMsgBar: "none",
      access: "none",
      loaderDisplay: "",
      checksend: "none",
      uploadAdditionalDocumentError: "none",
      uploadDocumentError: "none",
      noDcc: "none",
      notransmittal: "none",
      legalId: "",
      poNumberID: "",
      title: ""

    };
    this._queryParamGetting = this._queryParamGetting.bind(this);
    this._bindData = this._bindData.bind(this);
    this._getSubContractor = this._getSubContractor.bind(this);
    this._getProjectInformation = this._getProjectInformation.bind(this);
    this._getRevisionSettings = this._getRevisionSettings.bind(this);
    this._getTransmittalSettings = this._getTransmittalSettings.bind(this);
    this._subContactorChanged = this._subContactorChanged.bind(this);
    this._getDocumentIndex = this._getDocumentIndex.bind(this);
    this._documentIndexChange = this._documentIndexChange.bind(this);
    this._outlookDocumentChange = this._outlookDocumentChange.bind(this);
    this._dccChange = this._dccChange.bind(this);
    this._ownerChange = this._ownerChange.bind(this);
    this._onRecievedDatePickerChange = this._onRecievedDatePickerChange.bind(this);
    this._subContractorNumberChange = this._subContractorNumberChange.bind(this);
    this._poNumberChange = this._poNumberChange.bind(this);
    this._onIncrementRevisionChecked = this._onIncrementRevisionChecked.bind(this);
    this._commentschange = this._commentschange.bind(this);
    this._addindex = this._addindex.bind(this);
    this._onTransmittalSettingsChange = this._onTransmittalSettingsChange.bind(this);
    this._onRevisionCodingChange = this._onRevisionCodingChange.bind(this);
    this._onDatePickerChange = this._onDatePickerChange.bind(this);
    this._externalCommentsChange = this._externalCommentsChange.bind(this);
    this._addexternalindex = this._addexternalindex.bind(this);
    this._saveAsDraft = this._saveAsDraft.bind(this);
    this._idGeneration = this._idGeneration.bind(this);
    this._transmittalSequenceNumber = this._transmittalSequenceNumber.bind(this);
    this._addDocument = this._addDocument.bind(this);
    this._bindInboundTransmittalSavedData = this._bindInboundTransmittalSavedData.bind(this);
    this._submit = this._submit.bind(this);
    this._updateall = this._updateall.bind(this);
    this._userMessageSettings = this._userMessageSettings.bind(this);
    this._upload = this._upload.bind(this);
    this._onDocumentIndexFilter = this._onDocumentIndexFilter.bind(this);
    this.checkPermission = this.checkPermission.bind(this);

  }
  public async componentDidMount() {
    this.siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    this.reqWeb = Web(window.location.protocol + "//" + window.location.hostname + this.props.hubUrl);
    //Get Current User
    const user = await sp.web.currentUser.get();
    this.currentEmail = user.Email;
    this.currentUserTitle = user.Title;
    this.currentUserId = user.Id;
    // await this.checkPermission();
    await this._queryParamGetting();
    let getdccreviewer = [];
    getdccreviewer.push(this.currentUserTitle);
    this.setState({
      dcc: getdccreviewer[0],
      dccId: this.currentUserId
    });
    let today = new Date();
    this.today = today;
    this.setState({
      recievedDate: today,
      externalDate: today
    });
    this._bindData();
    this._userMessageSettings();
  }
  // Validation
  public componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please select valid Document"
      }
    });

  }
  // Check query parameter
  private async _queryParamGetting() {
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    let transmittalID = params.get('trid');
    console.log("transmittalID", transmittalID);
    //console.log(this.detailID);
    if (transmittalID != "" && transmittalID != null) {
      // alert(transmittalID);
      this.transmittalID = transmittalID;
      const inboundHeader = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalHeaderList).items.select("TransmittalStatus").getById(parseInt(transmittalID)).get();
      console.log(inboundHeader);
      if (inboundHeader.TransmittalStatus != "Completed") {
        this._bindInboundTransmittalSavedData(this.transmittalID);
        this.setState({
          access: "",
          loaderDisplay: "none"
        });
      }
      else {
        this.setState({
          access: "none",
          accessDeniedMsgBar: "",
          statusMessage: { isShowMessage: true, message: "Transmittal is already completed", messageType: 1 },
          loaderDisplay: "none",

        });
        setTimeout(() => {
          this.setState({ accessDeniedMsgBar: 'none', });
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
        }, 10000);
      }
    }

    else {

      this.setState({
        transmittalNo: "none",
        access: "",
        loaderDisplay: "none"
      });
    }
  }
  // Check permission
  public async checkPermission() {
    const laUrl = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.requestList).items.filter("Title eq 'EMEC_PermissionWebpart'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.permissionpostUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.permissionpostUrl;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': 'Project_SendTransmittal',
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.currentEmail

    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
    if (response.ok) {
      console.log(responseJSON['Status']);
      if (responseJSON['Status'] == "Valid") {
        this._queryParamGetting();
        this.setState({
          access: "",
          loaderDisplay: "none"
        });
      }
      else {
        this.setState({
          access: "none",
          loaderDisplay: "none",
          accessDeniedMsgBar: "",
          statusMessage: { isShowMessage: true, message: "You are not permitted to perform this operation", messageType: 1 },
        });
        setTimeout(() => {
          this.setState({ accessDeniedMsgBar: 'none', });
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
        }, 10000);
      }
    }
    else { }
  }
  // Bind data
  public async _bindData() {

    this._getProjectInformation();
    this._getSubContractor();
    this._getRevisionSettings();
    this._getTransmittalSettings();
    this._getDocumentIndex();
  }
  // Bind data with query parameter
  public async _bindInboundTransmittalSavedData(transmittalID) {
    this._getSubContractor();
    let transmittalOutlookDocumentArray = [];
    let sorted_transmittalOutlookDocumentArray = [];
    let transmittalDate;
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalHeaderList).items.select("Id,Title,TransmittalDate,DocumentController/ID,DocumentController/Title,SubContractorID,SubContractor").expand("DocumentController").get().then(async inboundHeader => {
      console.log(inboundHeader);
      for (let l = 0; l < inboundHeader.length; l++) {
        if (inboundHeader[l].Id == transmittalID) {
          transmittalDate = moment(inboundHeader[l].TransmittalDate).format("DD/MM/YYYY"),
            this.setState({
              dcc: inboundHeader[l].DocumentController.Title,
              dccId: inboundHeader[l].DocumentController.ID,
              transmittalInboundID: inboundHeader[l].Title,
              transmittalDate: transmittalDate,
              subContractorID: inboundHeader[l].SubContractorID,
              subContractor: inboundHeader[l].SubContractor,
            });
        }
      }
      const document = await sp.web.getList(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary).items.filter("From eq 'Sub-Contractor'").select("ID,BaseName,SubContractor").get();
      for (let i = 0; i < document.length; i++) {
        if (document[i].SubContractor == this.state.subContractor) {
          let transmittalOutlookDocument = {
            key: document[i].ID,
            text: document[i].BaseName
          };
          transmittalOutlookDocumentArray.push(transmittalOutlookDocument);
        }
      }
      sorted_transmittalOutlookDocumentArray = _.orderBy(transmittalOutlookDocumentArray, 'text', ['asc']);
      this.setState({
        transmittalOutlookDocumentArray: sorted_transmittalOutlookDocumentArray
      });
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalDetailList).items.select("TransmittalHeaderId,DocumentIndex/ID,DocumentIndex/Title,DocumentIndex/DocumentName,Owner/ID,Owner/Title,SubContractorDocumentNumber,ReceivedDate,Comments,ID").expand("DocumentIndex,Owner").filter("TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ").get().then(inboundTransmittalDetailList => {
        console.log("inboundTransmittalDetailList", inboundTransmittalDetailList);
        if (inboundTransmittalDetailList.length > 0) {
          for (var k = 0; k <= inboundTransmittalDetailList.length; k++) {
            if (inboundTransmittalDetailList[k].TransmittalHeaderId == this.transmittalID) {
              // alert(inboundTransmittalDetailList[k].DocumentIndex.ID);
              this.addDocument.push({
                DocumentIndexId: inboundTransmittalDetailList[k].DocumentIndex.ID,
                DocumentIndex: inboundTransmittalDetailList[k].DocumentIndex.DocumentName,
                OwnerId: inboundTransmittalDetailList[k].Owner.ID,
                Owner: inboundTransmittalDetailList[k].Owner.Title,
                SubContractorNumber: inboundTransmittalDetailList[k].SubContractorDocumentNumber,
                ReceivedDate: moment(inboundTransmittalDetailList[k].ReceivedDate).format("DD/MM/YYYY"),
                RecieveDate: inboundTransmittalDetailList[k].ReceivedDate,
                Comments: inboundTransmittalDetailList[k].Comments,
                DetailId: inboundTransmittalDetailList[k].ID,
                SubContractor: inboundTransmittalDetailList[k].SubContractor,
                transmittalSubContractorId: inboundTransmittalDetailList[k].Title
              });
            }
            this.setState({
              gridDocument: this.addDocument,
              showGrid: false,
              currentInboundDetailItem: inboundTransmittalDetailList
            });

          }
        }
      });
      sp.web.getList(this.props.siteUrl + "/" + this.props.additionalDocumentLibrary).items.filter("TransmittalIDId eq '" + this.transmittalID + "' ").get().then(inboundAdditionalDocumentsListName => {
        console.log("inboundAdditionalDocumentsListName", inboundAdditionalDocumentsListName);
        if (inboundAdditionalDocumentsListName.length > 0) {
          for (var k = 0; k < inboundAdditionalDocumentsListName.length; k++) {
            // alert(inboundAdditionalDocumentsListName[k].TransmittalIDId);
            if (inboundAdditionalDocumentsListName[k].TransmittalIDId == this.transmittalID) {
              this.addExternalDocument.push({
                DocName: inboundAdditionalDocumentsListName[k].Title,
                ExternalDate: moment(inboundAdditionalDocumentsListName[k].ReceivedDate).format("DD/MM/YYYY"),
                Comments: inboundAdditionalDocumentsListName[k].Comments,
                additionalId: inboundAdditionalDocumentsListName[k].Id
              });
            }
            this.setState({
              gridExternalDocument: this.addExternalDocument,
              showExternalGrid: false,
              currentInboundAdditionalItem: inboundAdditionalDocumentsListName,
            });
          }
        }
      });


    });
  }
  // Get user messages
  private async _userMessageSettings() {
    const userMessageSettings: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.userMessageSettings).items.select("Title,Message").filter("PageName eq 'InboundSib-Contractor'").get();
    console.log(userMessageSettings);
    for (var i in userMessageSettings) {
      if (userMessageSettings[i].Title == "InboundSub-ContractorSave") {
        this.dataSaved = userMessageSettings[i].Message;
      }
    }
  }
  // Get subcontractor
  public async _getSubContractor() {
    await this._getProjectInformation();
    let subContractorarray = [];
    let sorted_SubContractor = [];
    let subContractor;
    // const subContractoritems: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/SubContractorMaster").items.get();


    const subContractoritems: any[] = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/SubContractorMaster").items.filter("ProjectId eq '" + this.state.projectNumber + "'").get();
    for (let i = 0; i < subContractoritems.length; i++) {
      if (subContractoritems[i].ProjectId == this.state.projectNumber) {
        if (subContractoritems[i].Title == this.state.legalId) {
          subContractor = {
            key: subContractoritems[i].VendorId,
            text: subContractoritems[i].VendorName
          };
          subContractorarray.push(subContractor);
        }

      }
    }

    console.log(subContractorarray);
    sorted_SubContractor = _.orderBy(subContractorarray, 'text', ['asc']);
    this.setState({
      subContractorArray: sorted_SubContractor
    });
  }
  // Get project information
  public async _getProjectInformation() {
    const projectInformation = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.projectInformationListName).items.get();
    if (projectInformation.length > 0) {
      for (var k in projectInformation) {
        if (projectInformation[k].Key == "ProjectName") {
          this.setState({
            projectName: projectInformation[k].Title,
          });
        }
        if (projectInformation[k].Key == "ProjectNumber") {
          this.setState({
            projectNumber: projectInformation[k].Title,
          });
        }
        if (projectInformation[k].Key == "LegalEntityId") {
          this.setState({
            legalId: projectInformation[k].Title,
          });
        }
      }
    }
  }
  // get Revision Settings
  public async _getRevisionSettings() {
    let revisionSettingsArray = [];
    let sorted_RevisionSettings = [];
    const revisionSettingsItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.revisionLevelList).items.select("Title,ID").get();
    for (let i = 0; i < revisionSettingsItem.length; i++) {
      let revisionSettingsItemdata = {
        key: revisionSettingsItem[i].ID,
        text: revisionSettingsItem[i].Title
      };
      revisionSettingsArray.push(revisionSettingsItemdata);
    }
    sorted_RevisionSettings = _.orderBy(revisionSettingsArray, 'text', ['asc']);
    this.setState({
      revisionSettingsArray: sorted_RevisionSettings
    });
  }
  // Get Transmittal Settings
  public async _getTransmittalSettings() {
    let transmittalCodeSettingsArray = [];
    let sorted_transmittalCodeSettings = [];
    const transmittalCodeSettingsItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalCodeSettings).items.get();
    for (let i = 0; i < transmittalCodeSettingsItem.length; i++) {
      if (transmittalCodeSettingsItem[i].AcceptanceCode == false) {
        let transmittalCodeSettingsItemdata = {
          key: transmittalCodeSettingsItem[i].ID,
          text: transmittalCodeSettingsItem[i].Title
        };
        transmittalCodeSettingsArray.push(transmittalCodeSettingsItemdata);
      }
    }
    sorted_transmittalCodeSettings = _.orderBy(transmittalCodeSettingsArray, 'text', ['asc']);
    this.setState({
      transmittalSettingsArray: sorted_transmittalCodeSettings
    });
  }
  // Get Document Index
  public async _getDocumentIndex() {
    let documentIndexArray = [];
    let sorted_documentIndexArray = [];
    const documentIndexArrayItem: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getAll(5000);
    console.log(documentIndexArrayItem);
    for (let i = 0; i < documentIndexArrayItem.length; i++) {
      if (documentIndexArrayItem[i].ExternalDocument == true) {
        let documentIndexArrayItemdata = {
          key: documentIndexArrayItem[i].ID,
          text: documentIndexArrayItem[i].DocumentName
        };
        documentIndexArray.push(documentIndexArrayItemdata);

      }
    }
    sorted_documentIndexArray = _.orderBy(documentIndexArray, 'text', ['asc']);
    this.setState({
      documentIndexArray: sorted_documentIndexArray,
      items: sorted_documentIndexArray
    });
  }
  // Filter document index
  private _onDocumentIndexFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {

    if (text == "") {
      this.setState({
        searchDiv: "none",
        documentSelectedDiv: true,
      });
    }
    else {
      this.setState({
        items: text ? this.state.documentIndexArray.filter(i => i.DocumentID.toLowerCase().indexOf(text.toString().toLowerCase()) > -1) : this.state.documentIndexArray,
        searchDiv: "",
      });
    }

  }
  // On sub contractor change
  public async _subContactorChanged(option: { key: any; text: any }) {
    this.setState({ noSubContactor: "none", disableOutlook: false });
    let transmittalOutlookDocumentArray = [];
    let sorted_transmittalOutlookDocumentArray = [];
    this.setState({ subContractorID: option.key, subContractor: option.text });
    let purchasearray = [];
    let sorted_purchaseOrder = [];
    let purchaseitem;
    const purchaseitems = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/PurchaseOrderMaster").items.get();
    console.log(purchaseitems);
    for (let i = 0; i < purchaseitems.length; i++) {
      if (purchaseitems[i].VendorId == option.key) {
        if (purchaseitems[i].Title == this.state.legalId) {
          if (purchaseitems[i].ProjectId == this.state.projectNumber) {
            purchaseitem = {
              key: purchaseitems[i].ID,
              text: purchaseitems[i].PurchaseOrderNo
            };
            purchasearray.push(purchaseitem);
          }
        }

      }
    }
    console.log(purchasearray);
    sorted_purchaseOrder = _.orderBy(purchasearray, 'text', ['asc']);
    this.setState({
      purchaseOrderArray: sorted_purchaseOrder
    });
    const document = await sp.web.getList(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary).items.filter("From eq 'Sub-Contractor'").select("ID,BaseName,SubContractor").get();
    for (let i = 0; i < document.length; i++) {
      if (document[i].SubContractor == option.text) {
        let transmittalOutlookDocument = {
          key: document[i].ID,
          text: document[i].BaseName
        };
        transmittalOutlookDocumentArray.push(transmittalOutlookDocument);
      }
    }
    sorted_transmittalOutlookDocumentArray = _.orderBy(transmittalOutlookDocumentArray, 'text', ['asc']);
    if (sorted_transmittalOutlookDocumentArray.length == 0) {
      this.setState({
        disableOutlook: true
      });
    }
    this.setState({
      transmittalOutlookDocumentArray: sorted_transmittalOutlookDocumentArray
    });
  }
  // On document index change
  public async _documentIndexChange(option: { key: any; text: any }) {
    console.log(option);
    this.setState({ validDocumentIndex: "none" });
    if (this.state.gridDocument.length > 0) {
      let duplicate = this.state.gridDocument.filter(a => a.DocumentIndexId == option.key);
      if (duplicate.length != 0) {
        this.setState({
          documentSelectedDiv: false,
        });
      }
      else {
        const documentIndex = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.filter("ID eq '" + option.key + "'").get();
        const documentIndexItem = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.select("Owner/Title,Owner/ID,RevisionCoding/Title,RevisionCoding/ID,DocumentID,Title,SubcontractorDocumentNo,DocumentName").expand("Owner,RevisionCoding").filter("ID eq '" + option.key + "'").get();
        console.log(documentIndex);
        if (documentIndex[0].RevisionCodingId != null) {
          this.setState({
            revisionCodingId: documentIndexItem[0].RevisionCoding.ID
          });
        }
        this.setState({
          documentSelectedDiv: true,
          documentIndexID: option.key,
          documentIndexTitle: documentIndexItem[0].DocumentName,
          owner: documentIndexItem[0].Owner.Title,
          ownerId: documentIndexItem[0].Owner.ID,
          viewDocument: documentIndexItem[0].DocumentID,
          subContractorNumber: documentIndexItem[0].SubcontractorDocumentNo,
          validDocumentIndex: "none",
          title: documentIndexItem[0].Title

        });
      }
    }
    else {
      const documentIndex = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.filter("ID eq '" + option.key + "'").get();
      const documentIndexItem = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.select("Owner/Title,Owner/ID,RevisionCoding/Title,RevisionCoding/ID,DocumentID,Title,SubcontractorDocumentNo,DocumentName").expand("Owner,RevisionCoding").filter("ID eq '" + option.key + "'").get();
      console.log(documentIndex);
      if (documentIndex[0].RevisionCodingId != null) {
        this.setState({
          revisionCodingId: documentIndexItem[0].RevisionCoding.ID
        });
      }
      this.setState({
        documentIndexID: option.key,
        documentIndexTitle: documentIndexItem[0].DocumentName,
        owner: documentIndexItem[0].Owner.Title,
        ownerId: documentIndexItem[0].Owner.ID,
        documentSelectedDiv: true,
        viewDocument: documentIndexItem[0].DocumentID,
        subContractorNumber: documentIndexItem[0].SubcontractorDocumentNo,
        title: documentIndexItem[0].Title
      });
    }
  }
  // On outlook document change
  public async _outlookDocumentChange(option: { key: any; text: any }) {
    this.setState({
      validDocument: "none"
    });
    const document = await sp.web.getList(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary).items.getById(option.key).get();
    this.setState({
      subContractorNumber: document.SubContractorDocumentId,
      poNumber: document.PONumber,
      transmittalOutlookId: option.key
    });
  }
  // On upload project document
  public async _upload(e) {
    this.setState({
      validDocument: "none"
    });
    let doctype;
    let type;
    if (this.state.documentIndexID == "") {
      this.setState({ validDocumentIndex: "" });
      (document.querySelector("#newfile") as HTMLInputElement).value = null;
    }
    else {
      this.myfile = e.target.value;

      const di = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(parseInt(this.state.documentIndexID)).get();
      console.log(di);
      let sdid = di.SourceDocumentID;
      if (sdid != null) {
        let docname = di.DocumentName;
        var docsplitted = docname.split(".");
        doctype = docsplitted[docsplitted.length - 1];
        let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
        console.log(myfile);
        var splitted = myfile.name.split(".");
        type = splitted[splitted.length - 1];
        if (doctype != type) {
          this.setState({ validDocument: "" });
          (document.querySelector("#newfile") as HTMLInputElement).value = null;
        }
      }
      // let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
      // var splitted = myfile.name.split(".");
      // if (splitted.length > 2) {
      //   e.target.value = "";
      //   this.setState({
      //     uploadDocumentError: "",
      //   });
      //   setTimeout(() => {

      //     this.setState({
      //       uploadDocumentError: "none",
      //       });
      //   }, 2000);
      // }
    }
  }
  // On upload additional document
  public async _uploadadditional(e) {
    this.myfileadditional = e.target.value;
    let myfile = (document.querySelector("#externalFile") as HTMLInputElement).files[0];
    var splitted = myfile.name.split(".");
    console.log(splitted);
    // alert(splitted.length - 1);
    // if (splitted.length > 2) {
    //   e.target.value = "";
    //   this.setState({
    //     uploadAdditionalDocumentError: "",
    //   });
    //   setTimeout(() => {

    //     this.setState({
    //       uploadAdditionalDocumentError: "none",
    //     });
    //   }, 2000);
    // }

  }
  // On dcc change
  public _dccChange = (items: any[]) => {
    this.setState({ noDcc: "none" });
    let dccEmail;
    let dccName;
    let getSelectedDCC = [];
    for (let item in items) {
      dccEmail = items[item].secondaryText,
        dccName = items[item].text,
        getSelectedDCC.push(items[item].id);
    }
    this.setState({
      dccId: getSelectedDCC[0],
      dcc: dccName
    });
  }
  // On owner change
  public _ownerChange = (items: any[]) => {


    let getSelectedOwner = [];

    for (let item in items) {

      getSelectedOwner.push(items[item].id);
    }
    this.setState({
      ownerId: getSelectedOwner[0]
    });
  }
  // On received date change
  private _onRecievedDatePickerChange = (date?: Date): void => {
    this.setState({ recievedDate: date });
  }
  // on subcontractor number change
  public _subContractorNumberChange = (ev: React.FormEvent<HTMLInputElement>, subContractorNumber?: string) => {

    this.setState({ subContractorNumber: subContractorNumber || '' });

  }
  // on subcontractor contract number
  public _poNumberChange = (option: { key: any; text: any }) => {
    this.setState({ poNumber: option.text, poNumberID: option.key });
  }
  // On transmittal settings change
  public _onTransmittalSettingsChange(option: { key: any; text: any }) {
    this.setState({ notransmittal: "none", transmittalSettingsId: option.key });
  }
  // On revision coding change
  public _onRevisionCodingChange(option: { key: any; text: any }) {
    this.setState({ revisionCodingId: option.key });
  }
  // On increment revision checked
  private _onIncrementRevisionChecked = (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ isIncrement: true });
    }
    else {
      this.setState({ isIncrement: false });
    }
  }
  //Comment Change
  public _commentschange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ validComment: "none" });
    this.setState({ comments: comments });
  }
  // on add index
  public async _addindex() {
    let sourceDocumentId;
    let documenturl;
    let documentNameExtension;
    let transmittalOutlookName;
    let documentName = this.state.viewDocument + " " + this.state.title;
    // this.state.documentIndexTitle;
    var splitted;

    if (this.state.dcc == undefined) {
      this.setState({ noDcc: "" });
    }
    else if (this.state.documentIndexID == "") {
      this.setState({ validDocumentIndex: "" });
    }
    else if (this.state.transmittalSettingsId == null) {
      this.setState({ notransmittal: "" });
    }
    else if (this.state.comments == "") {
      this.setState({ validComment: "" });
    }
    else {
      if ((document.querySelector("#newfile") as HTMLInputElement).files[0] != null) {
        let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
        let myfileName = myfile.name;
        splitted = myfileName.split(".");
        documentNameExtension = documentName + '.' + splitted[splitted.length - 1];
        this.documentNameExtension = documentNameExtension;
        if (myfile.size) {

          this.addDocument.push({
            SubContractor: this.state.subContractorID,
            DocumentControllerId: this.currentUserId,
            DocumentIndexId: this.state.documentIndexID,
            DocumentIndex: this.state.documentIndexTitle,
            OwnerId: this.state.ownerId,
            Owner: this.state.owner,
            SubContractorNumber: this.state.subContractorNumber,
            ReceivedDate: moment(this.state.recievedDate).format("DD/MM/YYYY"),
            PONumber: this.state.poNumber,
            TransmittalCodeId: this.state.transmittalSettingsId,
            RevisionCodeId: this.state.revisionCodingId,
            IncrementRevision: this.state.isIncrement,
            Comments: this.state.comments,
            content: myfile,
            RecieveDate: this.state.recievedDate,
            DocumentName: documentNameExtension,
            fileAttach: "FileAttached",
            type: splitted[splitted.length - 1]
          });
          this.setState({
            gridDocument: this.addDocument,
            showGrid: false,
            comments: "",
            isIncrement: false,
            revisionCodingId: null,
            transmittalSettingsId: null,
            poNumber: "",
            subContractorNumber: "",
            documentIndexID: null,
            documentIndexTitle: "",
            ownerId: null,
            owner: "",
            transmittalOutlookId: null
          });
          this.myfile.value = "";
        }
      }
      else if (this.state.transmittalOutlookId != "" && this.state.transmittalOutlookId != null) {
        let content;
        await sp.web.getList(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary).items.select("LinkFilename,ID").get().then(async transmittalOutlookdoc => {
          console.log(transmittalOutlookdoc);
          for (let j = 0; j < transmittalOutlookdoc.length; j++) {
            if (transmittalOutlookdoc[j].ID == this.state.transmittalOutlookId) {
              transmittalOutlookName = transmittalOutlookdoc[j].LinkFilename;
            }
          }
        });
        splitted = transmittalOutlookName.split(".");
        documentNameExtension = documentName + '.' + splitted[splitted.length - 1];
        this.documentNameExtension = documentNameExtension;
        await sp.web.getFileByServerRelativeUrl(this.props.siteUrl + "/" + this.props.transmittalOutlookLibrary + "/" + transmittalOutlookName).getBuffer()
          .then(templateData => {
            console.log(templateData);
            content = templateData;
          });
        this.addDocument.push({
          SubContractor: this.state.subContractorID,
          DocumentControllerId: this.currentUserId,
          DocumentIndexId: this.state.documentIndexID,
          DocumentIndex: this.state.documentIndexTitle,
          OwnerId: this.state.ownerId,
          Owner: this.state.owner,
          SubContractorNumber: this.state.subContractorNumber,
          ReceivedDate: moment(this.state.recievedDate).format("DD/MM/YYYY"),
          RecieveDate: this.state.recievedDate,
          PONumber: this.state.poNumber,
          TransmittalCodeId: this.state.transmittalSettingsId,
          RevisionCodeId: this.state.revisionCodingId,
          IncrementRevision: this.state.isIncrement,
          Comments: this.state.comments,
          TransmittalOutlookId: this.state.transmittalOutlookId,
          SourceDocumentId: sourceDocumentId,
          SourceDocumentUrl: documenturl,
          DocumentName: documentNameExtension,
          content: content,
          fileAttach: "NoFileAttached",
          type: splitted[splitted.length - 1]
        });
        this.setState({
          gridDocument: this.addDocument,
          showGrid: false,
          comments: "",
          isIncrement: false,
          revisionCodingId: null,
          transmittalSettingsId: null,
          poNumber: "",
          subContractorNumber: "",
          documentIndexID: null,
          documentIndexTitle: "",
          ownerId: null,
          owner: "",
          transmittalOutlookId: null,
        });

      }
      else {
        this.setState({
          validDocument: ""
        });
      }
    }

  }
  // on date picker change
  public _onDatePickerChange = (date?: Date): void => {
    this.setState({ externalDate: date });
  }
  // on external document comment change
  public _externalCommentsChange = (ev: React.FormEvent<HTMLInputElement>, comments?: any) => {
    this.setState({ validAdditionalComment: "none" });
    this.setState({ externalComments: comments });
  }
  // Add external document
  public _addexternalindex() {
    if (this.state.dcc == undefined) {
      this.setState({ noDcc: "" });
    }
    else if (this.state.externalComments == "") {
      this.setState({ validAdditionalComment: "" });
    }
    else {
      if ((document.querySelector("#externalFile") as HTMLInputElement).files[0] != null) {
        let myfile = (document.querySelector("#externalFile") as HTMLInputElement).files[0];
        var docname = myfile.name;

        this.addExternalDocument.push({
          Content: myfile,
          DocName: docname,
          ExternalDate: moment(this.state.externalDate).format("DD/MM/YYYY"),
          ExtDate: this.state.externalDate,
          Comments: this.state.externalComments
        });
        console.log(this.addExternalDocument);
        this.setState({
          gridExternalDocument: this.addExternalDocument,
          showExternalGrid: false,
          externalComments: "",
        });
        this.myfileadditional.value = "";
      }
      else { }
    }
  }
  // Save as draft button click
  public async _saveAsDraft() {
    if (this.state.gridDocument.length == 0) {
      this.setState({ statusMessage: { isShowMessage: true, message: "Please add all mandatory fields ", messageType: 1 }, });
      setTimeout(() => {
        this.setState({ statusMessage: { isShowMessage: false, message: "Please add all mandatory fields ", messageType: 1 }, });
      }, 2000);
    }
    else {
      this.setState({ submitDisable: true });
      this.status = "Save";
      if (this.transmittalID == null || this.transmittalID == "") {
        console.log("Save as draft button clicked");
        this.setState({ checksend: "" });
        this._idGeneration();
      }
      else {
        this.setState({ checksend: "" });
        this._newAddDocument();
      }
    }
  }
  // submit button click
  public _submit() {
    if (this.state.gridDocument.length == 0) {
      this.setState({ statusMessage: { isShowMessage: true, message: "Please add all mandatory fields ", messageType: 1 }, });
      setTimeout(() => {
        this.setState({ statusMessage: { isShowMessage: false, message: "Please add all mandatory fields ", messageType: 1 }, });
      }, 2000);
    }
    else {
      this.setState({ submitDisable: true });
      this.status = "Submit";
      //  alert("add");
      if (this.transmittalID == null || this.transmittalID == "") {
        console.log("Save as draft button clicked");
        this.setState({ checksend: "" });
        this._idGeneration();
      }
      else {
        this.setState({ checksend: "" });
        this._newAddDocument();
      }
    }
  }
  // id generation
  public _idGeneration() {
    let prefix;
    let separator;
    let sequenceNumber;
    let title;
    let counter;
    let transmittalSubContractorId;
    let id;
    let increment;
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalIdSettings).items.filter("TransmittalCategory eq 'Inbound Sub-contractor'").get().then(transmittalIdSettingsItems => {
      console.log("transmittalIdSettingsItems", transmittalIdSettingsItems);
      prefix = transmittalIdSettingsItems[0].Prefix;
      separator = transmittalIdSettingsItems[0].Separator;
      sequenceNumber = transmittalIdSettingsItems[0].SequenceNumber;
      title = transmittalIdSettingsItems[0].Title;
      counter = transmittalIdSettingsItems[0].Counter;
      id = transmittalIdSettingsItems[0].ID;
      increment = counter + 1;
      var incrementvalue = increment.toString();
      this._transmittalSequenceNumber(incrementvalue, sequenceNumber);
      transmittalSubContractorId = prefix + separator + title + separator + this.state.projectNumber + separator + this.state.incrementSequenceNumber;
      console.log("transmittalID", transmittalSubContractorId);
      this.setState({
        transmittalSubContractorId: transmittalSubContractorId,
      });
    }).then(afterid => {
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalIdSettings).items.getById(id).update({
        Counter: increment
      });
      this._addDocument();
    });
  }
  // Transmittal sequence number
  private _transmittalSequenceNumber(incrementvalue, sequenceNumber) {
    var incrementSequenceNumber = incrementvalue;
    while (incrementSequenceNumber.length < sequenceNumber)
      incrementSequenceNumber = "0" + incrementSequenceNumber;
    console.log(incrementSequenceNumber);
    this.setState({
      incrementSequenceNumber: incrementSequenceNumber,
    });
  }
  // Add document
  public async _addDocument() {
    let additionalDocumentId;
    let documentName;
    let sourceDocumentId;
    let documenturl;
    let docServerUrl;
    let splitdocUrl;
    var splitted;
    let documentIdname;
    const inboundTransmittalHeader = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalHeaderList).items.add({
      Title: this.state.transmittalSubContractorId,
      SubContractor: this.state.subContractor,
      SubContractorID: Number(this.state.subContractorID),
      TransmittalStatus: "Draft",
      TransmittalCategory: "Sub-Contractor",
      TransmittalDate: this.today,
      DocumentControllerId: this.state.dccId,
    });

    this.setState({ inboundTransmittalHeaderId: inboundTransmittalHeader.data.ID });
    this.transmittalID = inboundTransmittalHeader.data.ID;

    const afterheaderid = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalHeaderList).items.getById(inboundTransmittalHeader.data.ID).update({
      TransmittalLink: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "Project workspace",
        Url: this.props.siteUrl + "/SitePages/" + this.props.inboundTransmittalSitePage + ".aspx?trid=" + inboundTransmittalHeader.data.ID + ""
      },
      TransmittalDetails: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "Transmittal Details",
        Url: this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalDetailList + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=6da3a1b3%2D0155%2D48d9%2Da7c7%2Dd2e862c07db5"
      },
      InboundAdditionalDetails: {
        "__metadata": { type: "SP.FieldUrlValue" },
        Description: "Inbound Additional Details",
        Url: this.props.siteUrl + "/" + this.props.additionalDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=TransmittalID&FilterValue1=" + inboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=5a376168-dc2b-49f0-aa7b-9c918fe1b614"
      },
    });
    if (afterheaderid) {
      if (this.state.gridDocument.length > 0) {

        for (let i = 0; i < this.state.gridDocument.length; i++) {

          await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalDetailList).items.add({
            Title: this.state.transmittalSubContractorId,
            TransmittalHeaderId: inboundTransmittalHeader.data.ID,
            DocumentIndexId: this.state.gridDocument[i].DocumentIndexId,
            Comments: this.state.gridDocument[i].Comments,
            OwnerId: this.state.gridDocument[i].OwnerId,
            ReceivedDate: this.state.gridDocument[i].RecieveDate,
            TransmittalCodeId: this.state.gridDocument[i].TransmittalCodeId,
            RevisionCodingId: this.state.gridDocument[i].RevisionCodeId,
            IncrementRevision: this.state.gridDocument[i].IncrementRevision,
            PONumber: this.state.gridDocument[i].PONumber,
            SubContractorDocumentNumber: this.state.gridDocument[i].SubContractorNumber
          });
          this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.state.gridDocument[i].DocumentIndexId + "";
          this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + this.state.gridDocument[i].DocumentIndexId + "&mode=expiry";
          this.documentIndexID = this.state.gridDocument[i].DocumentIndexId;
          const indexItems: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).get();
          console.log(indexItems);
          let documentid = indexItems.DocumentID;
          documentName = indexItems.DocumentName;
          let myfile = this.state.gridDocument[i].content;
          let myfileName;
          myfileName = this.state.gridDocument[i].DocumentName;
          var splitted = myfileName.split(".");
          documentIdname = documentid + '.' + splitted[splitted.length - 1];
          const fileuploaded = await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/").files.add(documentIdname, myfile, true);
          if (fileuploaded) {
            const item = await fileuploaded.file.getItem();
            sourceDocumentId = item["ID"];
            if (this.state.gridDocument[i].type == "pdf" || this.state.gridDocument[i].type == "Pdf" || this.state.gridDocument[i].type == "PDF") {
              documenturl = item["ServerRedirectedEmbedUrl"];
            }
            else {
              docServerUrl = item["ServerRedirectedEmbedUrl"];
              splitdocUrl = docServerUrl.split("&", 2);
              documenturl = splitdocUrl[0];
            }
            this.sourceDocumentID = sourceDocumentId;
            let reviewerId;
            if (indexItems.ReviewersId == null) { reviewerId = []; }
            else { reviewerId = indexItems.ReviewersId; }
            const updatelist = await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
              Title: indexItems.Title,
              DocumentID: indexItems.DocumentID,
              ReviewersId: {
                results: reviewerId
              },
              DocumentName: myfileName,
              BusinessUnit: indexItems.BusinessUnit,
              Category: indexItems.Category,
              SubCategory: indexItems.SubCategory,
              ApproverId: indexItems.ApproverId,
              Revision: "-",
              WorkflowStatus: "Draft",
              DocumentStatus: "Active",
              DocumentIndexId: this.state.gridDocument[i].DocumentIndexId,
              PublishFormat: indexItems.PublishFormat,
              Template: indexItems.Template,
              OwnerId: this.state.gridDocument[i].OwnerId,
              DepartmentName: indexItems.DepartmentName,
              RevisionHistory: {
                "__metadata": { type: "SP.FieldUrlValue" },
                Description: "Revision History",
                Url: this.revisionHistoryUrl
              },
              TransmittalDocument: indexItems.TransmittalDocument,
              ExternalDocument: indexItems.ExternalDocument,
              RevisionCodingId: this.state.gridDocument[i].RevisionCodeId,
              RevisionLevelId: indexItems.RevisionLevelId,
              DocumentControllerId: this.state.dccId,
              CustomerDocumentNo: indexItems.CustomerDocumentNo,
              SubcontractorDocumentNo: indexItems.SubcontractorDocumentNo
            });
            if (indexItems.ExpiryDate != null) {
              const updateexpiry = await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
                ExpiryDate: indexItems.ExpiryDate,
                ExpiryLeadPeriod: indexItems.ExpiryLeadPeriod
              });
            }
            if (updatelist) {
              if (indexItems.SourceDocumentID == null) {
                const log = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.add({
                  Title: indexItems.DocumentID,
                  Status: "Document Created",
                  LogDate: this.today,
                  Revision: "-",
                  DocumentIndexId: parseInt(this.documentIndexID),
                });
              }
              console.log(this.state.gridDocument[i]);
              const afterall = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(parseInt(this.state.gridDocument[i].DocumentIndexId)).update({
                SourceDocumentID: this.sourceDocumentID,
                DocumentName: myfileName,
                CreateDocument: true,
                TransmittalStatus: "Ongoing",
                // SourceDocument: {
                //   "__metadata": { type: "SP.FieldUrlValue" },
                //   Description: myfileName,
                //   Url: documenturl
                // },
                SourceDocument: {
                  "__metadata": { type: "SP.FieldUrlValue" },
                  Description: myfileName,
                  Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.state.gridDocument[i].DocumentIndexId) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
                },
                RevokeExpiry: {
                  "__metadata": { type: "SP.FieldUrlValue" },
                  Description: "Revoke",
                  Url: this.revokeUrl
                },
              });
              if (afterall) {
                this._triggerPermission(this.sourceDocumentID);
              }
            }
          }
        }

      }
      if (this.state.gridExternalDocument.length > 0) {
        for (let i in this.state.gridExternalDocument) {
          let myfile = this.state.gridExternalDocument[i].Content;
          splitted = this.state.gridExternalDocument[i].DocName.split(".", 2);
          let Name = myfile.name;
          let myfileName = this.state.transmittalSubContractorId + " " + myfile.name;

          if (myfile.size) {
            const fileUploaded = await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/" + this.props.additionalDocumentLibrary + "/").files.add(myfileName, myfile, true);
            console.log("File Uploaded");
            const item = await fileUploaded.file.getItem();
            console.log(item);
            additionalDocumentId = item["ID"];
            this.additionalDocumentId = additionalDocumentId;
            this.setState({ additionalDocumentId: additionalDocumentId });
            const additional = await sp.web.getList(this.props.siteUrl + "/" + this.props.additionalDocumentLibrary).items.getById(this.additionalDocumentId).update({
              TransmittalIDId: inboundTransmittalHeader.data.ID,
              ReceivedDate: this.state.gridExternalDocument[i].ExtDate,
              SubContractor: this.state.subContractor,
              SubContractorID: Number(this.state.subContractorID),
              Comments: this.state.gridExternalDocument[i].Comments,
              Title: Name
            });


          }
        }
      }
      if (this.status == "Submit") {

        //alert("insideupdatell");
        await this._updateall();
      }
      else {

        this.setState({
          checksend: "none",
          statusMessage: { isShowMessage: true, message: this.dataSaved, messageType: 4 },
          submitDisable: false
        });
        setTimeout(() => {
          window.location.replace(this.siteUrl);
        }, 8000);
      }

    }

  }
  // Add document after save as draft
  public async _newAddDocument() {
    let additionalDocumentId;
    let documentName;
    let sourceDocumentId;
    let documenturl;
    let docServerUrl;
    let splitdocUrl;
    var splitted;
    let headerid = parseInt(this.transmittalID);
    let documentIdname;
    const header = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalHeaderList).items.getById(headerid).update({
      DocumentControllerId: this.state.dccId
    });
    if (header) {
      if (this.state.gridDocument.length > this.state.currentInboundDetailItem.length) {
        for (var k = this.state.currentInboundDetailItem.length; k < this.state.gridDocument.length; k++) {
          const detail = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalDetailList).items.add({
            Title: this.state.transmittalInboundID,
            TransmittalHeaderId: headerid,
            DocumentIndexId: this.state.gridDocument[k].DocumentIndexId,
            Comments: this.state.gridDocument[k].Comments,
            OwnerId: this.state.gridDocument[k].OwnerId,
            ReceivedDate: this.state.gridDocument[k].RecieveDate,
            TransmittalCodeId: this.state.gridDocument[k].TransmittalCodeId,
            RevisionCodingId: this.state.gridDocument[k].RevisionCodeId,
            IncrementRevision: this.state.gridDocument[k].IncrementRevision,
            PONumber: this.state.gridDocument[k].PONumber,
            SubContractorDocumentNumber: this.state.gridDocument[k].SubContractorNumber,
          });
          this.revisionHistoryUrl = this.props.siteUrl + "/SitePages/" + this.props.revisionHistoryPage + ".aspx?did=" + this.state.gridDocument[k].DocumentIndexId + "";
          this.revokeUrl = this.props.siteUrl + "/SitePages/" + this.props.revokePage + ".aspx?did=" + this.state.gridDocument[k].DocumentIndexId + "&mode=expiry";
          this.documentIndexID = this.state.gridDocument[k].DocumentIndexId;
          let ownerId = this.state.gridDocument[k].OwnerId;
          let RevisionCodeId = this.state.gridDocument[k].RevisionCodeId;
          const indexItems: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.documentIndexID).get();
          console.log(indexItems);
          documentName = indexItems.DocumentName;
          let documentid = indexItems.DocumentID;
          let myfile = this.state.gridDocument[k].content;
          let myfileName = this.state.gridDocument[k].DocumentName;
          var splitted = myfileName.split(".");
          documentIdname = documentid + '.' + splitted[splitted.length - 1];
          const fileUploaded = await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/").files.add(documentIdname, myfile, true);
          console.log("File Uploaded");
          const item = await fileUploaded.file.getItem();
          console.log(item);
          sourceDocumentId = item["ID"];
          this.sourceDocumentID = sourceDocumentId;
          // if(this.state.gridDocument[k].type=="pdf"||this.state.gridDocument[k].type=="Pdf"||this.state.gridDocument[k].type=="PDF"){
          //   documenturl = item["ServerRedirectedEmbedUrl"];
          // }
          // else{
          // docServerUrl = item["ServerRedirectedEmbedUrl"];
          // splitdocUrl = docServerUrl.split("&", 2);
          // documenturl = splitdocUrl[0];
          // }
          let reviewerId;
          if (indexItems.ReviewersId == null) { reviewerId = []; }
          else { reviewerId = indexItems.ReviewersId; }
          const SDLib = await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(this.sourceDocumentID).update({
            Title: indexItems.Title,
            DocumentID: indexItems.DocumentID,
            ReviewersId: {
              results: reviewerId
            },
            DocumentName: myfileName,
            BusinessUnit: indexItems.BusinessUnit,
            Category: indexItems.Category,
            SubCategory: indexItems.SubCategory,
            ApproverId: indexItems.ApproverId,
            Revision: "-",
            WorkflowStatus: "Draft",
            DocumentStatus: "Active",
            DocumentIndexId: this.documentIndexID,
            PublishFormat: indexItems.PublishFormat,
            Template: indexItems.Template,
            OwnerId: ownerId,
            DepartmentName: indexItems.DepartmentName,
            RevisionHistory: {
              "__metadata": { type: "SP.FieldUrlValue" },
              Description: "Revision History",
              Url: this.revisionHistoryUrl
            },
            TransmittalDocument: indexItems.TransmittalDocument,
            ExternalDocument: indexItems.ExternalDocument,
            RevisionCodingId: RevisionCodeId,
            RevisionLevelId: indexItems.RevisionLevelId,
            DocumentControllerId: this.state.dccId,
            CustomerDocumentNo: indexItems.CustomerDocumentNo,
            SubcontractorDocumentNo: indexItems.SubcontractorDocumentNo

          });
          if (SDLib) {
            if (indexItems.SourceDocumentID == null) {
              const log = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentRevisionLogList).items.add({
                Title: indexItems.DocumentID,
                Status: "Document Created",
                LogDate: this.today,
                Revision: "-",
                DocumentIndexId: parseInt(this.documentIndexID),
              });
            }
            const didata = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(parseInt(this.documentIndexID)).update({
              SourceDocumentID: this.sourceDocumentID,
              DocumentName: myfileName,
              CreateDocument: true,
              TransmittalStatus: "Ongoing",
              // SourceDocument: {
              //   "__metadata": { type: "SP.FieldUrlValue" },
              //   Description: myfileName,
              //   Url: documenturl
              // },
              SourceDocument: {
                "__metadata": { type: "SP.FieldUrlValue" },
                Description: myfileName,
                Url: this.props.siteUrl + "/" + this.props.sourceDocumentLibrary + "/Forms/AllItems.aspx?FilterField1=DocumentIndex&FilterValue1=" + parseInt(this.documentIndexID) + "&FilterType1=Lookup&viewid=c46304af-9c51-4289-bea2-ddb05655f7c2"
              },
              RevokeExpiry: {
                "__metadata": { type: "SP.FieldUrlValue" },
                Description: "Revoke",
                Url: this.revokeUrl
              },
            });
            if (didata) {
              this._triggerPermission(this.sourceDocumentID);
            }
          }

        }
      }
      if (this.state.gridExternalDocument.length > this.state.currentInboundAdditionalItem.length) {
        for (let g = this.state.currentInboundAdditionalItem.length; g < this.state.gridExternalDocument.length; g++) {

          let myfile = this.state.gridExternalDocument[g].Content;
          let Name = myfile.name;
          splitted = this.state.gridExternalDocument[g].DocName.split(".", 2);
          let myfileName = this.state.transmittalInboundID + " " + myfile.name;
          if (myfile.size) {
            // add file to source library
            sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/" + this.props.additionalDocumentLibrary + "/").files.add(myfileName, myfile, true)
              .then(fileUploaded => {
                console.log("File Uploaded");
                fileUploaded.file.getItem().then(async item => {
                  console.log(item);
                  additionalDocumentId = item["ID"];
                  this.additionalDocumentId = additionalDocumentId;
                  this.setState({ additionalDocumentId: additionalDocumentId });
                  // update metadata
                  sp.web.getList(this.props.siteUrl + "/" + this.props.additionalDocumentLibrary).items.getById(this.additionalDocumentId).update({
                    TransmittalIDId: this.transmittalID,
                    ReceivedDate: this.state.gridExternalDocument[g].ExtDate,
                    SubContractor: this.state.subContractor,
                    SubContractorID: Number(this.state.subContractorID),
                    Comments: this.state.gridExternalDocument[g].Comments,
                    Title: Name
                  });
                });
              });
          }
        }
      }
      if (this.status == "Submit") {
        await this._updateall();
      }
      else {
        this.setState({
          checksend: "none",
          statusMessage: { isShowMessage: true, message: this.dataSaved, messageType: 4 },
          submitDisable: false
        });
        setTimeout(() => {
          window.location.replace(this.siteUrl);
        }, 8000);
      }
    }
  }
  // update metadatas
  public async _updateall() {
    // alert("update");
    let SourceDocumentID;
    let headerid;
    let DocumentID;
    if (this.transmittalID == null || this.transmittalID == "") {
      headerid = this.transmittalID;
    }
    else {
      headerid = this.transmittalID;
    }
    // alert(headerid);
    const headeradd = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalHeaderList).items.getById(parseInt(headerid)).update({
      TransmittalStatus: "Completed",
      // TransmittedBy:this.currentUserId
    });
    // alert(this.state.gridDocument.length);
    if (this.state.gridDocument.length > 0) {
      for (let h = 0; h < this.state.gridDocument.length; h++) {
        let diId = this.state.gridDocument[h].DocumentIndexId;
        const di: any = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(this.state.gridDocument[h].DocumentIndexId).get();
        console.log(di);
        // alert(di.ID);
        SourceDocumentID = di.SourceDocumentID;
        DocumentID = di.DocumentID;
        if (di) {
          const docUpdate = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(parseInt(diId)).update({
            TransmittalStatus: "Completed",
            SubcontractorDocumentNo: this.state.gridDocument[h].SubContractorNumber,
            PONumber: this.state.gridDocument[h].PONumber,
            TransmittalLocation: "IN from Sub-Contractor",
            TransmittalDocument: true
          });
          if (docUpdate) {
            const sourceUpdate = await sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibrary).items.getById(parseInt(SourceDocumentID)).update({
              TransmittalStatus: "Completed",
              SubcontractorDocumentNo: this.state.gridDocument[h].SubContractorNumber,
              PONumber: this.state.gridDocument[h].PONumber,
              TransmittalLocation: "IN from Sub-Contractor"
            });
            if (sourceUpdate) {
              sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalHistoryLogList).items.add({
                Title: DocumentID,
                Status: "IN from Sub-Contractor",
                DocumentIndexId: diId,
                LogDate: this.today,
              });
            }
          }
        }
      }
      this.setState({
        checksend: "none",
        statusMessage: { isShowMessage: true, message: this.dataSaved, messageType: 4 },
        submitDisable: false
      });
      setTimeout(() => {
        window.location.replace(this.siteUrl);
      }, 5000);
    }


  }
  // give permission
  protected async _triggerPermission(sourceDocumentID) {
    const laUrl = await this.reqWeb.getList("/sites/" + this.props.hubsite + "/Lists/" + this.props.requestList).items.filter("Title eq 'EMEC_DocumentPermission-Create Document'").get();
    console.log("Posturl", laUrl[0].PostUrl);
    this.postUrl = laUrl[0].PostUrl;
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'ItemId': sourceDocumentID
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);


  }
  // on cancel
  private _onCancel = () => {
    this.setState({
      cancelConfirmMsg: "",
      confirmDialog: false,
    });


  }
  //Cancel confirm
  private _confirmYesCancel = () => {
    this.setState({
      statusKey: "",
      comments: "",
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
    this.validator.hideMessages();
    window.location.replace(this.siteUrl);
  }
  //Not Cancel
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });

  }
  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmDialog: true,
    });
  }
  private dialogStyles = { main: { maxWidth: 500 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to cancel?',

  };
  private dialogDeleteProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to delete?',

  };
  private modalProps = {
    isBlocking: true,
  };
  // On deletebutton click
  private _openDeleteConfirmation = (items, key, type) => {

    if (this.transmittalID == "" || this.transmittalID == null || this.transmittalID == undefined) {
      this.setState({
        deleteConfirmMsg: "",
        confirmDeleteDialog: false,
      });
      this.validator.hideMessages();
      console.log(items[key]);
      console.log(items.DocumentIndexId);
      if (type == "ProjectDocuments") {
        this.typeForDelete = "ProjectDocuments";
        this.keyForDelete = key;
      } else if (type == "AdditionalDocuments") {
        this.typeForDelete = "AdditionalDocuments";
        this.keyForDelete = key;
      }
    }
    else {
      this.setState({
        deleteConfirmMsg: "",
        confirmDeleteDialog: false,
        tempDocIndexIDForDelete: items.DetailId,

      });
      this.docIndexId = items.DocumentIndexId;
      this.validator.hideMessages();
      console.log(items[key]);
      if (type == "ProjectDocuments") {
        // alert(items.outboundDetailsID);
        this.typeForDelete = "ProjectDocuments";
        this.keyForDelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.DetailId,
        });
      } else if (type == "AdditionalDocuments") {
        // alert("additionalid" + items.additionalDocumentID);
        this.typeForDelete = "AdditionalDocuments";
        this.keyForDelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.additionalId,
        });
      }
    }

  }
  // confirm delete item
  private _confirmDeleteItem = async (docID, items, key) => {
    if (this.transmittalID == "" || this.transmittalID == null || this.transmittalID == undefined) {
      this.setState({
        deleteConfirmMsg: "none",
        confirmDeleteDialog: true,
      });
      this.validator.hideMessages();
      if (this.typeForDelete == "ProjectDocuments") {
        this.itemDeleteFromGrid(items, key);
      }
      else if (this.typeForDelete == "AdditionalDocuments") {
        this.itemDeleteFromExternalGrid(items, key);
      }

    }
    else {
      this.setState({
        deleteConfirmMsg: "none",
        confirmDeleteDialog: true,
      });
      this.validator.hideMessages();
      console.log(items[key]);
      // alert(docID);
      if (this.typeForDelete == "ProjectDocuments") {
        // alert(docID);
        this.itemDeleteFromGrid(items, key);
        let list = sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalDetailList);
        await list.items.getById(parseInt(docID)).delete();
        const afterall = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.documentIndexList).items.getById(parseInt(this.docIndexId)).update({
          TransmittalStatus: "New",
        });
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.inboundTransmittalDetailList).items.select("TransmittalHeaderId,DocumentIndex/ID,DocumentIndex/Title,DocumentIndex/DocumentName,Owner/ID,Owner/Title,SubContractorDocumentNumber,ReceivedDate,Comments,ID").expand("DocumentIndex,Owner").filter("TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ").get().then(inboundTransmittalDetailList => {
          console.log("inboundTransmittalDetailList", inboundTransmittalDetailList);
          this.setState({
            currentInboundDetailItem: inboundTransmittalDetailList,
          });
        });
        this.setState({
          gridDocument: this.state.gridDocument,
        });
      }
      else if (this.typeForDelete == "AdditionalDocuments") {
        // alert("additional" + docID);
        this.itemDeleteFromExternalGrid(items, key);
        let list = sp.web.getList(this.props.siteUrl + "/" + this.props.additionalDocumentLibrary + "/");
        await list.items.getById(parseInt(docID)).delete();
        sp.web.getList(this.props.siteUrl + "/" + this.props.additionalDocumentLibrary).items.filter("TransmittalIDId eq '" + this.transmittalID + "' ").get().then(inboundAdditionalDocumentsListName => {
          console.log("inboundAdditionalDocumentsListName", inboundAdditionalDocumentsListName);
          this.setState({
            currentInboundAdditionalItem: inboundAdditionalDocumentsListName,
          });
        });
        this.setState({
          gridExternalDocument: this.state.gridExternalDocument,

        });
      }
    }
  }
  // Delete item from grid
  public itemDeleteFromGrid(items, key) {
    console.log(items);
    this.state.gridDocument.splice(key, 1);
    console.log("after removal", this.state.gridDocument);

    this.setState({
      gridDocument: this.state.gridDocument,
      deleteConfirmMsg: "none",
      confirmDeleteDialog: true,
    });
  }
  // Delete item from external grid
  public itemDeleteFromExternalGrid(items, key) {
    this.state.gridExternalDocument.splice(key, 1);
    console.log("after removal", this.state.gridExternalDocument);

    this.setState({
      gridExternalDocument: this.state.gridExternalDocument,
      deleteConfirmMsg: "none",
      confirmDeleteDialog: true,
    });
  }
  // on format date field
  private _onFormatDate = (date: Date): string => {
    const dat = date;
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  };
  public render(): React.ReactElement<IEmecInboundSubContractorProps> {
    const DeleteIcon: IIconProps = { iconName: 'Delete' };
    const AddIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
    return (
      <div className={styles.emecInboundSubContractor}>
        <div style={{ display: this.state.loaderDisplay }}>
          <ProgressIndicator label="Loading......" />
        </div>
        <div style={{ display: this.state.access }}>
          <div className={styles.border} >
            <div className={styles.alignCenter}>{this.props.webpartHeader}</div>

            <div className={styles.divrow} style={{ display: this.state.transmittalNo }}>
              <div className={styles.wdthrgt}><Label>Transmittal ID : {this.state.transmittalInboundID} </Label></div>
              <div className={styles.wdthlft}><Label>Transmittal Date : {this.state.transmittalDate}</Label></div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}><Label >Project Name : {this.state.projectName} </Label></div>
              <div className={styles.wdthlft}><Label >Project Number :{this.state.projectNumber} </Label></div>
            </div>

            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                {/* <Label >Sub-Contractor : </Label> */}
                <Dropdown
                  placeholder="Sub-Contractor"
                  label="Select Sub-Contractor"
                  options={this.state.subContractorArray}
                  onChanged={this._subContactorChanged}
                  selectedKey={this.state.subContractorID}
                />
                <div style={{ color: "#dc3545", display: this.state.noSubContactor }} >Please Select SubContactor</div>
              </div>
              <div className={styles.wdthlft}>
                <PeoplePicker
                  context={this.props.context}
                  titleText="DCC"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  disabled={false}
                  ensureUser={true}
                  onChange={(items) => this._dccChange(items)}
                  defaultSelectedUsers={[this.state.dcc]}
                  showHiddenInUI={false}
                  // isRequired={true}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
                <div style={{ color: "#dc3545", display: this.state.noDcc }}>
                  Please Select Document Controller</div>
              </div>
            </div>

            {/* <div style={{marginTop:"10px", border: "1px solid black", width:"80%"}}></div> */}
            <div className={styles.subSection}>Project Document  </div>


            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <Dropdown placeholder="Select Document Index"
                  label="Document Index"
                  options={this.state.documentIndexArray}
                  onChanged={this._documentIndexChange}
                  selectedKey={this.state.documentIndexID}
                  required
                />
                <div hidden={this.state.documentSelectedDiv} style={{ color: "#dc3545", fontWeight: "bold" }}>Sorry! Document is already selected</div>
                <div style={{ color: "#dc3545", display: this.state.validDocumentIndex }} >Please Select DocumentIndex</div>
              </div>
              <div className={styles.wdthlft}>
                <PeoplePicker
                  context={this.props.context}
                  titleText="Owner"
                  personSelectionLimit={1}
                  groupName={""} // Leave this blank in case you want to filter from all users
                  showtooltip={true}
                  disabled={false}
                  ensureUser={true}
                  onChange={(items) => this._ownerChange(items)}
                  defaultSelectedUsers={[this.state.owner]}
                  showHiddenInUI={false}
                  // isRequired={true}
                  principalTypes={[PrincipalType.User]}
                  resolveDelay={1000}
                />
              </div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <Label >Upload Document:</Label>
                <input type="file" name="myFile" id="newfile" onChange={(e) => this._upload(e)} ref={ref => this.myfile = ref} ></input>
              </div>
              <div className={styles.wdthlft} >
                <Dropdown
                  placeholder="Search Document"
                  label="Select Document"
                  options={this.state.transmittalOutlookDocumentArray}
                  onChanged={this._outlookDocumentChange}
                  selectedKey={this.state.transmittalOutlookId}
                  disabled={this.state.disableOutlook}
                />

              </div>
            </div>
            <div style={{ color: "#dc3545", display: this.state.uploadDocumentError, marginLeft: "9px" }}>Sorry this document is unable to process due to unwanted characters.Please rename the document and try again.</div>
            <div style={{ color: "#dc3545", display: this.state.validDocument }}>Please select valid Document</div>

            {/* <div className={styles.mt} >
            <div style={{display:"block"}}>  
            <div hidden ={this.state.documentSelectedDiv} style={{fontWeight :"bold"}}>Selected Document : {this.state.viewDocument}</div>
            <SearchBox placeholder="Document from Document Index" title="Document Index"  onSearch={newValue => console.log('value is ' + newValue)}  className={styles['ms-SearchBox']} onChange={this._onDocumentIndexFilter} /> 
            </div>
            <div  style={{ display: this.state.searchDiv,padding:"5px", marginLeft: "8px", height:"100px",width:"520px",boxShadow:"0 4px 15px rgba(0,0,0,0.2)",overflowY:"scroll"}} >
            {this.state.items.map((searchItems, key) => {
              return (
                <div style={{ padding: "1rem 0 0 0"}}>
                <table style={{ cursor: "pointer" }}>
                <tr>
                <td onClick={()=>this._documentIndexChange(searchItems)}>{searchItems.DocumentID}</td>                                        
                </tr>
                </table>
                </div>
                );
              })}
            </div>
          </div> */}

            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <DatePicker label="Received Date"
                  value={this.state.recievedDate}
                  onSelectDate={this._onRecievedDatePickerChange}
                  placeholder="Select a date"
                  formatDate={this._onFormatDate} />
              </div>
              <div className={styles.wdthlft}>
                <TextField label="Sub-Contractor Document Number"
                  onChange={this._subContractorNumberChange}
                  value={this.state.subContractorNumber} readOnly></TextField>
              </div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <Dropdown
                  placeholder="Sub-Contractor Contract Number"
                  label="Select Sub-Contractor Contract Number"
                  options={this.state.purchaseOrderArray}
                  onChanged={this._poNumberChange}
                  selectedKey={this.state.poNumberID}
                />
                {/* <TextField label="Sub-Contractor Contract Number" onChange={this._poNumberChange} value={this.state.poNumber}></TextField> */}
              </div>
              <div className={styles.wdthlft}>
                <Dropdown placeholder="Select Acceptance code" label="Acceptance code" options={this.state.transmittalSettingsArray} selectedKey={this.state.transmittalSettingsId} onChanged={this._onTransmittalSettingsChange} required />
                <div style={{ color: "#dc3545", display: this.state.notransmittal }}>Please select valid Acceptance code</div>
              </div>
            </div>
            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <Dropdown placeholder="Select Revision Code" label="Revision Code" options={this.state.revisionSettingsArray} selectedKey={this.state.revisionCodingId} onChanged={this._onRevisionCodingChange} required />
                <div style={{ color: "#dc3545" }}>{this.validator.message("RevisionCode", this.state.revisionCodingId, "required")}{" "}</div>
              </div>
              <div className={styles.wdthlft} style={{ marginTop: "3%" }}>
                <Checkbox label="Increment Revision ? " boxSide="end" onChange={this._onIncrementRevisionChecked} checked={this.state.isIncrement} />
              </div>
            </div>
            <div className={styles.divrow}>
              <div style={{ width: "95%" }} >< TextField label="Comments" id="comments" value={this.state.comments} onChange={this._commentschange} multiline required autoAdjustHeight></TextField></div>
              <div><IconButton iconProps={AddIcon} title="Addindex" ariaLabel="Addindex" onClick={this._addindex} style={{ padding: "58px 0px 0px 45px" }} /></div>
            </div>
            <div style={{ color: "#dc3545", display: this.state.validComment }}>Please enter description</div>
            <table className={styles.tableModal} hidden={this.state.showGrid} >
              <tr style={{ background: "#f4f4f4" }}>
                <th style={{ padding: "5px 10px" }}>Slno</th>
                <th style={{ padding: "5px 10px" }}>DocumentIndex</th>
                <th style={{ padding: "5px 10px" }}>Sub-Contractor Document Number</th>
                <th style={{ padding: "5px 10px" }}>ReceivedDate</th>
                <th style={{ padding: "5px 10px" }}>Owner</th>
                <th style={{ padding: "5px 10px" }}>Comments</th>
                <th style={{ padding: "5px 10px" }}>Delete</th>
              </tr>
              {this.state.gridDocument.map((items, key) => {
                return (
                  <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                    <td style={{ padding: "5px 10px" }}>{key + 1}</td>
                    <td style={{ padding: "5px 10px" }}>{items.DocumentIndex} </td>
                    <td style={{ padding: "5px 10px", textAlign: "center" }}>{items.SubContractorNumber} </td>
                    <td style={{ padding: "5px 10px" }}>{items.ReceivedDate}</td>
                    <td style={{ padding: "5px 10px" }}>{items.Owner} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Comments}</td>
                    <td style={{ padding: "5px 10px" }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key, "ProjectDocuments")} /></td>
                  </tr>
                );
              })}

            </table>

            <div className={styles.subSection}>Additional Document</div>

            <div className={styles.divrow}>
              <div className={styles.wdthrgt}>
                <Label >Upload Document:</Label>
                <input type="file" name="externalFile" id="externalFile" onChange={(e) => this._uploadadditional(e)} ref={ref => this.myfileadditional = ref}></input>
              </div>
              <div className={styles.wdthlft}>
                <DatePicker label="Date"
                  value={this.state.externalDate}
                  onSelectDate={this._onDatePickerChange}
                  placeholder="Select a date"
                  formatDate={this._onFormatDate} /></div>
            </div>
            <div style={{ color: "#dc3545", display: this.state.uploadAdditionalDocumentError, marginLeft: "9px" }}>Sorry this document is unable to process due to unwanted characters.Please rename the document and try again.</div>
            <div className={styles.divrow}>
              <div style={{ width: "95%" }} >< TextField label="Comments" id="comments" value={this.state.externalComments} onChange={this._externalCommentsChange} multiline required autoAdjustHeight></TextField></div>
              <div><IconButton iconProps={AddIcon} title="Addindex" ariaLabel="Addindex" onClick={this._addexternalindex} style={{ padding: "58px 0px 0px 45px" }} /></div>
            </div>
            <div style={{ color: "#dc3545", display: this.state.validAdditionalComment }}>Please enter description</div>
            <table className={styles.tableModal} hidden={this.state.showExternalGrid} >
              <tr style={{ background: "#f4f4f4" }}>
                <th style={{ padding: "5px 10px" }} >Slno</th>
                <th style={{ padding: "5px 10px" }}>Document Name</th>
                <th style={{ padding: "5px 10px" }}>ReceivedDate</th>
                <th style={{ padding: "5px 10px" }}>Comments</th>
                <th style={{ padding: "5px 10px" }}>Delete</th>
              </tr>
              {this.state.gridExternalDocument.map((items, key) => {
                return (
                  <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                    <td style={{ padding: "5px 10px" }}>{key + 1}</td>
                    <td style={{ padding: "5px 10px" }}>{items.DocName} </td>
                    <td style={{ padding: "5px 10px" }}>{items.ExternalDate}</td>
                    <td style={{ padding: "5px 10px" }}>{items.Comments}</td>
                    <td style={{ padding: "5px 10px" }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key, "AdditionalDocuments")} /></td>
                  </tr>
                );
              })}

            </table>

            <div> {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''} </div>
            <div style={{ display: this.state.checksend }}><Spinner label={'Please wait  transmittal is getting ready...'} /></div>
            <DialogFooter>

              <div className={styles.rgtalign}>
                <div style={{ fontStyle: "italic", fontSize: "12px" }}><span style={{ color: "red", fontSize: "23px" }}>*</span>fields are mandatory </div>
              </div>
              <div className={styles.rgtalign} >
                <PrimaryButton id="b2" className={styles.btn} onClick={this._saveAsDraft} disabled={this.state.submitDisable}>Save as Draft</PrimaryButton >
                <PrimaryButton id="b2" className={styles.btn} onClick={this._submit} disabled={this.state.submitDisable}>Submit</PrimaryButton >
                <PrimaryButton id="b1" className={styles.btn} onClick={this._onCancel}>Cancel</PrimaryButton >
              </div>

            </DialogFooter>
            {/* {/ Cancel Dialog Box /} */}
            <div style={{ display: this.state.cancelConfirmMsg }}>
              <div>
                <Dialog
                  hidden={this.state.confirmDialog}
                  dialogContentProps={this.dialogContentProps}
                  onDismiss={this._dialogCloseButton}
                  styles={this.dialogStyles}
                  modalProps={this.modalProps}>
                  <DialogFooter>
                    <PrimaryButton onClick={this._confirmYesCancel} text="Yes" />
                    <DefaultButton onClick={this._confirmNoCancel} text="No" />
                  </DialogFooter>
                </Dialog>
              </div>
            </div>
            {/* Delete Dialog Box  */}
            <div style={{ display: this.state.deleteConfirmMsg }}>
              <div>
                <Dialog
                  hidden={this.state.confirmDeleteDialog}
                  dialogContentProps={this.dialogDeleteProps}
                  onDismiss={this._dialogCloseButton}
                  styles={this.dialogStyles}
                  modalProps={this.modalProps}>
                  <DialogFooter>
                    <PrimaryButton onClick={() => this._confirmDeleteItem(this.state.tempDocIndexIDForDelete, "item", this.keyForDelete)} text="Yes" />
                    <DefaultButton onClick={this._confirmNoCancel} text="No" />
                  </DialogFooter>
                </Dialog>
              </div>
            </div>
          </div>
        </div>
        <div style={{ display: this.state.accessDeniedMsgBar }}>

          {this.state.statusMessage.isShowMessage ?
            <MessageBar
              messageBarType={this.state.statusMessage.messageType}
              isMultiline={false}
              dismissButtonAriaLabel="Close"
            >{this.state.statusMessage.message}</MessageBar>
            : ''}
        </div>
      </div>
    );
  }
}
