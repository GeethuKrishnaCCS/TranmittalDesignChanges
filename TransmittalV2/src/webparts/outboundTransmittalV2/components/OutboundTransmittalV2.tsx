import * as React from 'react';
import styles from './OutboundTransmittalV2.module.scss';
import { IOutboundTransmittalV2Props, IOutboundTransmittalV2State } from '../Interfaces/IOutboundTransmittalV2Props';
import { Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, FontWeights, IChoiceGroupOption, IDropdownOption, IDropdownStyles, IIconProps, IconButton, Label, MessageBar, Modal, PrimaryButton, Spinner, SpinnerSize, TextField, getTheme, mergeStyleSets } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { MultiSelect } from 'react-multi-select-component';
import SimpleReactValidator from 'simple-react-validator';
import { OBService } from '../Services/OBService';
import * as moment from 'moment';
import replaceString from 'replace-string';
import * as _ from 'lodash';
import { add } from 'lodash';
import Select from 'react-select';
import { Accordion, AccordionItem, AccordionItemButton, AccordionItemHeading, AccordionItemPanel } from 'react-accessible-accordion';
import CustomFileInput from './CustomFileInput';
import { DragDropFiles } from '@pnp/spfx-controls-react/lib/DragDropFiles';
import { IHttpClientOptions, HttpClient } from '@microsoft/sp-http';

export default class OutboundTransmittalV2 extends React.Component<IOutboundTransmittalV2Props, IOutboundTransmittalV2State, {}> {
  private validator: SimpleReactValidator;
  private _Service: OBService;
  //private reqWeb = Web(window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.hubSiteUrl);
  private emailsSelectedTo: any[] = [];
  private emailsSelectedCC: any[] = [];
  private contactToDisplay: any[] = [];
  private contactCCDisplay: any[] = [];
  private sortedArray: any[] = [];
  private transmittalID: string;
  private keyForDelete: any;
  private typeForDelete;
  private myfileadditional: { value: string; };
  constructor(props: IOutboundTransmittalV2Props) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      currentUser: null,
      totalNoOfFiles: "",
      transmitToKey: "",
      transmitTo: "",
      hideCustomer: "none",
      hideSubContractor: "none",
      selectedKeys: "",
      isCalloutVisible: true,
      customerContChBx: "",
      toggleMultiline: false,
      commentMultiline: false,
      showGrid: true,
      showExternalGrid: true,
      contactsForSearch: [],
      //dropdowns
      subContractorKey: "",
      subContractor: "",
      transmitForKey: "",
      transmitFor: "",
      subContractorItems: [],
      hideUnlockButton: "none",
      //
      projectInformation: [],
      projectName: "",
      customerName: "",
      transmittalNo: "none",
      contacts: [],
      isChecked: false,
      //contactsTo
      selectedContactsTo: "",
      selectedContactsCC: "",
      CC: "No",
      TypeOFDelete: "",
      transmitForItems: [],
      approvalLifeCycle: "",
      projectNumber: "",
      dueDate: null,
      hideDueDate: true,
      //serachDocuments
      searchDocuments: [],
      items: [],
      searchDiv: "none",
      searchText: "",
      //grid
      tempArrayForPublishedDocumentGrid: [],
      tempArrayForExternalDocumentGrid: [],
      publishDocumentsItemsForGrid: [],
      comments: "",
      itemsForGrid: [],
      itemsForExternalGrid: [],
      fileSize: "",
      fileSizeDiv: true,
      //previewDiv
      previewDiv: true,
      showReviewModal: false,
      externalComments: "",
      notes: "",
      transmittalType: "",
      incrementSequenceNumber: "",
      transmittalID: "",
      customerId: null,
      outboundTransmittalHeaderId: "",
      contactCCDisplay: "",
      contactToDisplay: "",
      publishedDocumentArray: [],
      documentSelect: "",
      documentSelectedDiv: true,
      dropDownReadonly: false,
      checkedItems: new Map(),
      transmittalTypekey: "",
      cancelConfirmMsg: "none",
      confirmCancelDialog: true,
      confirmDeleteDialog: true,
      currentOutboundDetailItem: [],
      currentOutboundAdditionalItem: [],
      sendAsSharedFolder: false,
      recieveInSharedFolder: false,
      sendAsMultipleFolder: false,
      hideButtonAfterSubmit: "",
      recallConfirmMsgDiv: "none",
      recallConfirmMsg: true,
      outboundRecallConfirmation: "",
      approvalRequired: false,
      notificationPreference: "",
      tempDocIndexIDForDelete: null,
      subContractorLabel: "none",
      subContractorDrpDwn: "",
      deleteConfirmation: "none",
      projectDocumentSelectKey: "",
      hideGridAddButton: false,
      selectedContactsToDisplayName: "",
      selectedContactsCCDisplayName: "",
      sourceDocumentItem: null,
      sendAsMultipleEmailCheckBoxDiv: "none",
      transmitTypeForDocument: "none",
      transmitTypeForLetter: "none",
      transmitTypeForDefault: "none",
      webpartView: "none",
      loaderDisplay: "none",
      accessDeniedMsgBar: "none",
      normalMsgBar: "none",
      fileSizeDivForRebind: "none",
      spinnerDiv: "none",
      uploadDocumentError: "none",
      selectedContactsTo1: null,
      selectedContactsToCCRebind: null,
      dueDateForBindingApprovalLifeCycle: null,
      contractNumber: "",
      legalId: "",
      coverLetterConfirmBox: "",
      coverLetterDialog: true,
      coverLetterNeeded: false,
      //people picker internal contacts
      internalCCContacts: [],
      internalCCContactsDisplayName: "",
      internalCCContactsDisplayNameForPreview: "",
      internalContactsEmail: "",
      settingsListArray: [],
      selectedVendor: [],
      searchContactsTo: [],
      selectedContactsToName: [],
      searchContactsCC: [],
      selectedContactsCCName: [],
      divForToAndCC: "none",
      divForToAndCCSearch: "",
      selectedDocuments: [],
      settingsListsItemsArray: [],
      documentFilters: []
    };
    this._Service = new OBService(this.props.context, window.location.protocol + "//" + window.location.hostname + this.props.hubSiteUrl);
    this._drpdwnTransmitTo = this._drpdwnTransmitTo.bind(this);
    this._drpdwnSubContractor = this._drpdwnSubContractor.bind(this);
    this._currentUser = this._currentUser.bind(this);
    this._showProjectDocumentGrid = this._showProjectDocumentGrid.bind(this);
    this._showExternalGrid = this._showExternalGrid.bind(this);
    this._hideGrid = this._hideGrid.bind(this);
    this._confirmAndSendBtnClick = this._confirmAndSendBtnClick.bind(this);
    this.projectInformation = this.projectInformation.bind(this);
    //this._queryParamGetting = this._queryParamGetting.bind(this);
    this._transmitForBind = this._transmitForBind.bind(this);
    this._loadPublishDocuments = this._loadPublishDocuments.bind(this);
    this._onDocumentClick = this._onDocumentClick.bind(this);
    this._onPreviewBtnClick = this._onPreviewBtnClick.bind(this);
    this._loadSourceDocuments = this._loadSourceDocuments.bind(this);
    this.itemDeleteFromGrid = this.itemDeleteFromGrid.bind(this);
    this.itemDeleteFromExternalGrid = this.itemDeleteFromExternalGrid.bind(this);
    // this._onSaveAsDraftBtnClick = this._onSaveAsDraftBtnClick.bind(this);
    this._trannsmittalIDGeneration = this._trannsmittalIDGeneration.bind(this);
    this._onTransmitType = this._onTransmitType.bind(this);
    this._transmittalSequenceNumber = this._transmittalSequenceNumber.bind(this);
    this.handleChange = this.handleChange.bind(this);
    this.getCheckboxesValue = this.getCheckboxesValue.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
    this._confirmNoCancel = this._confirmNoCancel.bind(this);
    // this.bindOutboundTransmittalSavedData = this.bindOutboundTransmittalSavedData.bind(this);
    this._openDeleteConfirmation = this._openDeleteConfirmation.bind(this);
    this.triggerOutboundTransmittal = this.triggerOutboundTransmittal.bind(this);
    this._recallTransmittalConfirmation = this._recallTransmittalConfirmation.bind(this);
    this._userMessageSettings = this._userMessageSettings.bind(this);
    // this._recallSubmit = this._recallSubmit.bind(this);
    this._confirmDeleteItem = this._confirmDeleteItem.bind(this);
    this._forCalculatingSize = this._forCalculatingSize.bind(this);
    this._loadSourceDocumentsForLetter = this._loadSourceDocumentsForLetter.bind(this);
    //  this._LAUrlGettingForPermission = this._LAUrlGettingForPermission.bind(this);
    //this.triggerProjectPermissionFlow = this.triggerProjectPermissionFlow.bind(this);
    //this._LAUrlGettingForRecall = this._LAUrlGettingForRecall.bind(this);
    // this.triggerProjectRecall = this.triggerProjectRecall.bind(this);
    this._coverLetterNeeded = this._coverLetterNeeded.bind(this);
    this.setSelectedContactsTo = this.setSelectedContactsTo.bind(this);
    this.setSelectedContactsCC = this.setSelectedContactsCC.bind(this);
    this.setSelectedDocuments = this.setSelectedDocuments.bind(this);
    this.loadSettingsList = this.loadSettingsList.bind(this);
  }
  public render(): React.ReactElement<IOutboundTransmittalV2Props> {
    const {
      hasTeamsContext,
    } = this.props;
    const TransmitTo: IDropdownOption[] = [
      { key: '1', text: 'Customer' },
      { key: '2', text: 'Sub-Contractor' },
    ];
    const options: IChoiceGroupOption[] = [
      { key: 'Document', text: 'Document' },
      { key: 'Letter', text: 'Letter' },
    ];
    const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: "50%" } };
    // const AddIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
    const DeleteIcon: IIconProps = { iconName: 'Delete' };
    const CancelIcon: IIconProps = { iconName: 'Cancel' };
    const theme = getTheme();
    const contentStyles = mergeStyleSets({
      container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',


      },
      header: [
        theme.fonts.xLargePlus,
        {
          flex: '1 1 auto',
          //borderTop: `4px solid ${theme.palette.themePrimary}`,
          color: theme.palette.neutralPrimary,
          display: 'flex',
          alignItems: 'center',
          fontWeight: FontWeights.semibold,
          padding: '12px 12px 14px 284px',
        },
      ],
      header1: [

        theme.fonts.xLargePlus,
        {
          flex: '1 1 auto',
          // borderTop: `4px solid ${theme.palette.themePrimary}`,
          color: theme.palette.neutralPrimary,
          display: 'flex',
          alignItems: 'center',
          fontWeight: FontWeights.semibold,
          padding: '10px 20px',
        },
      ],
      body: {
        flex: '4 4 auto',
        padding: '0 20px 20px ',
        overflowY: 'hidden',

        selectors: {
          p: { margin: '14px 0' },
          'p:first-child': { marginTop: 0 },
          'p:last-child': { marginBottom: 0 },
        },
      },
    });
    const iconButtonStyles = {
      root: {
        color: theme.palette.neutralPrimary,
        marginLeft: 'auto',
        marginTop: '4px',
        marginRight: '2px',
      },
      rootHovered: {
        color: theme.palette.neutralDark,
      },
    };
    const DownIcon: IIconProps = { iconName: 'ChevronDown' };

    return (
      <section className={`${styles.outboundTransmittalV2} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <div>

            <div className={styles.outboundTransmittalV2} >
              <Label className={styles.align}>{this.props.description}</Label>
              <div style={{ marginLeft: "522px" }} />

              <div className={styles.outSideBorder}>
                <div className={styles.transmittalNo}>
                  <Label style={{ display: this.state.transmittalNo }}>Transmittal No :  {this.state.transmittalNo}	</Label></div>
                <div style={{ width: "50%", marginLeft: "450px" }}>
                  <Label>Project :   {this.state.projectName}</Label></div>
              </div>
              <div className={styles.border}>
                <div className={styles.row}>
                  <div style={{ display: "flex", marginBottom: "10px" }}>
                    <div style={{ marginLeft: "10px", display: "flex", width: "100%" }}>
                      <Dropdown id="t3"
                        required={true}
                        selectedKey={this.state.transmitToKey}
                        placeholder="Select an option"
                        options={TransmitTo}
                        onChange={this._drpdwnTransmitTo}
                        style={{ width: "100%" }}
                        label="Transmit To"
                        disabled={this.state.dropDownReadonly} />
                    </div>
                    <div style={{ display: this.state.hideCustomer, width: "50%" }}>
                      <div style={{ display: "flex", marginTop: "22px" }}>
                        <Label>Customer : </Label>
                        <Label style={{ fontWeight: "bold", paddingLeft: "5px" }}>  {this.state.customerName}</Label>
                      </div>
                    </div>
                    {this.state.hideSubContractor === "" &&
                      <div style={{ width: "100%" }}>
                        <div style={{ display: "flex", marginLeft: "123px", marginTop: "3px" }}>
                          <Label required>Sub-Contractor : </Label>
                          <Select
                            placeholder="Select Sub-Contractor"
                            isMulti={false}
                            options={this.state.subContractorItems}
                            onChange={this._drpdwnSubContractor.bind(this)}
                            isSearchable={true}
                            value={this.state.subContractorKey}
                            maxMenuHeight={150}
                            className={styles.subContractorDropDwn}
                          />

                          <Label style={{ marginLeft: "10px", display: this.state.subContractorLabel }}>{this.state.subContractor} </Label>
                        </div>
                        <div style={{ color: "#dc3545", marginLeft: "123px" }}>
                          {this.validator.message("subContractor", this.state.subContractorKey, "required")}{" "}</div>
                      </div>}

                  </div>
                  <div style={{ color: "#dc3545" }}>{this.validator.message("transmitTo", this.state.transmitToKey, "required")}{" "}</div>
                  <hr />
                  <div >
                    <div style={{ marginBottom: "10px" }} className={styles.borderForToCC}>
                      <span className={styles.span} />
                      <div style={{ width: "97%", display: this.state.divForToAndCCSearch }}>
                        <label style={{ fontWeight: "bold", }}>To</label>
                        <MultiSelect options={this.state.contactsForSearch} value={this.state.selectedContactsToName}
                          onChange={this.setSelectedContactsTo}
                          labelledBy="To" hasSelectAll={true} />
                        <div style={{ color: "#dc3545" }}>{this.validator.message("selectedContactsTo", this.state.selectedContactsTo, "required")}{" "}</div>
                      </div>
                      <div style={{ width: "195%", display: this.state.divForToAndCC }}>
                        <Dropdown
                          placeholder="Select To contacts"
                          label="To"
                          defaultSelectedKeys={this.state.selectedContactsTo1}
                          multiSelect
                          multiSelectDelimiter={","}
                          options={this.state.contacts}
                          styles={dropdownStyles}
                          //onChange={this._onDrpdwnCntact}
                          title="To"
                        />
                        <div style={{ color: "#dc3545" }}>{this.validator.message("selectedContactsTo", this.state.selectedContactsTo, "required")}{" "}</div>
                      </div>
                      <span className={styles.span} />
                      <div style={{ width: "97%", display: this.state.divForToAndCCSearch }}>
                        <label style={{ fontWeight: "bold", }}>CC</label>
                        <MultiSelect options={this.state.contactsForSearch} value={this.state.selectedContactsCCName}
                          onChange={this.setSelectedContactsCC}
                          labelledBy="CC" hasSelectAll={true} />
                      </div>
                      <div style={{ width: "195%", display: this.state.divForToAndCC }}>
                        <Dropdown
                          placeholder="Select CC contacts"
                          label="CC "
                          defaultSelectedKeys={this.state.selectedContactsToCCRebind}
                          multiSelect
                          multiSelectDelimiter={","}
                          options={this.state.contacts}
                          styles={dropdownStyles}
                        // onChange={this._onDrpdwnCCContact}
                        />
                      </div>
                      <div style={{ width: "98%", fontWeight: "bold" }}>
                        <PeoplePicker
                          context={this.props.context}
                          titleText="Internal CC"
                          personSelectionLimit={20}
                          groupName={""} // Leave this blank in case you want to filter from all users
                          showtooltip={true}
                          required={true}
                          disabled={false}
                          ensureUser={true}
                          onChange={(items) => this._selectedInternalCCContacts(items)}
                          principalTypes={[PrincipalType.User]}
                          resolveDelay={1000}
                          defaultSelectedUsers={this.state.internalCCContactsDisplayName} />
                      </div>
                    </div>
                  </div>
                  {/* choice groups */}
                  <div  >
                    <div style={{ display: "flex" }}>
                      <div style={{ marginLeft: "26px", marginRight: "15px", display: this.state.transmitTypeForDocument }}>
                        <ChoiceGroup options={options}
                          onChange={this._onTransmitType}
                          label="Select any" required={true} defaultSelectedKey={'Document'} disabled={true} />
                      </div>
                      <div style={{ marginLeft: "26px", marginRight: "15px", display: this.state.transmitTypeForLetter }}>
                        <ChoiceGroup options={options}
                          onChange={this._onTransmitType}
                          label="Select any" required={true} defaultSelectedKey={'Letter'} disabled={true} />
                      </div>
                      <div style={{ marginLeft: "26px", marginRight: "15px", display: this.state.transmitTypeForDefault }}>
                        <ChoiceGroup options={options}
                          onChange={this._onTransmitType}
                          label="Select any" required={true} />
                      </div>
                      <div style={{ marginLeft: "150px", marginTop: "9px" }}>
                        <Label>Check if cover letter needed</Label>
                        <div className={styles.mt1}><Checkbox label="Cover Letter" title="Check if cover letter needed or not."
                          onChange={this._coverLetterNeeded}
                          checked={this.state.coverLetterNeeded} /></div>
                      </div>
                      <div style={{ marginLeft: "90px", marginTop: "9px" }}>
                        <Label>Select email type</Label>
                        <div className={styles.mt1}><Checkbox label="Send and Receive as shared folder"
                          onChange={this._onSendAsSharedFolder}
                          checked={this.state.sendAsSharedFolder} /></div>
                        <div className={styles.mt1} style={{ display: this.state.sendAsMultipleEmailCheckBoxDiv }}><Checkbox label="Send as multiple emails"
                          onChange={this._onSendAsMultipleFolder}
                          checked={this.state.sendAsMultipleFolder} /></div>
                      </div>

                    </div>
                    {/* transmittal type validationdiv */}
                    <div style={{ color: "#dc3545", marginLeft: "26px" }}>{this.validator.message("transmittalType", this.state.transmittalType, "required")}{" "}</div>
                  </div>
                  <hr />
                  {/* Notes */}
                  <div style={{ marginLeft: "9px" }}>
                    <TextField label="Notes" multiline placeholder="" value={this.state.notes}
                      onChange={this.notesOnchange}
                      style={{ marginLeft: "20px", width: "290px" }}
                    />
                  </div>
                  <hr />

                  {/* filesizeDiv */}
                  {this.state.itemsForGrid.length > 0 &&
                    <div hidden={this.state.fileSizeDiv} style={{ float: "right", color: (Number(this.state.fileSize) >= 25) ? "Red" : "Green" }}>Size : [{(this.state.fileSize < 1) ? this.state.fileSize + " MB" : this.state.fileSize + " MB"}]</div>
                  }
                  {/* project documents */}
                  {this.state.settingsListArray.length > 0 &&
                    <div style={{ display: "flex", marginRight: "10px" }}>
                      {this.state.settingsListArray.map((item, index) => {
                        return (
                          <div style={{ marginRight: "10px" }}>
                            <Dropdown id={item.Title}
                              required={true}
                              selectedKey={item.selectedKey}
                              placeholder="Select an option"
                              options={this.state.settingsListsItemsArray[item.Title] ? this.state.settingsListsItemsArray[item.Title] : []}
                              onChange={(_, e) => this.handleSettingsListItemsChange(index, e, item.Title)}
                              label={item.Title}
                              disabled={this.state.dropDownReadonly} /></div>);
                      })
                      }
                    </div>
                  }
                  <div style={{ padding: "12px 0 12px 12px" }}>
                    <div style={{ display: "block" }}>
                      <div hidden={this.state.documentSelectedDiv} style={{ fontWeight: "bold", color: "Red" }}> {this.state.documentSelect}</div>
                      <Label>Project Documents</Label>
                      <MultiSelect
                        options={this.state.searchDocuments}
                        value={this.state.selectedDocuments}
                        onChange={this.setSelectedDocuments}
                        labelledBy="Select"
                        hasSelectAll={true}
                      />
                      <div style={{ color: "#dc3545", marginLeft: "123px" }}>
                        {this.validator.message("projectDocuments", this.state.projectDocumentSelectKey, "required")}{" "}
                      </div>
                    </div>
                  </div>
                  {/* projectDocumentGrid */}
                  {this.state.itemsForGrid.length > 0 &&
                    <div style={{ width: '100%', }}>
                      <table className={styles['custom-table']} hidden={this.state.showGrid}>
                        <tr style={{ textAlign: "left" }}>
                          <th >Slno</th>
                          <th >Document Name</th>
                          <th >Revision No</th>
                          <th style={{ display: (this.state.transmitTo === "Customer") ? "" : "none" }}>Customer Document No</th>
                          <th style={{ display: (this.state.transmitTo === "Sub-Contractor") ? "" : "none" }}>SubContractor Document No</th>
                          <th style={{ display: (this.state.transmitTo === "Sub-Contractor") ? "none" : "none" }}>Acceptance Code</th>
                          <th >Size (in MB)</th>
                          <th >Transmit For</th>
                          <th >Due Date</th>
                          <th >Comments</th>
                          <th style={{ display: this.state.hideButtonAfterSubmit }}>Delete</th>
                        </tr>
                        {this.state.itemsForGrid.map((items, key) => {
                          return (
                            <tr key={key} style={{ textAlign: "left" }}>
                              <td >{key + 1}</td>
                              <td >{items.documentName} </td>
                              <td >{items.revision} </td>
                              <td style={{ display: (this.state.transmitTo === "Customer") ? "" : "none" }}>{items.customerDocumentNo} </td>
                              <td style={{ display: (this.state.transmitTo === "Sub-Contractor") ? "" : "none" }}>{items.subcontractorDocumentNo} </td>
                              <td style={{ display: (this.state.transmitTo === "Sub-Contractor") ? "none" : "none" }}>{items.acceptanceCodeTitle}</td>
                              <td >{items.fileSizeInMB}</td>
                              <td >
                                <Dropdown id={key + "TransmittedFor"}
                                  selectedKey={items.TransmittedFor}
                                  placeholder="Select an option"
                                  options={this.state.transmitForItems}
                                  onChange={(_, e) => this._drpdwnTransmitFor(key, e, items)}
                                /> </td>
                              <td >
                                <DatePicker
                                  value={items.dueDate}
                                  hidden={this.state.hideDueDate}
                                  //onse={(_, e) =>this._dueDatePickerChange(key, e)}
                                  minDate={this.state.dueDateForBindingApprovalLifeCycle}
                                  placeholder="Select a date..."
                                  ariaLabel="Select a date"
                                  formatDate={this._onFormatDate} /></td>
                              <td >  <TextField autoComplete="off" multiline
                                placeholder="" value={items.Comments}
                                onChange={(_, e) => this.onCommentChange(key, e)}
                              /></td>
                              <td style={{ display: this.state.hideButtonAfterSubmit }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key, "ProjectDocuments")} /></td>
                            </tr>
                          );
                        })}
                      </table>
                    </div>
                  }
                  <hr style={{ marginTop: "20px" }} />
                  <Accordion allowZeroExpanded >
                    <AccordionItem >
                      <AccordionItemHeading>
                        <AccordionItemButton >
                          <Label ><IconButton iconProps={DownIcon} />External Documents</Label>
                        </AccordionItemButton>
                      </AccordionItemHeading>
                      <AccordionItemPanel>
                        <div className={styles.divrow} >
                          <div className={styles.wdthfrst}>
                            <CustomFileInput onChange={this.uploadFile} key={1} />

                          </div>
                          <div className={`${styles.dragDropContainer} ${styles.wdthThirdColm}`}>
                            <DragDropFiles
                              dropEffect="copy"
                              enable={true}
                              onDrop={this._getDropFiles}
                              iconName="Upload"
                              labelMessage="Upload Files"
                            >
                              Drag and drop here...
                            </DragDropFiles></div>
                        </div>
                        {/* <div style={{ color: "#dc3545", display: this.state.uploadDocumentError, marginLeft: "9px" }}>Sorry this document is unable to process due to unwanted characters.Please rename the document and try again.</div>
                      <div style={{ width: "100%", display: "flex" }}>
                        <div style={{ width: "100%", padding: "10px 7px 10px 9px" }}> <TextField required={true} value={this.state.externalComments} placeholder="" 
                        // onChange={this.onCommentExternalChange} 
                        />
                        </div>
                        <div style={{ width: "5%", padding: "10px 7px 10px 9px", display: this.additionalDivHide }}>
                          <IconButton iconProps={AddIcon} title="Add External Documents" ariaLabel="Add" 
                          //onClick={this._showExternalGrid} 
                          style={{ marginTop: "-4px", display: this.state.hideButtonAfterSubmit }} />
                        </div>
                      </div>
                      <div style={{ color: "#dc3545", marginLeft: "123px" }}>{this.validator.message("externalcomments", this.state.externalComments, "required")}{" "}</div> */}

                      </AccordionItemPanel>
                    </AccordionItem>
                  </Accordion>
                  <div  >
                    <table className={styles['custom-table']} hidden={this.state.showExternalGrid} >
                      <tr style={{ textAlign: "left" }}>
                        <th >Slno</th>
                        <th>Document Name</th>
                        <th>Size (in MB)</th>
                        <th >Comments</th>
                        <th style={{ display: this.state.hideButtonAfterSubmit }}>Delete</th>
                      </tr>
                      {this.state.itemsForExternalGrid.map((items, key) => {
                        return (
                          <tr style={{ textAlign: "left" }}>
                            <td>{key + 1}</td>
                            <td >{items.documentName}</td>
                            <td>{items.fileSizeInMB}</td>
                            <td >
                              <TextField autoComplete="off" multiline
                                placeholder="" value={items.externalComments}
                                onChange={(_, e) => this.onExternalCommentChange(key, e)} />
                            </td>
                            <td style={{ display: this.state.hideButtonAfterSubmit }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key, "AdditionalDocuments")} /></td>
                          </tr>
                        );
                      })}
                    </table>
                  </div>
                </div>
                <div style={{ display: this.state.normalMsgBar }}>
                  {/* Show Message bar for Notification*/}
                  {this.state.statusMessage.isShowMessage ?
                    <MessageBar
                      messageBarType={this.state.statusMessage.messageType}
                      isMultiline={false}
                      dismissButtonAriaLabel="Close"
                    >{this.state.statusMessage.message}</MessageBar>
                    : ''}
                </div>
                <Spinner color="Blue" size={SpinnerSize.large} style={{ display: this.state.spinnerDiv, marginBottom: "10px" }} label={'Transmittal Sending.Please Wait...'} />
                <div style={{ display: "flex", padding: "10px 0px 12px 2px", float: "right", }}>
                  <PrimaryButton text="Save as draft" style={{ marginLeft: "auto", marginRight: "11px", display: this.state.hideButtonAfterSubmit }}
                  //onClick={() => this._onSaveAsDraftBtnClick()} 
                  />
                  <PrimaryButton text="Preview" style={{ marginRight: "11px", marginLeft: "auto" }}
                    onClick={this._onPreviewBtnClick}
                  />
                  <PrimaryButton text="Confirm & Send" style={{ marginRight: "11px", marginLeft: "auto", display: this.state.hideButtonAfterSubmit }}
                    onClick={this._confirmAndSendBtnClick}
                  />
                  <PrimaryButton text="Recall" style={{ marginRight: "11px", marginLeft: "auto", display: this.state.hideUnlockButton }} onClick={this._recallTransmittalConfirmation} />
                  <PrimaryButton text="Cancel" style={{ marginLeft: "auto" }} onClick={this._hideGrid} />
                </div>
              </div>
              <div />
              {/* div for preview  */}
              <div hidden={this.state.previewDiv}>
                <Modal
                  isOpen={this.state.showReviewModal}
                  onDismiss={this._closeModal}
                  containerClassName={contentStyles.container}>
                  <div style={{ marginLeft: "96%" }}>
                    <IconButton
                      iconProps={CancelIcon}
                      ariaLabel="Close popup modal"
                      onClick={this._closeModal}
                      styles={iconButtonStyles}
                    />
                  </div>

                  <div className={styles.wrap} style={{ backgroundColor: this.props.modalBGColor }}>
                    <div className={styles.borderStyle}><div className={styles.title}>{this.state.projectName}</div></div>

                    <div className={styles.wrapSection}>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Transmittal No:&nbsp; </span>{this.state.transmittalNo}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Customer Contract No:&nbsp; </span>{this.state.contractNumber}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Date: &nbsp;</span>{moment.utc(new Date()).format("DD/MM/YYYY")}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Transmit to: &nbsp;</span>{this.state.transmitTo}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Transmitted by: &nbsp; </span>{this.props.context.pageContext.user.displayName}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Send to:&nbsp; </span>{this.state.selectedContactsToDisplayName}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Send CC: &nbsp;</span>{this.state.selectedContactsCCDisplayName}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Internal CC: &nbsp;</span>{this.state.internalCCContactsDisplayNameForPreview}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Total no of files:&nbsp; </span>{this.state.totalNoOfFiles}</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Total Size:&nbsp; </span>{this.state.fileSize}MB</div>
                      <div className={styles.wrapIitem} style={{ width: "50%" }}> <span>Cover Letter Attached:&nbsp; </span>{this.state.coverLetterNeeded === true ? "Yes" : "No"}</div>
                      <div className={styles.wrapIitem} style={{ width: "100%", marginTop: "15px" }}><span>Note:&nbsp; </span>{this.state.notes}</div>
                    </div>

                    <div className={styles.wrapTable}>
                      <div className={styles.w100}><div className={styles.subtitle}>Project Documents</div></div>
                      <div className={styles.overflow}>
                        <div className={styles.divTable}>
                          <div className={styles.divTableBody}>
                            <div className={styles.divTableRow}>
                              <div className={styles.divTableCell}>Slno</div>
                              <div className={styles.divTableCell}>Document Name</div>
                              <div className={styles.divTableCell}>Revision</div>
                              <th className={styles.divTableCell} style={{ display: (this.state.transmitTo === "Customer") ? "" : "none" }}>Customer Document No</th>
                              <th className={styles.divTableCell} style={{ display: (this.state.transmitTo === "Sub-Contractor") ? "" : "none" }}>SubContractor Document No</th>
                              <div className={styles.divTableCell} style={{ display: this.state.transmitTo === "Sub-Contractor" ? "none" : "none" }}>AcceptanceCode</div>
                              <div className={styles.divTableCell}>Size(in MB)</div>
                              <div className={styles.divTableCell}>Transmit for</div>
                              <div className={styles.divTableCell}>Due date</div>
                              <div className={styles.divTableCell}>Comments</div>
                            </div>
                            {this.state.itemsForGrid.map((items, key) => {
                              return (
                                <div className={styles.divTableRow}>
                                  <div className={styles.divTableCell}>&nbsp;{key + 1}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.documentName}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.revision}</div>
                                  <td className={styles.divTableCell} style={{ display: (this.state.transmitTo === "Customer") ? "" : "none" }}>{items.customerDocumentNo} </td>
                                  <td className={styles.divTableCell} style={{ display: (this.state.transmitTo === "Sub-Contractor") ? "" : "none" }}>{items.subcontractorDocumentNo} </td>
                                  <div className={styles.divTableCell} style={{ display: (this.state.transmitTo === "Sub-Contractor") ? "none" : "none" }}>&nbsp;{items.acceptanceCodeTitle}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.fileSizeInMB}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.TransmittedFor}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.DueDate}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.comments}</div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                      {/* <div className={styles.textright} style={{width: "100%",marginTop:"0px"}}>Total Size: <span>[Size]</span></div> */}
                    </div>
                    <div className={styles.wrapTable}>
                      <div className={styles.w100}><div className={styles.subtitle}>Additional Documents</div></div>
                      <div className={styles.overflow}>
                        <div className={styles.divTable}>
                          <div className={styles.divTableBody}>
                            <div className={styles.divTableRow}>
                              <div className={styles.divTableCell}>Slno</div>
                              <div className={styles.divTableCell}>Document Name</div>
                              <div className={styles.divTableCell}>Size(in MB)</div>
                              <div className={styles.divTableCell}>Comments</div>
                            </div>
                            {this.state.itemsForExternalGrid.map((items, key) => {
                              return (
                                <div className={styles.divTableRow}>
                                  <div className={styles.divTableCell}>&nbsp;{key + 1}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.documentName}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.fileSizeInMB}</div>
                                  <div className={styles.divTableCell}>&nbsp;{items.externalComments}</div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                      <div className={styles.textright} style={{ width: "100%", marginTop: "0px" }}><span>Total Size: &nbsp;</span>{this.state.fileSize}MB</div>
                    </div>

                  </div>
                </Modal>
              </div>
              {/* Delete Dialog Box */}
              <div style={{ display: this.state.deleteConfirmation }}>
                <div>
                  <Dialog
                    hidden={this.state.confirmDeleteDialog}
                    dialogContentProps={this.dialogContentProps}
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
              {/* Recall Dialog Box */}
              <div style={{ display: this.state.recallConfirmMsgDiv }}>
                <div>
                  <Dialog
                    hidden={this.state.recallConfirmMsg}
                    dialogContentProps={this.dialogContentRecallProps}
                    onDismiss={this._dialogCloseButton}
                    styles={this.dialogStyles}
                    modalProps={this.modalProps}>
                    <DialogFooter>
                      <PrimaryButton
                        //onClick={this._recallSubmit}
                        text="Yes" />
                      <DefaultButton onClick={this._confirmNoCancel} text="No" />
                    </DialogFooter>
                  </Dialog>
                </div>
              </div>
              {/* Cancel Dialog Box */}
              <div style={{ display: this.state.cancelConfirmMsg }}>
                <div>
                  <Dialog
                    hidden={this.state.confirmCancelDialog}
                    dialogContentProps={this.dialogCancelContentProps}
                    onDismiss={this._dialogCloseButton}
                    styles={this.dialogStyles}
                    modalProps={this.modalProps}>
                    <DialogFooter>
                      <PrimaryButton onClick={() => this._cancelConfirmYes()} text="Yes" />
                      <DefaultButton onClick={this._confirmNoCancel} text="No" />
                    </DialogFooter>
                  </Dialog>
                </div>
              </div>
            </div>
            <div style={{ display: this.state.accessDeniedMsgBar }}>
              {/* Show Message bar for Notification*/}
              {this.state.statusMessage.isShowMessage ?
                <MessageBar
                  messageBarType={this.state.statusMessage.messageType}
                  isMultiline={false}
                  dismissButtonAriaLabel="Close"
                >{this.state.statusMessage.message}</MessageBar>
                : ''}
            </div>
          </div >
        </div >


      </section >
    );
  }
  //transmittal id generation
  public async _trannsmittalIDGeneration() {
    let prefix;
    let separator;
    let sequenceNumber;
    let title;
    let counter;
    let transmittalID;
    let transmitTo;
    { this.state.transmitTo === "Customer" ? transmitTo = "Outbound Customer" : transmitTo = "Outbound Sub-contractor"; }
    await this._Service.getItemForSelectInLists(this.props.siteUrl, this.props.transmittalIdSettingsListName, "*", "TransmittalCategory eq '" + transmitTo + "' and(TransmittalType eq '" + this.state.transmittalType + "')")
      .then(transmittalIdSettingsItems => {
        prefix = transmittalIdSettingsItems[0].Prefix;
        separator = transmittalIdSettingsItems[0].Separator;
        sequenceNumber = transmittalIdSettingsItems[0].SequenceNumber;
        title = transmittalIdSettingsItems[0].Title;
        counter = transmittalIdSettingsItems[0].Counter;
        let increment = counter + 1;
        let incrementValue = increment.toString();
        this._transmittalSequenceNumber(incrementValue, sequenceNumber);
        transmittalID = prefix + separator + title + separator + this.state.projectNumber + separator + this.state.incrementSequenceNumber;
        console.log("transmittalID", transmittalID);
        this.setState({
          transmittalNo: transmittalID,
        });
        //counter updation
        let counterData = {
          Counter: increment,
        }
        this._Service.updateSiteItem(this.props.siteUrl, this.props.transmittalIdSettingsListName, transmittalIdSettingsItems[0].ID, counterData);
      });
  }
  //for preview section
  public _onPreviewBtnClick() {
    let totalFiles;
    totalFiles = add(this.state.itemsForGrid.length, this.state.itemsForExternalGrid.length);
    // alert(totalFiles);
    this.setState({
      totalNoOfFiles: totalFiles,
      previewDiv: false,
      showReviewModal: true,
    });


  }
  private async _confirmAndSendBtnClick() {
    // let sourceDocumentId;
    // let hidden = 1;
    //let statusCount = 0;
    let forTransmittalStatus;
    //total files
    let totalFiles;
    let convertKBtoMB;
    totalFiles = add(this.state.itemsForGrid.length, this.state.itemsForExternalGrid.length);
    //sizecalculating
    let totalsizeProjects = 0;
    let totalAdditional = 0;
    //size recalcalculating
    if (this.state.itemsForGrid.length > 0 || this.state.itemsForExternalGrid.length > 0) {
      for (let i = 0; i < this.state.itemsForGrid.length; i++) {
        totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
      }
      for (let k = 0; k < this.state.itemsForExternalGrid.length; k++) {
        totalAdditional = Number(totalAdditional) + Number(this.state.itemsForExternalGrid[k].fileSizeInMB);
      }
      let totalSize = add(totalAdditional, totalsizeProjects);
      convertKBtoMB = Number(totalSize).toFixed(2);
      this.setState({
        fileSize: Number(convertKBtoMB)
      });
      console.log(this.state.fileSize);
    }
    //if size greater than 10 mb 
    if (Number(convertKBtoMB) > 10 && (this.state.sendAsSharedFolder == false)) {
      this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "File size is greater than 10 MB.Please select the checkbox Send and Receive as shared folder", messageType: 1 }, });
    }
    else {
      if (this.transmittalID == null || this.transmittalID == "") {
        let selectedContactsTo = this.state.selectedContactsTo.toString();
        let selectedContactsCC = this.state.selectedContactsCC.toString();
        console.log(selectedContactsTo);
        if (this.state.transmitTo != "" && this.state.itemsForGrid.length != 0 && this.state.transmittalTypekey != "" && this.state.selectedContactsTo != null && this.validator.fieldValid("selectedContactsTo")) {
          this.setState({
            spinnerDiv: "",
            hideButtonAfterSubmit: "none",
            hideUnlockButton: "none",
          });
          await this._trannsmittalIDGeneration();
          //header list
          try {
            let headetData = {
              Title: this.state.transmittalNo,
              TransmittalCategory: this.state.transmitTo,
              Customer: this.state.customerName,
              CustomerID: (this.state.transmitTo == "Customer") ? this.state.customerId : "",
              SubContractor: this.state.subContractor,
              SubContractorID: (this.state.subContractorKey).toString(),
              ToEmails: selectedContactsTo,
              CCEmails: selectedContactsCC,
              Notes: this.state.notes,
              TransmittalStatus: (this.state.transmitTo != "Customer") ? "Completed" : "Ongoing",
              TransmittalType: this.state.transmittalType,
              TransmittedById: this.state.currentUser,
              SendAsSharedFolder: this.state.sendAsSharedFolder,
              ReceiveInSharedFolder: this.state.recieveInSharedFolder,
              SendAsMultipleEmails: this.state.sendAsMultipleFolder,
              TransmittalSize: (convertKBtoMB).toString(),
              TransmittalDate: new Date(),
              TotalFiles: (totalFiles).toString(),
              ToName: this.state.selectedContactsToDisplayName,
              CCName: this.state.selectedContactsCCDisplayName,
              CoverLetter: this.state.coverLetterNeeded,
              InternalCCId: this.state.internalCCContacts,
            }
            this._Service.createNewSiteProcess(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, headetData)
              .then(async outboundTransmittalHeader => {
                this.setState({ outboundTransmittalHeaderId: outboundTransmittalHeader.data.ID });
                for (let i = 0; i < this.state.itemsForGrid.length; i++) {
                  if (this.state.itemsForGrid[i].approvalRequired == true) {
                    forTransmittalStatus = "true";
                  }
                }
                let linkUpdationInheader = {
                  TransmittalLink: {
                    Description: "Project workspace",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.outBoundTransmittalSitePage + ".aspx?trid=" + outboundTransmittalHeader.data.ID + ""
                  },
                  TransmittalDetails: {
                    Description: "Transmittal Details",
                    Url: this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" + outboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=6da3a1b3%2D0155%2D48d9%2Da7c7%2Dd2e862c07db5"
                  },
                  OutboundAdditionalDetails: {
                    Description: "Outbound Additional Details",
                    Url: this.props.siteUrl + "/" + this.props.outboundAdditionalDocumentsListName + "/Forms/AllItems.aspx?FilterField1=TransmittalID&FilterValue1=" + outboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=bcc64a99-0907-4416-b9f6-8001acf1e000"
                  },
                  TransmittalStatus: (this.state.transmitTo === "Customer") ? forTransmittalStatus != "true" ? "Completed" : "Ongoing" : "Completed",
                }
                this._Service.updateSiteItem(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, outboundTransmittalHeader.data.ID, linkUpdationInheader);
                //outbound Details
                if (this.state.itemsForGrid.length > 0) {
                  this.state.itemsForGrid.forEach((i, index) => {
                    let obDetailData = {
                      Title: i.documentName,
                      TransmittalHeaderId: outboundTransmittalHeader.data.ID,
                      DocumentIndexId: i.documentIndexId,
                      Revision: i.revision,
                      TransmittalRevision: i.revision,
                      DueDate: i.dueDate,
                      Size: i.fileSizeInMB,
                      SentComments: i.comments,
                      CustomerAcceptanceCodeId: i.acceptanceCode,
                      TransmitFor: i.TransmitFor,
                      ApprovalRequired: i.approvalRequired,
                      TransmittalStatus: (i.approvalRequired == true && this.state.transmitTo == "Customer") ? "Ongoing" : "Completed",
                      DocumentLibraryID: i.publishDoumentlibraryID,
                      Slno: (Number(index) + Number(1)).toString(),
                      CustomerDocumentNo: i.customerDocumentNo,
                      SubcontractorDocumentNo: i.subcontractorDocumentNo,
                    }
                    this._Service.createNewSiteProcess(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, obDetailData);

                  })
                }
                //outbound additional
                if (this.state.itemsForExternalGrid.length > 0) {
                  this.state.itemsForExternalGrid.forEach(async (file, key) => {
                    const splitted = file.documentName.split(".");
                    const documentNameExtension = splitted.slice(0, -1).join('.') + "_" + this.state.transmittalNo + '.' + splitted[splitted.length - 1];
                    console.log(documentNameExtension);
                    let sourceDocumentMetadata = {
                      Title: file.documentName,
                      TransmittalIDId: this.state.outboundTransmittalHeaderId,
                      Size: file.fileSizeInMB,
                      Comments: file.externalComments,
                      SentDate: new Date(),
                      TransmittalStatus: "Ongoing",
                      Slno: (Number(key) + Number(1)).toString(),
                    }
                    await this._Service.uploadDocument(documentNameExtension, file.content, this.props.outboundAdditionalDocumentsListName, sourceDocumentMetadata)
                  });
                }
                //need to add additional documents code 
                //add document index updations
                this.triggerOutboundTransmittal(Number(this.state.outboundTransmittalHeaderId));
                this.setState({
                  hideButtonAfterSubmit: "none",
                  hideUnlockButton: "none",
                  spinnerDiv: "",
                });
              });
          }
          catch (error) {
            console.error("Error creating outbound transmittal:", error);
          }
          this.validator.hideMessages();
        }
        else {
          this.validator.showMessages();
          this.forceUpdate();
        }
      }


    }
  }
  public uploadFile = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = event.target.files;
    if (files && files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const duplicate = this.state.itemsForExternalGrid.some(tempItem => tempItem.documentName === file.name);
        const filename = file.name; // Replace this with your filename
        const fileExtension = filename.split('.').pop();
        const isZipFile = fileExtension === 'zip';
        const isAudioFile = ['mp3', 'wav', 'ogg', 'aac'].includes(fileExtension);
        const isVideoFile = ['mp4', 'avi', 'mkv', 'mov'].includes(fileExtension);
        if (isZipFile || isAudioFile || isVideoFile) {
          // Ignore ZIP, audio, and video files and image files
          // Optionally, you can display an error message or handle these files differently           
        } else {
          if (!duplicate) {
            let tempExternalFile = {
              documentName: file.name,
              fileSize: (((file.size / 1024)).toFixed(2)),
              fileSizeInMB: (((file.size / 1024) * 0.0009765625).toFixed(2)),
              externalComments: this.state.externalComments,
              content: file,
            }
            this.setState(prevState => ({
              itemsForExternalGrid: [...prevState.itemsForExternalGrid, tempExternalFile],
              showExternalGrid: false,
              fileSizeDiv: false,
            }));
          }
        }

      }

    }
  }
  private _getDropFiles = async (files: any) => {
    if (files.length > 0) {
      if (files !== "") {
        files.forEach((item, key) => {
          const duplicate = this.state.itemsForExternalGrid.some(tempItem => tempItem.documentName === item.name);
          const filename = item.name;
          const fileExtension = filename.split('.').pop();
          const isZipFile = fileExtension === 'zip';
          const isAudioFile = ['mp3', 'wav', 'ogg', 'aac'].includes(fileExtension);
          const isVideoFile = ['mp4', 'avi', 'mkv', 'mov'].includes(fileExtension);
          // const allowedExtensions = ['jpg', 'jpeg', 'png', 'gif'].includes(fileExtension);
          if (isZipFile || isAudioFile || isVideoFile) {
            // Ignore ZIP, audio, and video files and image files
            // Optionally, you can display an error message or handle these files differently           
          } else {
            if (!duplicate) {
              let tempExternalFile = {
                documentName: item.name,
                fileSize: (((item.size / 1024)).toFixed(2)),
                fileSizeInMB: (((item.size / 1024) * 0.0009765625).toFixed(2)),
                externalComments: this.state.externalComments,
                content: item,
              }

              this.setState(prevState => ({
                itemsForExternalGrid: [...prevState.itemsForExternalGrid, tempExternalFile],
                showExternalGrid: false,
                fileSizeDiv: false,
              }));
            }
          }
        });
      }
    }
  }
  private onCommentChange = (index, event) => {
    const newMultiline = this.stripHtmlTags(event).length > 50;
    if (newMultiline !== this.state.toggleMultiline) {
      this.setState({
        toggleMultiline: true,
      });
    }
    const { itemsForGrid } = this.state;
    const updatedItems = [...itemsForGrid];
    updatedItems[index].comments = this.stripHtmlTags(event);
    this.setState({ itemsForGrid: updatedItems });
  }
  private onExternalCommentChange = (index, event) => {
    const newMultiline = this.stripHtmlTags(event).length > 50;
    if (newMultiline !== this.state.toggleMultiline) {
      this.setState({
        toggleMultiline: true,
      });
    }
    const { itemsForExternalGrid } = this.state;
    const updatedItems = [...itemsForExternalGrid];
    updatedItems[index].externalComments = this.stripHtmlTags(event);
    this.setState({ itemsForExternalGrid: updatedItems });
  }
  //stripHtmlTags
  public stripHtmlTags = (html) => {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    return doc.body.textContent || '';
  };

  public UNSAFE_componentWillMount = () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields"
      }
    });
  }
  public async componentDidMount() {
    await this.projectInformation();
    this._userMessageSettings();
    this._currentUser();
    this._transmitForBind();
    this._queryParamGetting();
    // //this._LAUrlGetting();  
    await this.loadSettingsList();
    this._loadPublishDocuments("");
  }
  private loadSettingsList = async () => {
    await this._Service.getListItems(this.props.context.pageContext.site.serverRelativeUrl, "SettingsList")
      .then(settings => {
        const tempArray = settings.filter(item => item.Active === true);
        console.log(tempArray);
        this.setState({ settingsListArray: tempArray })
        const resultArrays = [];
        // Use Promise.all to handle multiple Promises
        Promise.all(tempArray.map(element =>
          this._Service.getListItems(this.props.context.pageContext.site.serverRelativeUrl, element.Title)
        ))
          .then(results => {
            // results is an array containing the resolved values of each Promise
            tempArray.forEach((element, index) => {
              console.log(results[index])
              let listItems = results[index].map(item => ({
                key: item.Title,
                text: item.Title
              }));
              resultArrays[element.Title] = listItems;

            });
            this.setState({ settingsListsItemsArray: resultArrays })
          })
          .catch(error => {
            // Handle errors here
            console.error(error);
          });
      });
  }
  //Row Comment item changes
  public handleSettingsListItemsChange = (index, event, item) => {
    const { settingsListArray, documentFilters } = this.state;
    const updatedItems = [...settingsListArray];
    // Update the selected key for the specific item
    updatedItems[index].selectedKey = event.key;
    documentFilters[item] = event.key;
    // Update the state
    this.setState({ settingsListArray: updatedItems });
    this._loadPublishDocuments(item);
  };
  // document selection
  private setSelectedDocuments = async (option) => {
    await this.setState({
      projectDocumentSelectKey: option.value,
      documentSelectedDiv: true,
      hideGridAddButton: false,
      selectedDocuments: option,
    });
    const tempFile = [];
    if (option.length !== 0) {
      option.forEach(selectedDocuments => {
        const duplicate = this.state.itemsForGrid.some(tempItem => tempItem.documentIndexId === selectedDocuments.DocumentIndexId);
        if (!duplicate) {
          let temp = {
            publishDoumentlibraryID: selectedDocuments.value,
            documentIndexId: selectedDocuments.DocumentIndexId,
            DueDate: moment(this.state.dueDate).format("DD/MM/YYYY"),
            dueDate: this.state.dueDate,
            comments: this.state.comments,
            revision: selectedDocuments.Revision,
            documentID: selectedDocuments.DocumentID,
            documentName: selectedDocuments.DocumentName,
            fileSize: (((selectedDocuments.FileSizeDisplay / 1024)).toFixed(2)),
            fileSizeInMB: (Number((selectedDocuments.FileSizeDisplay / 1024) * 0.0009765625).toFixed(2)),
            transmitFor: this.state.transmitFor,
            approvalRequired: this.state.approvalRequired,
            transmitForKey: this.state.transmitForKey,
            temporary: "",
            customerDocumentNo: selectedDocuments.CustomerDocumentNo,
          };
          tempFile.push(temp)
          this.setState(prevState => ({
            itemsForGrid: [...prevState.itemsForGrid, temp],
          }));
        }
      });
      this.setState({
        showGrid: false,
        projectDocumentSelectKey: "",
        fileSizeDiv: false,
        searchText: "",
      });
    }


  }

  protected async triggerOutboundTransmittal(transmittalID) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const laUrl = await this._Service.getHubItemsWithFilter(this.props.masterListName, "Title eq 'EMEC_OutboundTransmittal'", this.props.hubSiteUrl);
    console.log("Posturl", laUrl[0].PostUrl);
    const postURL = laUrl[0].PostUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'TransmittalNo': transmittalID,
      'ProjectName': this.state.projectName,
      'ContractNumber': this.state.contractNumber,
      'ProjectNumber': this.state.projectNumber,
      'CoverLetterNeeded': (this.state.coverLetterNeeded == true ? "Yes" : "NO"),
      'InternalContactsEmails': this.state.internalContactsEmail,
      'InternalContactsDisplayNames': this.state.internalCCContactsDisplayNameForPreview,
      'OutboundTransmittalDetails': this.state.itemsForGrid
    });
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: body
    };
    // let responseText: string = "";
    let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
    let responseJSON = await response.json();
    //responseText = JSON.stringify(responseJSON);
    console.log(responseJSON);
    if (response.ok) {
      // alert(response.text);
      if (responseJSON['Status'] == "MailSend") {
        this.setState({
          hideButtonAfterSubmit: "none",
          hideUnlockButton: "none",
          normalMsgBar: "",
          spinnerDiv: "none",
          statusMessage: { isShowMessage: true, message: "Transmittal Send Successfully", messageType: 4 },
        });
        setTimeout(() => {
          window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
        }, 10000);
      }
      else {

      }
    }
    else { }

  }

  private async _userMessageSettings() {
    // const userMessageSettings: any[] = await this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.userMessageSettings)
    //   .items.select("Title,Message").filter("PageName eq 'OutboundTransmittal'")();
    // for (var i in userMessageSettings) {
    //   if (userMessageSettings[i].Title === "OutboundTransmittalRecall") {
    //     this.setState({ outboundRecallConfirmation: userMessageSettings[i].Message });
    //   }
    // }
  }
  //for query param gettings
  private _queryParamGetting() {
    let params = new URLSearchParams(window.location.search);
    this.transmittalID = params.get('trid');
    if (this.transmittalID != "" && this.transmittalID != null) {
      this._Service.getItembyID(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, Number(this.transmittalID))
        .then(transmittalHeaderItems => {
          if (transmittalHeaderItems.TransmittalStatus == "Ongoing") {
            this.setState({
              hideButtonAfterSubmit: "none",
              hideUnlockButton: "none",
            });
            // this._LAUrlGettingForPermission();
          }
          else if (transmittalHeaderItems.TransmittalStatus == "Completed") {
            this.setState({
              hideButtonAfterSubmit: "none",
              hideUnlockButton: "none",
            });
          }
          this.setState({
            divForToAndCCSearch: "none",
            divForToAndCC: "",
          });
        });
      //  this.bindOutboundTransmittalSavedData(this.transmittalID);
      this.setState({
        transmittalNo: "",
        webpartView: "",
      });
    }
    else {
      this.setState({
        transmittalNo: "none",
        transmitTypeForDefault: "",
        loaderDisplay: "none",
        webpartView: "",
        divForToAndCCSearch: "",
        divForToAndCC: "none",
      });
    }
  }
  //settings for date format
  private _onFormatDate = (date: Date): string => {
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  }
  //rebinding the fields after save as draft
  // private async bindOutboundTransmittalSavedData(transmittalID: string) {
  //   //sizecalculating
  //   let totalsizeProjects = 0;
  //   let totalAdditional = 0;
  //   let customerArray: { key: any; text: string; }[] = [];
  //   let subContractor: { key: any; text: string; }[] = [];
  //   let toMails;
  //   let ccMails;
  //   let toMailsDisplay;
  //   let ccMailsDisplay;
  //   let multipledoc;
  //   let finalTo: any[] = [];
  //   let finalCC: any[] = [];
  //   let finalToDisplay: any[] = [];
  //   let finalCCDisplay: any[] = [];
  //   let contactsToRebind: any[] = [];
  //   let tempInternalCCID: any[] = [];
  //   let tempInternalCCName: any[] = [];
  //   let tempInternalCCNameForMail: any[] = [];

  //   //binding from outbound transmittal header
  //   await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalHeaderListName).items.filter("ID eq '" + this.transmittalID + "' ").get().then(async (outboundTransmittalHeader: { InternalCCId: null; }[]) => {
  //     console.log("outboundTransmittalHeader", outboundTransmittalHeader);
  //     console.log(outboundTransmittalHeader[0].ToEmails);
  //     toMails = outboundTransmittalHeader[0].ToEmails === null ? "" : outboundTransmittalHeader[0].ToEmails.split(';');
  //     ccMails = outboundTransmittalHeader[0].CCEmails === null ? "" : outboundTransmittalHeader[0].CCEmails.split(';');
  //     toMailsDisplay = outboundTransmittalHeader[0].ToName === null ? "" : outboundTransmittalHeader[0].ToName.split(',');
  //     ccMailsDisplay = outboundTransmittalHeader[0].CCName === null ? "" : outboundTransmittalHeader[0].CCName.split(',');
  //     console.log(toMails);
  //     let toMailKey = "";
  //     for (let k = 0; k < toMails.length; k++) {
  //       finalTo.push(toMails[k]);
  //     }
  //     for (let k = 0; k < ccMails.length; k++) {
  //       finalCC.push(ccMails[k]);
  //     }
  //     for (let k = 0; k < toMailsDisplay.length; k++) {
  //       finalToDisplay.push(toMailsDisplay[k]);
  //     }
  //     for (let k = 0; k < ccMailsDisplay.length; k++) {
  //       finalCCDisplay.push(ccMailsDisplay[k]);
  //     }
  //     // contactsToRebind = replaceString(outboundTransmittalHeader[0].ToEmails, ';', ',');
  //     console.log(contactsToRebind);
  //     this.emailsSelectedTo = finalTo;
  //     this.emailsSelectedCC = finalCC;
  //     this.contactToDisplay = finalToDisplay;
  //     this.contactCCDisplay = finalCCDisplay;
  //     this.setState({
  //       transmittalNo: outboundTransmittalHeader[0].Title,
  //       selectedContactsToDisplayName: outboundTransmittalHeader[0].ToName,
  //       selectedContactsTo1: finalTo,
  //       selectedContactsTo: (outboundTransmittalHeader[0].ToEmails === null) ? " " : outboundTransmittalHeader[0].ToEmails,
  //       // selectedContactsTo: contactsToRebind,
  //       selectedContactsToCCRebind: finalCC,
  //       selectedContactsCC: (outboundTransmittalHeader[0].CCEmails === null) ? " " : outboundTransmittalHeader[0].CCEmails,
  //       selectedContactsCCDisplayName: outboundTransmittalHeader[0].CCName,
  //       notes: outboundTransmittalHeader[0].Notes,
  //       transmittalTypekey: outboundTransmittalHeader[0].TransmittalType,
  //       transmittalType: outboundTransmittalHeader[0].TransmittalType,
  //       fileSize: outboundTransmittalHeader[0].TransmittalSize,
  //     });
  //     if (outboundTransmittalHeader[0].TransmittalType === "Document" && outboundTransmittalHeader[0].TransmittalCategory === "Customer") {
  //       this._loadPublishDocuments();
  //     }
  //     else if (outboundTransmittalHeader[0].TransmittalType === "Letter" && outboundTransmittalHeader[0].TransmittalCategory === "Customer") {
  //       this._loadSourceDocumentsForLetter();
  //     }
  //     else if (outboundTransmittalHeader[0].TransmittalType === "Document" && outboundTransmittalHeader[0].TransmittalCategory === "Sub-Contractor") {
  //       this._loadSourceDocuments();
  //     }
  //     else if (outboundTransmittalHeader[0].TransmittalType === "Letter" && outboundTransmittalHeader[0].TransmittalCategory === "Sub-Contractor") {
  //       this._loadSourceDocumentsForLetter();
  //     }
  //     if (outboundTransmittalHeader[0].TransmittalType === "Document") {
  //       this.setState({
  //         transmittalTypekey: "Document",
  //         transmitTypeForDocument: "",
  //         transmitTypeForLetter: "none",
  //         transmitTypeForDefault: "none",

  //       });
  //     }
  //     else if (outboundTransmittalHeader[0].TransmittalType === "Letter") {
  //       this.setState({
  //         transmittalTypekey: "Letter",
  //         transmitTypeForDocument: "none",
  //         transmitTypeForLetter: "",
  //         transmitTypeForDefault: "none",
  //       });
  //     }
  //     if (outboundTransmittalHeader[0].TransmittalSize === null) {
  //       this.setState({
  //         fileSizeDiv: true,
  //       });
  //     }
  //     else {
  //       this.setState({
  //         fileSizeDiv: true,
  //       });
  //     }
  //     //Transmit To CUSTOMER
  //     if (outboundTransmittalHeader[0].TransmittalCategory === "Customer") {
  //       // this._loadPublishDocuments();
  //       this.setState({
  //         transmitToKey: "1",
  //         transmitTo: "Customer",
  //         hideCustomer: "",
  //         customerName: outboundTransmittalHeader[0].Customer,
  //         dropDownReadonly: true,

  //       });
  //       this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.contactListName).items
  //         // .filter("CompanyId eq '" + this.state.customerId + "' ").get()
  //         .filter("CustomerOrVendorID eq '" + this.state.customerId + "'  and  LegalEntityId eq '" + this.state.legalId + "'").get()
  //         .then((contacts: { [x: string]: { Email: string; }; }) => {
  //           console.log("contacts", contacts);
  //           for (var k in contacts) {
  //             if (contacts[k].Active === true) {
  //               console.log("contacts", contacts);
  //               this.setState({
  //                 contacts: contacts,
  //               });
  //               let transmitForItemdata = {
  //                 key: contacts[k].Email,
  //                 text: contacts[k].Title + " " + (contacts[k].LastName != null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
  //               };
  //               customerArray.push(transmitForItemdata);
  //             }
  //           }
  //           this.setState({
  //             contacts: customerArray,
  //           });
  //         });
  //       //binding from outbound tranmittal details
  //       let selectHeaderItems = "Id,DocumentIndex/ID,DocumentIndex/Title,DueDate,SentComments,Revision,Title,Size,TransmittedFor/ID,TransmittedFor/Title,Temporary,TransmittalHeader/ID,DocumentLibraryID,ID,ApprovalRequired,CustomerDocumentNo,SubcontractorDocumentNo";
  //       sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName).items.select(selectHeaderItems).expand("DocumentIndex,TransmittedFor,TransmittalHeader").filter("TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ").get().then((outboundTransmittalDetailsListName: string | any[]) => {
  //         console.log("outboundTransmittalDetailsListName", outboundTransmittalDetailsListName);
  //         if (outboundTransmittalDetailsListName.length >= 2) {
  //           this.setState({
  //             sendAsMultipleEmailCheckBoxDiv: "",
  //           });
  //         }
  //         if (outboundTransmittalDetailsListName.length > 0) {
  //           for (var k = 0; k <= outboundTransmittalDetailsListName.length; k++) {
  //             this.state.tempArrayForPublishedDocumentGrid.push({
  //               publishDoumentlibraryID: outboundTransmittalDetailsListName[k]['DocumentLibraryID'],
  //               //transmittalHeaderID:outboundTransmittalDetailsListName[k].TransmittalHeader.ID,
  //               outboundDetailsID: outboundTransmittalDetailsListName[k].ID,
  //               documentIndexId: outboundTransmittalDetailsListName[k].DocumentIndex['ID'],
  //               DueDate: moment(outboundTransmittalDetailsListName[k].DueDate).format("DD/MM/YYYY"),
  //               dueDate: outboundTransmittalDetailsListName[k].DueDate,
  //               comments: outboundTransmittalDetailsListName[k].SentComments,
  //               revision: outboundTransmittalDetailsListName[k].Revision,
  //               // documentID: outboundTransmittalDetailsListName[k].DocumentID,
  //               documentName: outboundTransmittalDetailsListName[k].Title,
  //               //acceptanceCode:outboundTransmittalDetailsListName[k].CustomerAcceptanceCode.ID,
  //               fileSize: outboundTransmittalDetailsListName[k].Size,
  //               fileSizeInMB: outboundTransmittalDetailsListName[k].Size,
  //               transmitFor: outboundTransmittalDetailsListName[k].TransmittedFor.Title,
  //               transmitForKey: outboundTransmittalDetailsListName[k].TransmittedFor.ID,
  //               temporary: outboundTransmittalDetailsListName[k].Temporary,
  //               customerDocumentNo: outboundTransmittalDetailsListName[k].CustomerDocumentNo,
  //             });
  //             this.setState({
  //               itemsForGrid: this.state.tempArrayForPublishedDocumentGrid,
  //               showGrid: false,
  //               currentOutboundDetailItem: outboundTransmittalDetailsListName,
  //             });
  //             if (outboundTransmittalDetailsListName[k].Size >= 10 && outboundTransmittalDetailsListName.length >= 2) {
  //               this.setState({
  //                 sendAsMultipleEmailCheckBoxDiv: "none",
  //               });
  //               multipledoc = "Yes";
  //             }

  //           }
  //           for (let i = 0; i < this.state.itemsForExternalGrid.length; i++) {
  //             if (this.state.itemsForExternalGrid[i].fileSizeInMB >= 10 && this.state.itemsForExternalGrid.length >= 2) {
  //               this.setState({
  //                 sendAsMultipleEmailCheckBoxDiv: "none",
  //               });
  //             }
  //           }
  //         }

  //       });
  //     }
  //     //Transmit To SUBCONTRACTOR
  //     else if (outboundTransmittalHeader[0].TransmittalCategory === "Sub-Contractor") {
  //       this.setState({
  //         transmitToKey: "2",
  //         transmitTo: "Sub-Contractor",
  //         hideSubContractor: "",
  //         hideCustomer: "none",
  //         subContractorLabel: "",
  //         subContractorDrpDwn: "none",
  //       });
  //       if (outboundTransmittalHeader[0].SubContractorID != null) {
  //         this.setState({
  //           subContractorKey: outboundTransmittalHeader[0].SubContractorID,
  //           subContractor: outboundTransmittalHeader[0].SubContractor,
  //           dropDownReadonly: true,
  //         });
  //         const outBoundTransmittalDetails = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName).items.filter("TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ").get();
  //         console.log(outBoundTransmittalDetails);
  //         //binding from outbound tranmittal details
  //         let selectHeaderItems = "Id,DocumentIndex/ID,DocumentIndex/Title,DueDate,SentComments,Revision,Title,Size,TransmittedFor/ID,TransmittedFor/Title,Temporary,TransmittalHeader/ID,DocumentLibraryID,ID,CustomerAcceptanceCode/ID,CustomerAcceptanceCode/Title,CustomerDocumentNo,SubcontractorDocumentNo";
  //         sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName).items.select(selectHeaderItems).expand("DocumentIndex,TransmittedFor,TransmittalHeader,CustomerAcceptanceCode").filter("TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ").get().then((outboundTransmittalDetailsListName: string | any[]) => {
  //           console.log("outboundTransmittalDetailsListName", outboundTransmittalDetailsListName);
  //           if (outboundTransmittalDetailsListName.length >= 2) {
  //             this.setState({
  //               sendAsMultipleEmailCheckBoxDiv: "",
  //             });
  //           }
  //           if (outboundTransmittalDetailsListName.length > 0) {
  //             for (var k = 0; k < outboundTransmittalDetailsListName.length; k++) {
  //               this.state.tempArrayForPublishedDocumentGrid.push({
  //                 outboundDetailsID: outboundTransmittalDetailsListName[k].ID,
  //                 documentIndexId: outboundTransmittalDetailsListName[k].DocumentIndex['ID'],
  //                 DueDate: moment(outboundTransmittalDetailsListName[k].DueDate).format("DD/MM/YYYY "),
  //                 dueDate: outboundTransmittalDetailsListName[k].DueDate,
  //                 comments: outboundTransmittalDetailsListName[k].SentComments,
  //                 revision: outboundTransmittalDetailsListName[k].Revision,
  //                 // documentID: outboundTransmittalDetailsListName[k].DocumentID,
  //                 documentName: outboundTransmittalDetailsListName[k].Title,
  //                 acceptanceCodeTitle: (outBoundTransmittalDetails[k].CustomerAcceptanceCodeId != null) ? outboundTransmittalDetailsListName[k].CustomerAcceptanceCode.Title : "",
  //                 fileSize: outboundTransmittalDetailsListName[k].Size,
  //                 fileSizeInMB: outboundTransmittalDetailsListName[k].Size,
  //                 transmitFor: outboundTransmittalDetailsListName[k].TransmittedFor.Title,
  //                 transmitForKey: outboundTransmittalDetailsListName[k].TransmittedFor.ID,
  //                 temporary: outboundTransmittalDetailsListName[k].Temporary,
  //                 subcontractorDocumentNo: outboundTransmittalDetailsListName[k].SubcontractorDocumentNo,
  //               });
  //               this.setState({
  //                 itemsForGrid: this.state.tempArrayForPublishedDocumentGrid,
  //                 showGrid: false,
  //                 currentOutboundDetailItem: outboundTransmittalDetailsListName,
  //               });
  //               //multiple mail checkbox 
  //               if (outboundTransmittalDetailsListName[k].Size >= 10 && outboundTransmittalDetailsListName.length >= 2) {
  //                 this.setState({
  //                   sendAsMultipleEmailCheckBoxDiv: "none",
  //                 });
  //                 multipledoc = "Yes";
  //               }
  //               //multiple mail checkbox wrt additional
  //               for (let i = 0; i < this.state.itemsForExternalGrid.length; i++) {
  //                 if (this.state.itemsForExternalGrid[i].fileSizeInMB >= 10 && this.state.itemsForExternalGrid.length >= 2) {
  //                   this.setState({
  //                     sendAsMultipleEmailCheckBoxDiv: "none",
  //                   });
  //                 }
  //               }
  //             }
  //           }

  //         });
  //       }
  //       //Transmit To SUBCONTRACTOR
  //       if (this.state.transmitTo === "Sub-Contractor") {
  //         this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.contactListName).items
  //           .filter("CustomerOrVendorID eq '" + outboundTransmittalHeader[0].SubContractorID + "' and  LegalEntityId eq '" + this.state.legalId + "' ").get()
  //           .then((contacts: { [x: string]: { Email: string; }; }) => {
  //             for (var k in contacts) {
  //               if (contacts[k].Active === true) {
  //                 let transmitForItemdata = {
  //                   key: contacts[k].Email,
  //                   text: contacts[k].Title + " " + (contacts[k].LastName != null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
  //                 };
  //                 subContractor.push(transmitForItemdata);
  //               }
  //             }
  //             this.setState({
  //               contacts: subContractor
  //             });
  //           });
  //       }
  //     }
  //     //for size div visible
  //     if (outboundTransmittalHeader[0].TransmittalStatus === "Draft") {
  //       this.setState({
  //         fileSizeDiv: false,
  //       });
  //     }
  //     if (outboundTransmittalHeader[0].ReceiveInSharedFolder === true) {
  //       this.setState({
  //         recieveInSharedFolder: true,
  //       });
  //     }
  //     if (outboundTransmittalHeader[0].SendAsMultipleEmails === true) {
  //       this.setState({
  //         sendAsMultipleFolder: true,
  //       });
  //     }
  //     if (outboundTransmittalHeader[0].SendAsSharedFolder === true) {
  //       this.setState({
  //         sendAsSharedFolder: true,
  //       });
  //     }
  //     if (outboundTransmittalHeader[0].CoverLetter === true) {
  //       this.setState({
  //         coverLetterNeeded: true,
  //       });
  //     }
  //     if (outboundTransmittalHeader[0].InternalCCId != null) {
  //       sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalHeaderListName).items.getById(this.transmittalID).select("InternalCC/ID, InternalCC/Title,InternalCC/EMail").expand("InternalCC").get().then((internalContacts: { InternalCC: { [x: string]: { EMail: any; }; }; }) => {
  //         for (var k in internalContacts.InternalCC) {
  //           tempInternalCCID.push(internalContacts.InternalCC[k].ID);
  //           this.setState({
  //             internalCCContacts: tempInternalCCID,
  //           });
  //           tempInternalCCName.push(internalContacts.InternalCC[k].Title,);
  //           tempInternalCCNameForMail.push(internalContacts.InternalCC[k].EMail,);
  //         }

  //         var InternalEmailID = tempInternalCCNameForMail.toString();
  //         let InternalEmailIDSemicolonAttached = replaceString(InternalEmailID, ',', ';');
  //         console.log("internalCCEmail", InternalEmailIDSemicolonAttached);
  //         this.setState({
  //           internalCCContactsDisplayName: tempInternalCCName,
  //           internalCCContactsDisplayNameForPreview: tempInternalCCName.toString(),
  //           internalContactsEmail: InternalEmailIDSemicolonAttached,
  //         });
  //         console.log("tempInternalCCName", tempInternalCCName);
  //       });
  //     }

  //   });
  //   //  //binding from outbound transmittal additional documents 
  //   sp.web.getList(this.props.siteUrl + "/" + this.props.outboundAdditionalDocumentsListName).items.filter("TransmittalIDId eq '" + this.transmittalID + "' ").get().then((outboundAdditionalDocumentsListName: string | any[]) => {
  //     console.log("outboundAdditionalDocumentsListName", outboundAdditionalDocumentsListName);
  //     if (outboundAdditionalDocumentsListName.length > 0) {
  //       for (var k = 0; k <= outboundAdditionalDocumentsListName.length; k++) {
  //         this.state.tempArrayForExternalDocumentGrid.push({
  //           additionalDocumentID: outboundAdditionalDocumentsListName[k]['ID'],
  //           documentName: outboundAdditionalDocumentsListName[k].Title,
  //           fileSize: outboundAdditionalDocumentsListName[k].Size,
  //           fileSizeInMB: outboundAdditionalDocumentsListName[k].Size,
  //           externalComments: outboundAdditionalDocumentsListName[k].Comments,
  //           //   content:myfile,
  //         });
  //         this.setState({
  //           currentOutboundAdditionalItem: outboundAdditionalDocumentsListName,
  //           showExternalGrid: false,
  //           //fileSizeDiv: false,
  //           itemsForExternalGrid: this.state.tempArrayForExternalDocumentGrid,
  //         });
  //         if (outboundAdditionalDocumentsListName[k].Size >= 10 && this.state.itemsForGrid.length >= 2) {
  //           this.setState({
  //             sendAsMultipleEmailCheckBoxDiv: "none",
  //           });
  //           multipledoc = "Yes";
  //         }
  //       }
  //       if (this.state.itemsForGrid.length > 0 || outboundAdditionalDocumentsListName.length > 0) {
  //         for (let i = 0; i < this.state.itemsForGrid.length; i++) {
  //           totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
  //         }
  //         for (let k = 0; k < outboundAdditionalDocumentsListName.length; k++) {
  //           totalAdditional = Number(totalAdditional) + Number(outboundAdditionalDocumentsListName[k].Size);
  //         }

  //         let totalSize = add(totalAdditional, totalsizeProjects);
  //         let convertKBtoMB = Number(totalSize).toFixed(2);
  //         this.setState({
  //           fileSize: Number(convertKBtoMB),
  //           fileSizeDiv: false,
  //         });
  //         console.log(this.state.fileSize);
  //       }

  //     }
  //   });
  //   if (this.state.fileSize > "10") {
  //     this.setState({
  //       fileSizeDivForRebind: "",
  //     });
  //   }
  //   if (multipledoc === "Yes") {
  //     this.setState({
  //       sendAsMultipleEmailCheckBoxDiv: "none",
  //     });
  //   }
  //   else {
  //     this.setState({
  //       sendAsMultipleEmailCheckBoxDiv: " ",
  //     });
  //   }
  // }
  //tranmit to dropdown
  public _drpdwnTransmitTo(event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void {
    const customerArray: { key: any; text: string; }[] = [];
    const customerArraySearch: { value: any; label: string; }[] = [];
    this.setState({
      transmitToKey: (option.key).toString(),
      transmitTo: option.text,
      searchDocuments: [],
      contactsForSearch: [],
      showGrid: true,
      itemsForGrid: [],
      tempArrayForPublishedDocumentGrid: [],
      publishDocumentsItemsForGrid: [],
      selectedContactsToName: [],
      selectedContactsCCName: [],
      selectedContactsToDisplayName: "",
      selectedContactsCCDisplayName: "",
      transmittalType: "",
      transmittalTypekey: "",
      coverLetterNeeded: false
    });
    if (option.text === "Customer") {
      this._Service.getHubItemsWithFilter(this.props.contactListName, "CustomerOrVendorID eq '" + this.state.customerId + "'  and  LegalEntityId eq '" + this.state.legalId + "'", this.props.hubSiteUrl)
        .then(contacts => {
          for (let k in contacts) {
            if (contacts[k].Active === true) {
              const transmitForItemdata = {
                key: contacts[k].Email,
                text: contacts[k].Title + " " + (contacts[k].LastName !== null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
              };
              let transmitForItemdataSearch = {
                value: contacts[k].Email,
                label: contacts[k].Title + " " + (contacts[k].LastName !== null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
              };
              customerArray.push(transmitForItemdata);
              customerArraySearch.push(transmitForItemdataSearch);
            }
          }
        });
      return this.setState({
        hideCustomer: "",
        subContractorItems: [],
        subContractorKey: "",
        hideSubContractor: "none",
        contacts: customerArray,
        contactsForSearch: customerArraySearch
      });
    }
    else if (option.text === "Sub-Contractor") {
      this.setState({
        documentSelect: "",
        documentSelectedDiv: true,
      });
      const subcontractorArray: { value: any; label: any; }[] = [];
      this._Service.getHubItemsWithFilter("SubContractorMaster", "ProjectId eq '" + this.state.projectNumber + "' and  Title eq '" + this.state.legalId + "'", this.props.hubSiteUrl)
        .then(subcontractor => {
          for (let i = 0; i < subcontractor.length; i++) {
            let subcontractorItemdata = {
              value: subcontractor[i].VendorId,
              label: subcontractor[i].VendorName
            };
            subcontractorArray.push(subcontractorItemdata);
            this.setState({
              contacts: subcontractorArray,
            });
          }
          return this.setState({
            contacts: [],
            hideSubContractor: "",
            hideCustomer: "none",
            subContractorItems: subcontractorArray
          });
        });
    }
  }
  //bin
  public _drpdwnSubContractor(option) {
    const subContractor: { key: any; text: string; }[] = [];
    let subContractorArray = [];
    const subContractorArraySearch: { value: any; label: string; }[] = [];
    this.setState({
      subContractorKey: option.value,
      subContractor: option.label,
      searchDocuments: [],
      contactsForSearch: [],
      showGrid: true,
      itemsForGrid: [],
      tempArrayForPublishedDocumentGrid: [],
      publishDocumentsItemsForGrid: [],
      selectedContactsToName: [],
      selectedContactsCCName: [],
      selectedContactsToDisplayName: "",
      selectedContactsCCDisplayName: "",
      transmittalType: "",
      transmittalTypekey: "",
      coverLetterNeeded: false
    });
    this._Service.getHubItemsWithFilter(this.props.contactListName, "CustomerOrVendorID eq '" + option.value + "' and  LegalEntityId eq '" + this.state.legalId + "' ", this.props.hubSiteUrl)
      .then(contacts => {
        for (let k in contacts) {
          if (contacts[k].Active === true) {
            let transmitForItemdata = {
              key: contacts[k].Email,
              text: contacts[k].Title + " " + (contacts[k].LastName != null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
            };
            let transmitForItemdataSearch = {
              value: contacts[k].Email,
              label: contacts[k].Title + " " + (contacts[k].LastName != null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
            };
            subContractorArray.push(transmitForItemdata);
            subContractorArraySearch.push(transmitForItemdataSearch);
            subContractor.push(transmitForItemdata);
          }
        }
        this.setState({
          contacts: subContractor,
          contactsForSearch: subContractorArraySearch,
          subContractorKey: option.value,
          subContractor: option.label
        });
      });

  }
  public _drpdwnTransmitFor = async (index, event, item) => {
    //this.setState({ transmitForKey: (option.key).toString(), transmitFor: option.text });
    const { itemsForGrid } = this.state;
    const updatedItems = [...itemsForGrid];
    updatedItems[index].TransmittedFor = event.key;
    updatedItems[index].TransmitFor = event.text;
    this.setState({ itemsForGrid: updatedItems });

    // const select = "ApprovalRequired,AcceptanceCode";
    // const filter = "Title eq '" + option.text + "'";
    // this._Service.getItemForSelectInLists(this.props.siteUrl, this.props.transmittalCodeSettingsListName, select, filter)
    //   .then((transmitfor: { ApprovalRequired: any; }[]) => {
    //     this.setState({
    //       approvalRequired: transmitfor[0].ApprovalRequired,
    //     });
    //   });
  }
  private _hideGrid() {
    this.setState({
      confirmCancelDialog: false,
      cancelConfirmMsg: "",
    });
  }
  // private _dueDatePickerChange = (index: number,event: any)=>{
  //   const { itemsForGrid } = this.state;
  //   const updatedItems = [...itemsForGrid];
  //   updatedItems[index].dueDate = this.stripHtmlTags(event);
  //   this.setState({ itemsForGrid: updatedItems });
  // }

  private notesOnchange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState({ notes: newText || '' });
  }
  // private onCommentExternalChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
  //   const newMultiline = newText.length > 50;
  //   if (newMultiline !=== this.state.toggleMultiline) {
  //     this.setState({
  //       toggleMultiline: true,
  //     });
  //   }
  //   this.setState({ externalComments: newText || '' });
  //}

  private _closeModal = (): void => {
    this.setState({ showReviewModal: false });
  }
  //temporary array for external documents grid.
  private _showExternalGrid() {
    this.setState({
      fileSizeDivForRebind: "none",
    });
    if (this.state.externalComments !== "") {
      this.validator.hideMessages();
      //sizecalculating
      let totalsizeProjects = 0;
      let totalAdditional = 0;
      if ((document.querySelector("#newfile") as HTMLInputElement).files[0] !== null) {
        let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
        if (myfile.size) {
          const duplicate = this.state.itemsForExternalGrid.filter(extItem => extItem.documentName === myfile.name)
          if (!duplicate) {
            let tempExternalFile = {
              documentName: myfile.name,
              fileSize: (((myfile.size / 1024)).toFixed(2)),
              fileSizeInMB: (((myfile.size / 1024) * 0.0009765625).toFixed(2)),
              externalComments: this.state.externalComments,
              content: myfile,
            }
            this.setState(prevState => ({
              itemsForExternalGrid: [...prevState.itemsForExternalGrid, tempExternalFile],
              showExternalGrid: false,
              fileSizeDiv: false,
            }));
          }
        }
      }
      //for calculating document size
      if (this.state.itemsForGrid.length > 0 || this.state.tempArrayForExternalDocumentGrid.length > 0) {
        for (let i = 0; i < this.state.itemsForGrid.length; i++) {
          totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
        }
        for (let k = 0; k < this.state.tempArrayForExternalDocumentGrid.length; k++) {
          totalAdditional = Number(totalAdditional) + Number(this.state.tempArrayForExternalDocumentGrid[k].fileSizeInMB);
        }

        let totalSize = add(totalAdditional, totalsizeProjects);
        let convertKBtoMB = Number(totalSize).toFixed(2);
        this.setState({
          fileSize: Number(convertKBtoMB)
        });
        console.log(this.state.fileSize);
        if (this.state.itemsForGrid.length >= 2 && Number(convertKBtoMB) < 9.99) {
          this.setState({
            sendAsMultipleEmailCheckBoxDiv: "",
          });
        }
        for (let i = 0; i < this.state.tempArrayForExternalDocumentGrid.length; i++) {
          if (this.state.tempArrayForExternalDocumentGrid[i].fileSizeInMB >= 10 && this.state.itemsForGrid.length >= 2) {
            this.setState({
              sendAsMultipleEmailCheckBoxDiv: "none",
            });
          }
        }

      }
      this.myfileadditional.value = "";
      this.setState({
        externalComments: "",
        fileSizeDiv: false,
      });
    }
    else {
      this.validator.showMessages();
      this.forceUpdate();
    }
  }
  //temporary array for grid
  private _showProjectDocumentGrid() {
    this.setState({ normalMsgBar: "none", statusMessage: { isShowMessage: false, message: "Please click the project add button", messageType: 1 }, });
    //sizecalculating
    let totalsizeProjects = 0;
    let totalAdditional = 0;
    this.setState({
      documentSelectedDiv: true,
      fileSizeDivForRebind: "none",
    });

    if (this.state.itemsForGrid.length > 0) {
      let duplicate = this.state.itemsForGrid.filter(a => a.publishDoumentlibraryID === this.state.projectDocumentSelectKey);
      if (duplicate.length !== 0) {
        this.setState({
          documentSelectedDiv: false,
          documentSelect: "Already selected document.Please select another.",
          hideGridAddButton: true,
        });
      }
      else {
        if (this.validator.fieldValid("transmitTo") && this.validator.fieldValid("projectDocuments") && this.validator.fieldValid("transmitForKey")) {
          this.validator.hideMessages();
          // let sizeOfDocument;
          if (this.state.transmitTo === "Customer") {
            // let  sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
            // alert((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay/1024).toFixed(3))
            this.state.tempArrayForPublishedDocumentGrid.push({
              publishDoumentlibraryID: this.state.publishDocumentsItemsForGrid[0].ID,
              documentIndexId: this.state.publishDocumentsItemsForGrid[0].DocumentIndex.ID,
              DueDate: moment(this.state.dueDate).format("DD/MM/YYYY"),
              dueDate: this.state.dueDate,
              comments: this.state.comments,
              revision: this.state.publishDocumentsItemsForGrid[0].Revision,
              documentID: this.state.publishDocumentsItemsForGrid[0].DocumentID,
              documentName: this.state.publishDocumentsItemsForGrid[0].DocumentName,
              fileSize: (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(2)),
              fileSizeInMB: (Number((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024) * 0.0009765625).toFixed(2)),
              transmitFor: this.state.transmitFor,
              approvalRequired: this.state.approvalRequired,
              transmitForKey: this.state.transmitForKey,
              temporary: "",
              customerDocumentNo: this.state.publishDocumentsItemsForGrid[0].CustomerDocumentNo,

            });
            console.log("temporaryGrid", this.state.tempArrayForPublishedDocumentGrid);
            this.setState({
              itemsForGrid: this.state.tempArrayForPublishedDocumentGrid,
              showGrid: false,
              projectDocumentSelectKey: "",
              fileSizeDiv: false,
              searchText: "",

            });
            if (this.state.itemsForGrid.length > 0 || this.state.itemsForExternalGrid.length > 0) {
              for (let i = 0; i < this.state.itemsForGrid.length; i++) {
                totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
              }
              for (let k = 0; k < this.state.itemsForExternalGrid.length; k++) {
                totalAdditional = Number(totalAdditional) + Number(this.state.itemsForExternalGrid[k].fileSizeInMB);
              }

              const totalSize = add(totalAdditional, totalsizeProjects);
              const convertKBtoMB = Number(totalSize).toFixed(2);
              this.setState({
                fileSize: Number(convertKBtoMB)
              });
              console.log(this.state.fileSize);
              if (this.state.itemsForGrid.length >= 2 && Number(convertKBtoMB) < 9.99) {
                this.setState({
                  sendAsMultipleEmailCheckBoxDiv: "",
                });
              }
              for (let i = 0; i < this.state.itemsForGrid.length; i++) {
                if (this.state.itemsForGrid[i].fileSizeInMB >= 10 && this.state.itemsForGrid.length >= 2) {
                  this.setState({
                    sendAsMultipleEmailCheckBoxDiv: "none",
                  });
                }
              }

            }

          }
          else if (this.state.transmitTo === "Sub-Contractor") {
            // sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
            // alert((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay/1024).toFixed(3))
            this.state.tempArrayForPublishedDocumentGrid.push({
              publishDoumentlibraryID: this.state.publishDocumentsItemsForGrid[0].ID,
              documentIndexId: this.state.publishDocumentsItemsForGrid[0].DocumentIndex.ID,
              DueDate: moment(this.state.dueDate).format("DD/MM/YYYY"),
              dueDate: this.state.dueDate,
              comments: this.state.comments,
              revision: this.state.publishDocumentsItemsForGrid[0].Revision,
              documentID: this.state.publishDocumentsItemsForGrid[0].DocumentID,
              documentName: this.state.publishDocumentsItemsForGrid[0].DocumentName,
              acceptanceCode: (this.state.sourceDocumentItem === null) ? " " : this.state.publishDocumentsItemsForGrid[0].AcceptanceCode.ID,
              acceptanceCodeTitle: (this.state.sourceDocumentItem === null) ? "" : this.state.publishDocumentsItemsForGrid[0].AcceptanceCode.Title,
              fileSize: (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(2)),
              fileSizeInMB: (Number((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024) * 0.0009765625).toFixed(2)),
              transmitFor: this.state.transmitFor,
              approvalRequired: this.state.approvalRequired,
              transmitForKey: this.state.transmitForKey,
              temporary: "",
              subcontractorDocumentNo: this.state.publishDocumentsItemsForGrid[0].SubcontractorDocumentNo,
            });
            console.log(this.state.tempArrayForPublishedDocumentGrid);
            this.setState({
              itemsForGrid: this.state.tempArrayForPublishedDocumentGrid,
              showGrid: false,
              fileSizeDiv: false,
              searchText: "",
              projectDocumentSelectKey: "",
            });
            if (this.state.itemsForGrid.length > 0 || this.state.itemsForExternalGrid.length > 0) {
              for (let i = 0; i < this.state.itemsForGrid.length; i++) {
                totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
              }
              for (let k = 0; k < this.state.itemsForExternalGrid.length; k++) {
                totalAdditional = Number(totalAdditional) + Number(this.state.itemsForExternalGrid[k].fileSizeInMB);
              }

              const totalSize = add(totalAdditional, totalsizeProjects);
              const convertKBtoMB = Number(totalSize).toFixed(2);
              this.setState({
                fileSize: Number(convertKBtoMB)
              });
              console.log(this.state.fileSize);
              if (this.state.itemsForGrid.length >= 2 && Number(convertKBtoMB) < 9.99) {
                this.setState({
                  sendAsMultipleEmailCheckBoxDiv: "",
                });
              }
              for (let i = 0; i < this.state.itemsForGrid.length; i++) {
                if (this.state.itemsForGrid[i].fileSizeInMB >= 10 && this.state.itemsForGrid.length >= 2) {
                  this.setState({
                    sendAsMultipleEmailCheckBoxDiv: "none",
                  });
                }
              }
            }

          }
        }
        else {
          this.validator.showMessages();
          this.forceUpdate();
        }
      }
    }
    else {
      if (this.validator.fieldValid("transmitTo") && this.validator.fieldValid("projectDocuments") && this.validator.fieldValid("transmitForKey")) {
        this.validator.hideMessages();
        //let  sizeOfDocument;
        if (this.state.transmitTo === "Customer") {
          // sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
          this.state.tempArrayForPublishedDocumentGrid.push({
            publishDoumentlibraryID: this.state.publishDocumentsItemsForGrid[0].ID,
            documentIndexId: this.state.publishDocumentsItemsForGrid[0].DocumentIndex.ID,
            DueDate: moment(this.state.dueDate).format("DD/MM/YYYY "),
            dueDate: this.state.dueDate,
            comments: this.state.comments,
            revision: this.state.publishDocumentsItemsForGrid[0].Revision,
            documentID: this.state.publishDocumentsItemsForGrid[0].DocumentID,
            documentName: this.state.publishDocumentsItemsForGrid[0].DocumentName,
            fileSize: (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(2)),
            fileSizeInMB: (Number((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024) * 0.0009765625).toFixed(2)),
            transmitFor: this.state.transmitFor,
            approvalRequired: this.state.approvalRequired,
            transmitForKey: this.state.transmitForKey,
            temporary: "",
            customerDocumentNo: this.state.publishDocumentsItemsForGrid[0].CustomerDocumentNo,
          });
          console.log("SizeinMb", Number((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024) * 0.0009765625).toFixed(2));
          console.log("ProjectSelectedDocument", this.state.publishDocumentsItemsForGrid);
          this.setState({
            itemsForGrid: this.state.tempArrayForPublishedDocumentGrid,
            showGrid: false,
            fileSize: Number((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024) * 0.0009765625).toFixed(2),
            fileSizeDiv: false,
            searchText: "",
            projectDocumentSelectKey: "",

          });
          if (this.state.itemsForGrid.length > 0 || this.state.itemsForExternalGrid.length > 0) {
            for (let i = 0; i < this.state.itemsForGrid.length; i++) {
              totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
            }
            for (let k = 0; k < this.state.itemsForExternalGrid.length; k++) {
              totalAdditional = Number(totalAdditional) + Number(this.state.itemsForExternalGrid[k].fileSizeInMB);
            }

            const totalSize = add(totalAdditional, totalsizeProjects);
            const convertKBtoMB = Number(totalSize).toFixed(2);
            this.setState({
              fileSize: Number(convertKBtoMB)
            });
            console.log(this.state.fileSize);
          }
        }
        else if (this.state.transmitTo === "Sub-Contractor") {
          // sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
          // alert((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay/1024).toFixed(3))
          this.state.tempArrayForPublishedDocumentGrid.push({
            publishDoumentlibraryID: this.state.publishDocumentsItemsForGrid[0].ID,
            documentIndexId: this.state.publishDocumentsItemsForGrid[0].DocumentIndex.ID,
            DueDate: moment(this.state.dueDate).format("DD/MM/YYYY"),
            dueDate: this.state.dueDate,
            comments: this.state.comments,
            revision: this.state.publishDocumentsItemsForGrid[0].Revision,
            documentID: this.state.publishDocumentsItemsForGrid[0].DocumentID,
            documentName: this.state.publishDocumentsItemsForGrid[0].DocumentName,
            acceptanceCode: (this.state.sourceDocumentItem === null) ? null : this.state.publishDocumentsItemsForGrid[0].AcceptanceCode.ID,
            acceptanceCodeTitle: (this.state.sourceDocumentItem === null) ? "" : this.state.publishDocumentsItemsForGrid[0].AcceptanceCode.Title,
            fileSize: (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(2)),
            fileSizeInMB: (Number((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024) * 0.0009765625).toFixed(2)),
            transmitFor: this.state.transmitFor,
            approvalRequired: this.state.approvalRequired,
            transmitForKey: this.state.transmitForKey,
            temporary: "",
            subcontractorDocumentNo: this.state.publishDocumentsItemsForGrid[0].SubcontractorDocumentNo,
          });
          console.log(this.state.tempArrayForPublishedDocumentGrid);
          this.setState({
            itemsForGrid: this.state.tempArrayForPublishedDocumentGrid,
            showGrid: false,
            fileSize: Number((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024) * 0.0009765625).toFixed(2),
            fileSizeDiv: false,
            searchText: "",
            projectDocumentSelectKey: "",
          });
          if (this.state.itemsForGrid.length > 0 || this.state.itemsForExternalGrid.length > 0) {
            for (let i = 0; i < this.state.itemsForGrid.length; i++) {
              totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
            }
            for (let k = 0; k < this.state.itemsForExternalGrid.length; k++) {
              totalAdditional = Number(totalAdditional) + Number(this.state.itemsForExternalGrid[k].fileSizeInMB);
            }

            const totalSize = add(totalAdditional, totalsizeProjects);
            const convertKBtoMB = Number(totalSize).toFixed(2);
            this.setState({
              fileSize: Number(convertKBtoMB)
            });
            console.log(this.state.fileSize);
          }
          if (this.state.itemsForGrid.length >= 2) {
            this.setState({
              sendAsMultipleEmailCheckBoxDiv: "",
            });
          }
        }
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
    this.setState({
      comments: "",
      transmitForKey: null,
    });
  }
  //Save as draft 

  private _forCalculatingSize() {
    let totalsizeProjects = 0;
    let totalAdditional = 0;
    if (this.state.itemsForGrid.length > 0 || this.state.itemsForExternalGrid.length > 0) {
      for (let i = 0; i < this.state.itemsForGrid.length; i++) {
        totalsizeProjects = Number(totalsizeProjects) + Number(this.state.itemsForGrid[i].fileSizeInMB);
      }
      for (let k = 0; k < this.state.itemsForExternalGrid.length; k++) {
        totalAdditional = Number(totalAdditional) + Number(this.state.itemsForExternalGrid[k].fileSizeInMB);
      }

      const totalSize = add(totalAdditional, totalsizeProjects);
      const convertKBtoMB = Number(totalSize).toFixed(2);
      this.setState({
        fileSize: Number(convertKBtoMB)
      });
      console.log(this.state.fileSize);
    }
    if (Number(this.state.fileSize) > 10 && (this.state.sendAsSharedFolder === true)) {
      this.setState({
        normalMsgBar: "none",
        statusMessage: { isShowMessage: false, message: this.state.transmittalNo, messageType: 4 },

      });
    }
    else if (Number(this.state.fileSize) < 10 && (this.state.sendAsSharedFolder === false)) {
      this.setState({
        normalMsgBar: "none",
        statusMessage: { isShowMessage: false, message: this.state.transmittalNo, messageType: 4 },

      });
    }
  }

  //transmittal id generation
  // public async _trannsmittalIDGeneration() {
  //   let prefix;
  //   let separator;
  //   let sequenceNumber;
  //   let title;
  //   let counter;
  //   let transmittalID;
  //   let transmitTo;
  //   if (this.state.transmitTo === "Customer") { transmitTo = "Outbound Customer"; }
  //   else if (this.state.transmitTo === "Sub-Contractor") { transmitTo = "Outbound Sub-contractor"; }
  //   await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalIdSettingsListName).items.filter("TransmittalCategory eq '" + transmitTo + "' and(TransmittalType eq '" + this.state.transmittalType + "')").get().then((transmittalIdSettingsItems: { ID: any; }[]) => {
  //     console.log("transmittalIdSettingsItems", transmittalIdSettingsItems);
  //     prefix = transmittalIdSettingsItems[0].Prefix;
  //     separator = transmittalIdSettingsItems[0].Separator;
  //     sequenceNumber = transmittalIdSettingsItems[0].SequenceNumber;
  //     title = transmittalIdSettingsItems[0].Title;
  //     counter = transmittalIdSettingsItems[0].Counter;
  //     let increment = counter + 1;
  //     var incrementValue = increment.toString();
  //     this._transmittalSequenceNumber(incrementValue, sequenceNumber);
  //     transmittalID = prefix + separator + title + separator + this.state.projectNumber + separator + this.state.incrementSequenceNumber;
  //     console.log("transmittalID", transmittalID);
  //     this.setState({
  //       transmittalNo: transmittalID,
  //     });
  //     //counter updation
  //     sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalIdSettingsListName).items.getById(transmittalIdSettingsItems[0].ID).update({
  //       Counter: increment,
  //     });
  //   });
  // }
  //protected async triggerProjectPermissionFlow(PostUrl) {
  //   //alert("triggerProjectPermissionFlow")
  //   let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
  //   // alert("In function");
  //   // alert(transmittalID);
  //   const postURL = PostUrl;
  //   const requestHeaders: Headers = new Headers();
  //   requestHeaders.append("Content-type", "application/json");
  //   const body: string = JSON.stringify({
  //     'PermissionTitle': 'Project_UnlockTransmittal',
  //     'SiteUrl': siteUrl,
  //     'CurrentUserEmail': this.props.context.pageContext.user.email
  //   });
  //   const postOptions: IHttpClientOptions = {
  //     headers: requestHeaders,
  //     body: body
  //   };
  //   let responseText: string = "";
  //   // let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  //   // let responseJSON = await response.json();
  //   // responseText = JSON.stringify(responseJSON);
  //   // console.log(responseJSON);
  //   // if (response.ok) {
  //   //   console.log(responseJSON['Status']);
  //   //   if (responseJSON['Status'] === "Valid") {
  //   //     // this.setState({
  //   //     //   loaderDisplay: "none",
  //   //     //   webpartView: "",
  //   //     // });
  //   //     //this._queryParamGetting();
  //   //     this.permissionForRecall = "Yes";
  //   //     this.setState({
  //   //       hideButtonAfterSubmit: "none",
  //   //       hideUnlockButton: "",
  //   //     });
  //   //   }
  //   //   else {
  //   //     // this.setState({
  //   //     //   webpartView: "none",
  //   //     //   loaderDisplay: "none",
  //   //     //   accessDeniedMsgBar: "",
  //   //     //   statusMessage: { isShowMessage: true, message: "You are not permitted to perform this operations", messageType: 1 },
  //   //     // });
  //   //     // setTimeout(() => {
  //   //     //   window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //   //     // }, 20000);
  //   //   }

  //   // }
  //   // else { }

  // }
  // protected async triggerOutboundTransmittal(transmittalID) {
  //   let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
  //   const postURL = this.postUrl;
  //   const requestHeaders: Headers = new Headers();
  //   requestHeaders.append("Content-type", "application/json");
  //   const body: string = JSON.stringify({
  //     'SiteURL': siteUrl,
  //     'TransmittalNo': transmittalID,
  //     'ProjectName': this.state.projectName,
  //     'ContractNumber': this.state.contractNumber,
  //     'ProjectNumber': this.state.projectNumber,
  //     'CoverLetterNeeded': (this.state.coverLetterNeeded === true ? "Yes" : "NO"),
  //     'InternalContactsEmails': this.state.internalContactsEmail,
  //     'InternalContactsDisplayNames': this.state.internalCCContactsDisplayNameForPreview,
  //     'OutboundTransmittalDetails': this.state.itemsForGrid
  //   });
  //   const postOptions: IHttpClientOptions = {
  //     headers: requestHeaders,
  //     body: body
  //   };
  //   let responseText: string = "";
  //   // let response = await this.props.context.httpClient.post(postURL, HttpClient.configurations.v1, postOptions);
  //   // let responseJSON = await response.json();
  //   // responseText = JSON.stringify(responseJSON);
  //   // console.log(responseJSON);
  //   // if (response.ok) {
  //   //   // alert(response.text);
  //   //   if (responseJSON['Status'] === "MailSend") {
  //   //     this.setState({
  //   //       hideButtonAfterSubmit: "none",
  //   //       hideUnlockButton: "none",
  //   //       normalMsgBar: "",
  //   //       spinnerDiv: "none",
  //   //       statusMessage: { isShowMessage: true, message: "Transmittal Send Successfully", messageType: 4 },
  //   //     });
  //   //     setTimeout(() => {
  //   //       window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //   //     }, 10000);
  //   //   }
  //   //   else {

  //   //   }
  //   // }
  //   // else { }

  // }
  //incrementing transmittal id sequence number
  private _transmittalSequenceNumber(incrementValue: any, sequenceNumber: number) {
    let incrementSequenceNumber = incrementValue;
    while (incrementSequenceNumber.length < sequenceNumber)
      incrementSequenceNumber = "0" + incrementSequenceNumber;
    console.log(incrementSequenceNumber);
    this.setState({
      incrementSequenceNumber: incrementSequenceNumber,
    });
  }
  private getCheckboxesValue(event: any) {
    let value = "";
    console.log(event);
    this.state.checkedItems.forEach((hobbyinfo: any, index: string) => {
      if (hobbyinfo) {
        value = value === "" ? index : value + "," + index;
      }
    });
    // alert(value);
  }
  private handleSubmit(event: { preventDefault: () => void; }) {
    console.log(this.state);
    event.preventDefault();
  }
  private handleChange(event: { target: { checked: any; value: any; }; }) {
    let isChecked = event.target.checked;
    let item = event.target.value;

    this.setState(prevState => ({ checkedItems: prevState.checkedItems.set(item, isChecked) }));
    console.log(this.state.checkedItems.item);
  }
  //Delete button click
  private _openDeleteConfirmation = (items: { [x: string]: any; outboundDetailsID: any; additionalDocumentID: any; }, key: number, type: string) => {
    if (this.transmittalID === "" || this.transmittalID === null) {
      this.setState({
        deleteConfirmation: "",
        confirmDeleteDialog: false,
      });
      this.validator.hideMessages();
      console.log(items[key]);
      if (type === "ProjectDocuments") {
        this.setState({ TypeOFDelete: "ProjectDocuments" });
        this.keyForDelete = key;
      } else if (type === "AdditionalDocuments") {
        this.setState({ TypeOFDelete: "AdditionalDocuments" });
        this.keyForDelete = key;
      }
    }
    else {
      this.setState({
        deleteConfirmation: "",
        confirmDeleteDialog: false,
        tempDocIndexIDForDelete: items.outboundDetailsID,
      });
      this.validator.hideMessages();
      console.log(items[key]);
      if (type === "ProjectDocuments") {
        // alert(items.outboundDetailsID);
        this.setState({ TypeOFDelete: "ProjectDocuments" });
        this.keyForDelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.outboundDetailsID,
        });
      } else if (type === "AdditionalDocuments") {
        // alert("additionalid" + items.additionalDocumentID);
        this.setState({ TypeOFDelete: "AdditionalDocuments" });
        this.keyForDelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.additionalDocumentID,
        });
      }
    }

  }
  private _confirmDeleteItem = async (docID: any, items: string, key: string) => {
    if (this.transmittalID == "" || this.transmittalID == null) {
      this.setState({
        confirmDeleteDialog: true,
        deleteConfirmation: "none"
      });
      this.validator.hideMessages();
      if (this.state.TypeOFDelete == "ProjectDocuments") {
        this.itemDeleteFromGrid(items, key);
      }
      else if (this.state.TypeOFDelete == "AdditionalDocuments") {
        this.itemDeleteFromExternalGrid(items, key);
      }

    }
    else {
      this.setState({
        confirmDeleteDialog: true,
        deleteConfirmation: "none"
      });
      this.validator.hideMessages();
      console.log(items[key]);

      if (this.typeForDelete == "ProjectDocuments") {
        // alert(docID);

        // if (docID) {
        //   let list = sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName);
        //   await list.items.getById(parseInt(docID)).delete();
        //   let selectHeaderItems = "Id,DocumentIndex/ID,DocumentIndex/Title,DueDate,SentComments,Revision,Title,Size,TransmittedFor/ID,TransmittedFor/Title,Temporary,TransmittalHeader/ID,DocumentLibraryID,ID,ApprovalRequired";
        //   sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName).items.select(selectHeaderItems).expand("DocumentIndex,TransmittedFor,TransmittalHeader").filter("TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ").get().then(outboundTransmittalDetailsListName => {
        //     console.log("outboundTransmittalDetailsListName", outboundTransmittalDetailsListName);
        //     this.setState({
        //       currentOutboundDetailItem: outboundTransmittalDetailsListName,
        //     });
        //   });
        //   this.setState({
        //     itemsForGrid: this.state.itemsForGrid,
        //   });
        // }
        this.itemDeleteFromGrid(items, key);
      }
      else if (this.typeForDelete == "AdditionalDocuments") {
        // if (docID) {
        //   let list = sp.web.getList(this.props.siteUrl + "/" + this.props.outboundAdditionalDocumentsListName + "/");
        //   await list.items.getById(parseInt(docID)).delete();
        //   sp.web.getList(this.props.siteUrl + "/" + this.props.outboundAdditionalDocumentsListName + "/").items.filter("TransmittalIDId eq '" + this.transmittalID + "' ").get().then(listItems => {
        //     this.setState({
        //       currentOutboundAdditionalItem: listItems,
        //     });
        //   });
        //   this.setState({
        //     itemsForExternalGrid: this.state.itemsForExternalGrid,
        //   });
        // }
        this.itemDeleteFromExternalGrid(items, key);
      }
    }
  }
  //deleting
  public itemDeleteFromGrid(items: any, key: any) {
    console.log(items);
    const updatedFiles = this.state.itemsForGrid.filter((item, index) => index !== key);
    this.setState({
      itemsForGrid: updatedFiles,
      documentSelectedDiv: true,
      projectDocumentSelectKey: "",
    });
    console.log("after removal", this.state.itemsForGrid);
    console.log(items.fileSize);
    this._forCalculatingSize();
    //for project documents
    for (let i = 0; i < this.state.itemsForGrid.length; i++) {
      if (this.state.itemsForGrid[i].fileSizeInMB >= 10 && this.state.itemsForGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "none",
        });
      }
      else if (this.state.itemsForGrid[i].fileSizeInMB <= 10 && this.state.itemsForGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "",
        });
      }
    }
    //for additional
    for (let i = 0; i < this.state.itemsForExternalGrid.length; i++) {
      if (this.state.itemsForExternalGrid[i].fileSizeInMB >= 10 && this.state.itemsForExternalGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "none",
        });
      }
      else if (this.state.itemsForExternalGrid[i].fileSizeInMB <= 10 && this.state.itemsForExternalGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "",
        });
      }
    }
  }
  public itemDeleteFromExternalGrid(items: any, key: any) {
    this.state.itemsForExternalGrid.splice(key, 1);
    console.log("after removal", this.state.itemsForExternalGrid);
    console.log(items.fileSize);
    this.setState({
      itemsForExternalGrid: this.state.itemsForExternalGrid,
    });
    this._forCalculatingSize();
    this.setState({
      externalComments: "",
    });
    this.myfileadditional.value = "";
    //for additiona multiple doc checkbox 
    for (let i = 0; i < this.state.itemsForExternalGrid.length; i++) {
      if (this.state.itemsForExternalGrid[i].fileSizeInMB >= 10 && this.state.itemsForGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "none",
        });
      }
      else if (this.state.itemsForExternalGrid[i].fileSizeInMB <= 10 && this.state.itemsForGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "",
        });
      }
    }
    for (let i = 0; i < this.state.itemsForGrid.length; i++) {
      if (this.state.itemsForGrid[i].fileSizeInMB >= 10 && this.state.itemsForGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "none",
        });
      }
      else if (this.state.itemsForGrid[i].fileSizeInMB <= 10 && this.state.itemsForGrid.length >= 2) {
        this.setState({
          sendAsMultipleEmailCheckBoxDiv: "",
        });
      }
    }
  }
  //confirm cancel button click
  private _confirmNoCancel = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmCancelDialog: true,
      deleteConfirmation: "none",
      confirmDeleteDialog: true,
      recallConfirmMsg: true,
      recallConfirmMsgDiv: "none",
    });
    this.validator.hideMessages();
  }
  //confirm cancel button click
  private _cancelConfirmYes = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmCancelDialog: true,
      documentSelectedDiv: true,
      documentSelect: "",
    });

    window.location.replace(window.location.protocol + "//" + window.location.hostname + "/" + this.props.siteUrl);
    this.validator.hideMessages();
  }
  private dialogStyles = { main: { maxWidth: 500 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to delete?',
  };
  private dialogCancelContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to Cancel?',
    //subText: '<b>Do you want to cancel? </b> ',
  };
  public _coverLetterNeeded = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ coverLetterNeeded: true, });
    }
    else if (!isChecked) { this.setState({ coverLetterNeeded: false, }); }
  }
  private modalProps = {
    isBlocking: true,
  };
  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      cancelConfirmMsg: "none",
      confirmCancelDialog: true,
      confirmDeleteDialog: true,
      recallConfirmMsg: true,
      recallConfirmMsgDiv: "none",
      deleteConfirmation: "none",
    });

  }
  public _onSendAsSharedFolder = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ sendAsSharedFolder: true, recieveInSharedFolder: true, });
      if (this.state.normalMsgBar === "") {
        this.setState({
          normalMsgBar: "none",
          statusMessage: { isShowMessage: false, message: "Recalled" + this.state.transmittalNo, messageType: 4 },
        });
      }
    }
    else if (!isChecked) { this.setState({ sendAsSharedFolder: false, }); }


  }
  public _onRecieveInSharedFolder = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ recieveInSharedFolder: true, });
    }
    else if (!isChecked) { this.setState({ recieveInSharedFolder: false, }); }
  }
  //for To fields
  // private _onDrpdwnCntact = async (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
  //   let checkedContacts: string;
  //   let checkedContactsDisplay: string;
  //   let contacts: any;
  //   if (option.selected) {
  //     contacts = {
  //       key: option.key,
  //       text: option.text.split("<"),//splitting the < to split mail id to inserting ToName and CCName
  //     };
  //     this.emailsSelectedTo.push(contacts.key);
  //     this.contactToDisplay.push(contacts.text[0]);
  //     checkedContacts = (this.emailsSelectedTo).toString();
  //     checkedContactsDisplay = (this.contactToDisplay).toString();
  //     let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
  //     this.setState({
  //       selectedContactsTo: checkedContactsSemicolonAttached,
  //       selectedContactsToDisplayName: checkedContactsDisplay,
  //     });
  //     console.log("checkedContacts", checkedContacts);
  //   }
  //   else {
  //     this.emailsSelectedTo.splice(index, 1);
  //     this.contactToDisplay.splice(index, 1);
  //     let newarray = this.emailsSelectedTo.filter(element => element !=== option.key);
  //     checkedContacts = (newarray).toString();
  //     let splittedName = (option.text).split("<");
  //     console.log(splittedName[0]);
  //     let newarrayContactDisplay = this.contactToDisplay.filter(element => element !=== splittedName[0]);
  //     checkedContactsDisplay = (newarrayContactDisplay).toString();
  //     this.contactToDisplay = newarrayContactDisplay;
  //     console.log("afterFilter", this.contactToDisplay);
  //     let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
  //     this.setState({
  //       selectedContactsTo: checkedContactsSemicolonAttached,
  //       selectedContactsToDisplayName: checkedContactsDisplay,
  //     });
  //     console.log("checkedContacts", checkedContacts);
  //   }
  // }
  // //for CC fields
  // private _onDrpdwnCCContact = async (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
  //   let checkedContacts: string;
  //   let checkedContactsDisplay: string;
  //   if (option.selected) {
  //     let contacts = {
  //       key: option.key,
  //       text: option.text.split("<"),
  //     };
  //     this.emailsSelectedCC.push(contacts.key);
  //     this.contactCCDisplay.push(contacts.text[0]);
  //     checkedContacts = (this.emailsSelectedCC).toString();
  //     checkedContactsDisplay = (this.contactCCDisplay).toString();
  //     let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
  //     this.setState({
  //       selectedContactsCC: checkedContactsSemicolonAttached,
  //       selectedContactsCCDisplayName: checkedContactsDisplay,
  //     });
  //     console.log("checkedContactsCC", checkedContacts);
  //   }
  //   else {

  //     this.emailsSelectedCC.splice(index, 1);
  //     console.log("beforeFilter", this.contactCCDisplay);
  //     this.contactCCDisplay.splice(index, 1);
  //     let newarray = this.emailsSelectedCC.filter(element => element !== option.key);
  //     checkedContacts = (newarray).toString();
  //     let splittedName = (option.text).split("<");
  //     console.log(splittedName[0]);
  //     let newarrayContactDisplay = this.contactCCDisplay.filter(element => element !== splittedName[0]);
  //     checkedContactsDisplay = (newarrayContactDisplay).toString();
  //     this.contactCCDisplay = newarrayContactDisplay;
  //     console.log("afterFilter", this.contactCCDisplay);
  //     let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
  //     this.setState({
  //       selectedContactsCC: checkedContactsSemicolonAttached,
  //       selectedContactsCCDisplayName: checkedContactsDisplay,
  //     });
  //     console.log("checkedContactsCC", checkedContacts);
  //   }
  // }
  // Internal CC People Picker Change
  public _selectedInternalCCContacts = (items: any[]) => {
    let getSelectedInternalID = [];
    let getSelectedInternalDisplayName = [];
    let getSelectedInternalEmailID = [];
    items.forEach(item => {
      getSelectedInternalID.push(items[item].id);
      getSelectedInternalDisplayName.push(items[item].text);
      getSelectedInternalEmailID.push(items[item].secondaryText);
    })

    let displayInternalName = getSelectedInternalDisplayName.toString();
    let InternalEmailID = getSelectedInternalEmailID.toString();
    let InternalEmailIDSemicolonAttached = replaceString(InternalEmailID, ',', ';');
    this.setState({ internalCCContacts: getSelectedInternalID, internalCCContactsDisplayNameForPreview: displayInternalName, internalContactsEmail: InternalEmailIDSemicolonAttached });
  }
  public _onSendAsMultipleFolder = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ sendAsMultipleFolder: true, });
    }
    else if (!isChecked) { this.setState({ sendAsMultipleFolder: false, }); }
  }
  private dialogContentRecallProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: "Do you want to Recall ?",
  };
  private _recallTransmittalConfirmation() {
    this.setState({
      recallConfirmMsgDiv: "",
      recallConfirmMsg: false,
    });
  }

  // sending Email for owners
  // private async _sendAnEmailUsingMSGraph(email: any, type: string, name: any, documentName: any): Promise<void> {
  //   let Subject;
  //   let Body;
  //   const emailNoficationSettings: any[] = await this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.emailNotificationSettings)
  //     .items.filter("Title eq '" + type + "'").get();
  //   Subject = emailNoficationSettings[0].Subject;
  //   Body = emailNoficationSettings[0].Body;
  //   //Replacing the email body with current values
  //   let replacedSubject1 = replaceString(Subject, '[DocumentName]', documentName);
  //   let replacedSubject = replaceString(replacedSubject1, '[TransmittalNo]', this.state.transmittalNo);
  //   let replaceRequester = replaceString(Body, '[Sir/Madam],', name);
  //   let replaceBody = replaceString(replaceRequester, '[DocumentName]', documentName);
  //   let replacelink = replaceString(replaceBody, '[TransmittalNo]', this.state.transmittalNo);

  //   let FinalBody = replacelink;
  //   if (email) {
  //     //Create Body for Email  
  //     let emailPostBody: any = {
  //       "message": {
  //         "subject": replacedSubject,
  //         "body": {
  //           "contentType": "HTML",
  //           "content": FinalBody
  //         },
  //         "toRecipients": [
  //           {
  //             "emailAddress": {
  //               "address": email
  //             }
  //           }
  //         ],
  //       }
  //     };
  //     //Send Email uisng MS Graph  
  //     this.props.context.msGraphClientFactory 
  //     .getClient("3")
  //     .then((client): void => {
  //         client
  //             .api('/me/sendMail')
  //             .post(emailPostBody);
  //     });
  //   }
  //   // }
  // }
  public _uploadadditional(e: { target: { value: any; }; currentTarget: { value: any; }; }) {
    this.myfileadditional = e.target.value;
    let documentNameExtension;
    console.log(this.myfileadditional);
    console.log(e.target.value);
    console.log(e.currentTarget.value);
    let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
    let splitted = myfile.name.split(".");
    console.log(splitted);
    console.log(splitted.length);
    console.log(splitted[splitted.length - 1]);
    for (let r = 0; r < splitted.length - 1; r++) {
      documentNameExtension = splitted.slice(0, -1).join('.') + "_TR00011" + '.' + splitted[splitted.length - 1];
    }
    // documentNameExtension = splitted[0] + "_TR00011" + '.' + splitted[splitted.length - 1];
    console.log(documentNameExtension);
    //let docName = documentNameExtension;
  }
  private setSelectedContactsTo = async (option: any[]) => {
    let checkedContacts: string;
    let checkedContactsDisplay: string;
    this.emailsSelectedTo = [];
    this.contactToDisplay = [];
    this.setState({
      searchContactsTo: [],
      selectedContactsTo: " ",
      selectedContactsToDisplayName: "",
    });
    let selectedContactsIdArray = [];
    let selectedContactsNameArray = [];
    console.log("option " + option)

    for (let i = 0; i < option.length; i++) {
      selectedContactsIdArray.push(option[i].value);
      selectedContactsNameArray.push(option[i].label.split("<")[0]);
    }
    console.log("email", selectedContactsIdArray);
    console.log("Name", selectedContactsNameArray);
    this.emailsSelectedTo.push(selectedContactsIdArray);
    this.contactToDisplay.push(selectedContactsNameArray);
    checkedContacts = (this.emailsSelectedTo).toString();
    checkedContactsDisplay = (this.contactToDisplay).toString();
    let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
    this.setState({
      selectedContactsTo: checkedContactsSemicolonAttached,
      selectedContactsToDisplayName: checkedContactsDisplay,
      selectedVendor: option,
    });
    console.log("checkedContacts", checkedContacts);
    this.setState({
      selectedContactsToName: option,
      searchContactsTo: selectedContactsIdArray
    });
  }
  private setSelectedContactsCC = async (option: any[]) => {
    let checkedContacts: string;
    let checkedContactsDisplay: string;
    this.emailsSelectedCC = [];
    this.contactCCDisplay = [];
    let selectedContactsCCArray = [];
    let selectedContactsCCNameArray = [];
    this.setState({
      searchContactsCC: [],
      selectedContactsCC: " ",
      selectedContactsCCDisplayName: "",
    });

    for (let i = 0; i < option.length; i++) {
      selectedContactsCCArray.push(option[i].value);
      selectedContactsCCNameArray.push(option[i].label.split("<")[0]);
    }
    this.emailsSelectedCC.push(selectedContactsCCArray);
    this.contactCCDisplay.push(selectedContactsCCNameArray);
    checkedContacts = (this.emailsSelectedCC).toString();
    checkedContactsDisplay = (this.contactCCDisplay).toString();
    const checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');

    this.setState({
      selectedContactsCC: checkedContactsSemicolonAttached,
      selectedContactsCCDisplayName: checkedContactsDisplay,
    });
    console.log("checkedContacts", checkedContacts);
    this.setState({
      selectedContactsCCName: option,
      searchContactsCC: selectedContactsCCArray
    });
  }
  // from current site
  //page load from project information list
  public projectInformation = async () => {
    await this._Service.getListItems(this.props.siteUrl, this.props.projectInformationListName)
      .then((projectInformation: any[]) => {
        if (projectInformation.length > 0) {
          projectInformation.forEach(PI => {
            if (PI.Key === "ProjectName") {
              this.setState({
                projectName: PI.Title,
              });
            }
            if (PI.Key === "Customer") {
              this.setState({
                customerName: PI.Title,
              });
            }
            if (PI.Key === "ContractNumber") {
              this.setState({
                contractNumber: PI.Title,
              });
            }
            if (PI.Key === "ApprovalCycle") {
              this.setState({
                approvalLifeCycle: PI.Title,
              });
              const dueDate = new Date();
              const days = PI.Title;
              console.log(Number(days));
              dueDate.setDate(dueDate.getDate() + Number(days));
              this.setState({
                hideDueDate: false,
                dueDate: dueDate,
                dueDateForBindingApprovalLifeCycle: dueDate,
              });
            }
            if (PI.Key === "ProjectNumber") {
              this.setState({
                projectNumber: PI.Title,
              });
            }
            if (PI.Key === "CustomerID") {
              this.setState({
                customerId: PI.Title,
              });
            }
            if (PI.Key === "LegalEntityId") {
              this.setState({
                legalId: PI.Title,
              });
            }

          })
        }
      });
  }
  //Current User
  private async _currentUser() {
    this._Service.getCurrentUserId().then((currentUser: { Id: any; }) => {
      this.setState({
        currentUser: currentUser.Id,
      });
    });
  }
  //transmit for
  private _transmitForBind() {
    const transmitForArray: { key: any; text: any; }[] = [];
    this._Service.getTransmitFor(this.props.siteUrl, this.props.transmittalCodeSettingsListName)
      .then((transmitFor: string | any[]) => {
        for (let i = 0; i < transmitFor.length; i++) {
          const transmitForItemdata = {
            key: transmitFor[i].ID,
            text: transmitFor[i].Title
          };
          transmitForArray.push(transmitForItemdata);
        }
        this.setState({
          transmitForItems: transmitForArray
        });
      });
  }
  //for customers documents from published docs
  public async _loadPublishDocuments(item) {
    const publishedDocumentArray: { value: any; label: any; }[] = [];
    let transmitForItemdata;
    let filter;
    const publishedDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.publishDocumentLibraryName;
    if (item !== "") {
      filter = "TransmittalStatus ne 'Ongoing' and (TransmittalDocument ne '" + false + "') and (DocumentStatus eq 'Active') and (WorkflowStatus eq 'Published') and (Category eq '" + this.state.documentFilters['Category'] + "')";
    }
    else {
      filter = "TransmittalStatus ne 'Ongoing' and (TransmittalDocument ne '" + false + "') and (DocumentStatus eq 'Active') and (WorkflowStatus eq 'Published') "
    }
    this._Service.getLibraryItems(publishedDocumentsDl, filter)
      .then(async (publishDocumentsItems) => {
        console.log("PublishDocumentForCustomerCount", publishDocumentsItems.length);
        this.sortedArray = _.orderBy(publishDocumentsItems, 'FileLeafRef', ['asc']);
        if (publishDocumentsItems.length > 0) {
          this._Service.getDIItems(this.props.context.pageContext.web.serverRelativeUrl, "DocumentIndex")
            .then((DIndexItems: any[]) => {
              console.log("PublishDocumentForCustomerFromIndex", DIndexItems.length);
              const filteredIndexItems = this.sortedArray.filter((item) =>
                DIndexItems.some((pdItem: { ID: any; }) => pdItem.ID === item.DocumentIndexId)
              );
              if (filteredIndexItems.length > 0) {
                filteredIndexItems.forEach((filteredItems: any) => {
                  transmitForItemdata = {
                    value: filteredItems.ID,
                    label: filteredItems.DocumentName,
                    FileLeafRef: filteredItems.FileLeafRef,
                    DocumentID: filteredItems.DocumentID,
                    Revision: filteredItems.Revision,
                    FileSizeDisplay: filteredItems.FileSizeDisplay,
                    DocumentName: filteredItems.DocumentName,
                    DocumentIndexId: filteredItems.DocumentIndexId,
                    WorkflowStatus: filteredItems.WorkflowStatus,
                    CustomerDocumentNo: filteredItems.CustomerDocumentNo,
                    SubcontractorDocumentNo: filteredItems.SubcontractorDocumentNo,
                  };
                  publishedDocumentArray.push(transmitForItemdata);
                });
                this.setState({
                  searchDocuments: publishedDocumentArray
                });
              }
              else {
                this.setState({
                  searchDocuments: publishedDocumentArray
                });
                if (publishedDocumentArray.length === 0) {
                  this.setState({
                    documentSelectedDiv: false,
                    documentSelect: "No documents for transmittal "
                  });
                }
              }
            });
        }
        else {
          console.log("No documents for transmittal");
          this.setState({
            documentSelectedDiv: false,
            documentSelect: "No documents for transmittal "
          });
        }
      }).catch((err: any) => {
        console.log("Error = ", err);
      });
  }   //for subcontractors  documents from published docs
  public async _loadSourceDocuments() {
    //for customer values from sourceDocuments
    const sourceDocumentArray: { value: any; label: any; }[] = [];
    const sourceDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.sourceDocumentLibraryName;
    this._Service.getSourceLibraryItems(sourceDocumentsDl)
      .then((sourceDocumentArrayItems: string | any[]) => {
        console.log("SourceDocumentForCustomer", sourceDocumentArrayItems.length);
        if (sourceDocumentArrayItems.length > 0) {
          this.sortedArray = _.orderBy(sourceDocumentArrayItems, 'FileLeafRef', ['asc']);
          this.sortedArray.forEach(sourceItems => {
            const transmitForItemdata = {
              value: sourceItems.ID,
              label: sourceItems.DocumentName
            };
            sourceDocumentArray.push(transmitForItemdata);
          });
          this.setState({
            searchDocuments: sourceDocumentArray
          });
        }
        else {
          console.log("No documents for transmittal");
          this.setState({
            documentSelectedDiv: false,
            documentSelect: "No documents for transmittal "
          });
        }
      }).catch((err: any) => {
        console.log("Error = ", err);
        this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: false, message: err, messageType: 1 }, });
      });

  }
  //for subcontractors letters documents from published docs
  public async _loadSourceDocumentsForLetter() {
    // let temDoc: [];
    const publishedDocumentArray: { value: any; label: any; }[] = [];
    let transmitForItemdata;
    const publishedDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.publishDocumentLibraryName;
    this._Service.getLibraryItems(publishedDocumentsDl, "")
      .then(async (publishDocumentsItems: string | any[]) => {
        console.log("PublishDocumentForCustomerCount", publishDocumentsItems.length);
        this.sortedArray = _.orderBy(publishDocumentsItems, 'FileLeafRef', ['asc']);
        if (publishDocumentsItems.length > 0) {
          this._Service.getDIItems(this.props.context.pageContext.web.serverRelativeUrl, "DocumentIndex")
            .then((DIndexItems: any[]) => {
              console.log("PublishDocumentFormIndex", DIndexItems.length);
              const filteredIndexItems = this.sortedArray.filter((item) =>
                DIndexItems.some((pdItem: { ID: any; }) => pdItem.ID === item.DocumentIndexId)
              );
              if (filteredIndexItems.length > 0) {
                filteredIndexItems.forEach((filteredItems: any) => {
                  if (filteredItems.Category === "Project - Official Letter") {
                    transmitForItemdata = {
                      value: filteredItems.ID,
                      label: filteredItems.DocumentName
                    };
                    publishedDocumentArray.push(transmitForItemdata);
                  }
                });
                this.setState({
                  searchDocuments: publishedDocumentArray
                });
                if (publishedDocumentArray.length === 0) {
                  this.setState({
                    documentSelectedDiv: false,
                    documentSelect: "No Project - Official Letter  for transmittal "
                  });
                }
              }
              else {
                this.setState({ searchDocuments: publishedDocumentArray });
                if (publishedDocumentArray.length === 0) {
                  this.setState({
                    documentSelectedDiv: false,
                    documentSelect: "No documents for transmittal "
                  });
                }
              }
            });
        }
        else {
          //alert("No documents all transmittal status is ONGOING");
          console.log("No documents for transmittal");
          this.setState({
            documentSelectedDiv: false,
            documentSelect: "No documents for transmittal "
          });
        }
      }).catch((err: any) => {
        console.log("Error = ", err);
      });


  }
  //project documents grid binding
  private async _onDocumentClick(ID) {
    let selectedDocuments = [];
    selectedDocuments.push(...ID);
    console.log(selectedDocuments);
    // this.setState({ projectDocumentSelectKey: ID.value, documentSelectedDiv: true, hideGridAddButton: false, });
    this.setState({
      searchDiv: "none",
    });
    // if (this.state.transmitTo == "Customer") {
    //   sp.web.getList(this.props.siteUrl + "/" + this.props.publishDocumentLibraryName)
    //     .items.select("ID,DocumentID,FileSizeDisplay,DocumentName,Revision,DocumentIndex/ID,CustomerDocumentNo,SubcontractorDocumentNo")
    //     .expand("DocumentIndex").filter("ID eq '" + ID.value + "' ").get().then((publishDocumentsItemsForGrid: any) => {
    //       console.log("publishDocumentsItemsForGrid", publishDocumentsItemsForGrid);
    //       this.setState({
    //         publishDocumentsItemsForGrid: publishDocumentsItemsForGrid
    //       });
    //     });
    // }
    // else if (this.state.transmitTo == "Sub-Contractor") {
    //   const sourceDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.sourceDocumentLibraryName;
    //   const SourcedocumentItem = await this._Service.getDLItemById(sourceDocumentsDl, ID.value);
    //   this.setState({
    //     sourceDocumentItem: SourcedocumentItem.AcceptanceCodeId,
    //   });
    //   const selectItems = "ID,DocumentID,FileSizeDisplay,DocumentName,Revision,AcceptanceCode/ID,AcceptanceCode/Title,DocumentIndex/ID,CustomerDocumentNo,SubcontractorDocumentNo";
    //   const filterItems = "ID eq '" + ID.value + "' ";
    //   const expandItems = "AcceptanceCode,DocumentIndex";
    //   this._Service.getItemForSelectInDL(sourceDocumentsDl, selectItems, filterItems, expandItems)
    //     .then((sourceDocumentsItemsForGrid: any) => {
    //       console.log("sourceDocumentsItemsForGrid", sourceDocumentsItemsForGrid);
    //       this.setState({
    //         publishDocumentsItemsForGrid: sourceDocumentsItemsForGrid,
    //       });
    //     });
    // }
  }
  //transmittal type 
  private _onTransmitType(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
    console.dir(option);
    this.setState({
      transmittalTypekey: option.key,
      transmittalType: option.text,
      documentSelectedDiv: true,
    });
    if (option.text == 'Letter') {
      this._loadSourceDocumentsForLetter();
    }
    else if (option.text === 'Document') {
      this._loadPublishDocuments("");
    }
  }

}
