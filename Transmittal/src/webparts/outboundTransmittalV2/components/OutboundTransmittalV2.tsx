import * as React from 'react';
import styles from './OutboundTransmittalV2.module.scss';
import type { IOutboundTransmittalV2Props, IOutboundTransmittalV2State } from '../Interfaces/IOutboundTransmittalV2Props';
import { escape } from '@microsoft/sp-lodash-subset';
import { Checkbox, ChoiceGroup, DatePicker, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, FontWeights, IChoiceGroupOption, IDropdownOption, IDropdownStyles, IIconProps, IModalProps, ITextFieldStyles, ITooltipHostStyles, IconButton, Label, MessageBar, Modal, PrimaryButton, ProgressIndicator, Spinner, SpinnerSize, TextField, getTheme, mergeStyleSets } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import * as moment from 'moment';
import { OBService } from '../Services/OBService';
import SimpleReactValidator from 'simple-react-validator';
import { Web } from '@pnp/sp/webs';
import replaceString, { ReplacementFunction } from 'replace-string';
import * as _ from 'lodash';
import { add } from 'lodash';
import { MSGraphClientV3, HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { MultiSelect } from "react-multi-select-component";
import Select from 'react-select';
import { Accordion } from '@pnp/spfx-controls-react';
export default class OutboundTransmittalV2 extends React.Component<IOutboundTransmittalV2Props, IOutboundTransmittalV2State, {}> {
  private validator: SimpleReactValidator;
  private _Service: OBService;
  private reqWeb = Web(window.location.protocol + "//" + window.location.hostname + "/sites/" + this.props.hubSiteUrl);
  private emailsSelectedTo: any[] = [];
  private postUrl: any;
  private emailsSelectedCC: any = [];
  private contactCCDisplay: any = [];
  private contactToDisplay: any = [];
  private sortedArray: any = [];
  private transmittalID: any;
  private currentDate = new Date();
  private outboundRecallConfirmation: any;
  private status: any;
  private typeForDelete: any;
  private today = this.currentDate.toLocaleString();
  private keyForDelete: any;
  private additionalDivHide: any;
  private DocumentID: any;
  private forRecall: any;
  private postUrlForPermission: any;
  private myfileadditional: any;
  private postUrlForRecall: any;
  private permissionForRecall: any;
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
      vendorarray: [],
      selectedVendor: [],
      searchContactsTo: [],
      selectedContactsToName: [],
      searchContactsCC: [],
      selectedContactsCCName: [],
      divForToAndCC: "none",
      divForToAndCCSearch: "",
    };
    this._Service = new OBService(this.props.context);
  }

  public render(): React.ReactElement<IOutboundTransmittalV2Props> {
    const TransmitTo: IDropdownOption[] = [
      { key: '1', text: 'Customer' },
      { key: '2', text: 'Sub-Contractor' },
    ];
    const dialogContentProps = {
      type: DialogType.normal,
      title: (this.state.transmitTo === "Customer") ? "Select Customer Contacts" : "Select Sub-Contractors",
      closeButtonAriaLabel: 'Close',
    };
    const options: IChoiceGroupOption[] = [
      { key: 'Document', text: 'Document' },
      { key: 'Letter', text: 'Letter' },
    ];
    const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: "50%" } };
    const dragOptions: Partial<IModalProps> = { dragOptions: undefined, };
    const contactSearch: IIconProps = { iconName: 'ProfileSearch' };
    const AddIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
    const DeleteIcon: IIconProps = { iconName: 'Delete' };
    const CancelIcon: IIconProps = { iconName: 'Cancel' };
    const hostStyles: Partial<ITooltipHostStyles> = { root: { display: 'inline-block' } };
    const calloutProps = { gapSpace: 0 };
    const multiline: Partial<ITextFieldStyles> = { root: { height: "50px" } };
    const theme = getTheme();
    const contentStyles = mergeStyleSets({
      container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',


      },
      header: [
        // eslint-disable-next-line deprecation/deprecation
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
        // eslint-disable-next-line deprecation/deprecation
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
    return (
      <div>
        <div>
          <div style={{ display: this.state.loaderDisplay }}>
            <ProgressIndicator label="Loading......" />
          </div>
          <div className={styles.outboundTransmittalV2} style={{ display: this.state.webpartView }}>
            <Label className={styles.align}>{this.props.description}</Label>
            <div style={{ marginLeft: "522px" }}>
            </div>
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
                  <div style={{ display: this.state.hideSubContractor, width: "100%" }}>
                    <div style={{ display: "flex", marginLeft: "123px", marginTop: "3px" }}>
                      <Select
                        placeholder="Select Sub-Contractor"
                        isMulti={false}
                        options={this.state.subContractorItems}
                        onChange={this._drpdwnSubContractor.bind(this)}
                        isSearchable={true}
                        value={this.state.subContractorKey}
                        maxMenuHeight={150}
                        isClearable={false}
                      //styles={{: "10px", width: "248px", marginTop: "22px", display: this.state.subContractorDrpDwn }}
                      />

                      <Label style={{ marginLeft: "10px", display: this.state.subContractorLabel }}>{this.state.subContractor} </Label>
                    </div>
                    <div style={{ color: "#dc3545", marginLeft: "123px" }}>{this.validator.message("subContractor", this.state.subContractorKey, "required")}{" "}</div>
                  </div>
                </div>
                <div style={{ color: "#dc3545" }}>{this.validator.message("transmitTo", this.state.transmitToKey, "required")}{" "}</div>
                <hr />
                <div >
                  <div style={{ marginBottom: "10px" }} className={styles.borderForToCC}>
                    <span className={styles.span}></span>
                    <div style={{ width: "97%", display: this.state.divForToAndCCSearch }}>
                      <label style={{ fontWeight: "bold", }}>To</label>
                      <MultiSelect options={this.state.contactsForSearch} value={this.state.selectedContactsToName} onChange={this.setSelectedContactsTo} labelledBy="To" hasSelectAll={true} />
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
                        onChange={this._onDrpdwnCntact}
                        title="To"
                      />
                      <div style={{ color: "#dc3545" }}>{this.validator.message("selectedContactsTo", this.state.selectedContactsTo, "required")}{" "}</div>
                    </div>
                    <span className={styles.span}></span>
                    <div style={{ width: "97%", display: this.state.divForToAndCCSearch }}>
                      <label style={{ fontWeight: "bold", }}>CC</label>
                      <MultiSelect options={this.state.contactsForSearch} value={this.state.selectedContactsCCName} onChange={this.setSelectedContactsCC} labelledBy="CC" hasSelectAll={true} />
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
                        onChange={this._onDrpdwnCCContact}
                      />
                    </div>
                    <div style={{ width: "98%", fontWeight: "bold" }}>
                      <PeoplePicker
                        context={this.props.context as any}
                        titleText="Internal CC"
                        personSelectionLimit={20}
                        groupName={""} // Leave this blank in case you want to filter from all users
                        showtooltip={true}
                        required={true}
                        disabled={false}
                        ensureUser={true}
                        onChange={(items: any) => this._selectedInternalCCContacts(items)}
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
                      <ChoiceGroup options={options} onChange={this._onTransmitType} label="Select any" required={true} defaultSelectedKey={'Document'} disabled={true} />
                    </div>
                    <div style={{ marginLeft: "26px", marginRight: "15px", display: this.state.transmitTypeForLetter }}>
                      <ChoiceGroup options={options} onChange={this._onTransmitType} label="Select any" required={true} defaultSelectedKey={'Letter'} disabled={true} />
                    </div>
                    <div style={{ marginLeft: "26px", marginRight: "15px", display: this.state.transmitTypeForDefault }}>
                      <ChoiceGroup options={options} onChange={this._onTransmitType} label="Select any" required={true} />
                    </div>
                    <div style={{ marginLeft: "150px", marginTop: "9px" }}>
                      <Label>Check if cover letter needed</Label>
                      <div className={styles.mt1}><Checkbox label="Cover Letter" title="Check if cover letter needed or not." onChange={this._coverLetterNeeded}
                        checked={this.state.coverLetterNeeded} /></div>
                    </div>
                    <div style={{ marginLeft: "90px", marginTop: "9px" }}>
                      <Label>Select email type</Label>
                      <div className={styles.mt1}><Checkbox label="Send and Receive as shared folder" onChange={this._onSendAsSharedFolder}
                        checked={this.state.sendAsSharedFolder} /></div>
                      <div className={styles.mt1} style={{ display: this.state.sendAsMultipleEmailCheckBoxDiv }}><Checkbox label="Send as multiple emails" onChange={this._onSendAsMultipleFolder} checked={this.state.sendAsMultipleFolder} /></div>
                    </div>

                  </div>
                  {/* transmittal type validationdiv */}
                  <div style={{ color: "#dc3545", marginLeft: "26px" }}>{this.validator.message("transmittalType", this.state.transmittalType, "required")}{" "}</div>
                </div>
                <hr />
                {/* Notes */}
                <div style={{ marginLeft: "9px" }}>
                  <TextField label="Notes" multiline placeholder="" value={this.state.notes} onChange={this.notesOnchange} style={{ marginLeft: "20px", width: "290px" }} />
                </div>
                <hr />
                {/* filesizeDiv */}
                {this.state.itemsForGrid.length > 0 &&
                  <div hidden={this.state.fileSizeDiv} style={{ float: "right", color: (Number(this.state.fileSize) >= 25) ? "Red" : "Green" }}>Size : [{(this.state.fileSize < 1) ? this.state.fileSize + " MB" : this.state.fileSize + " MB"}]</div>
                }
                {/* project documents */}
                <div style={{ padding: "12px 0 12px 12px" }}>
                  <div style={{ display: "block" }}>
                    <div hidden={this.state.documentSelectedDiv} style={{ fontWeight: "bold", color: "Red" }}> {this.state.documentSelect}</div>
                    <Select
                      placeholder="Select Project Documents"
                      isMulti={false}
                      options={this.state.searchDocuments}
                      onChange={this._onDocumentClick.bind(this)}
                      isSearchable={true}
                      value={this.state.projectDocumentSelectKey}
                      maxMenuHeight={150}
                      isClearable={false}
                    />
                    <div style={{ color: "#dc3545", marginLeft: "123px" }}>
                      {this.validator.message("projectDocuments", this.state.projectDocumentSelectKey, "required")}{" "}
                    </div>
                  </div>
                </div>
                <div style={{ display: "flex" }}>
                  <div style={{ padding: "8px 0px 0 11px" }}>
                    <div style={{ width: "100%" }}>
                      <Dropdown id="t3"
                        selectedKey={this.state.transmitForKey}
                        placeholder="Select an option"
                        options={this.state.transmitForItems}
                        onChanged={this._drpdwnTransmitFor} style={{ width: "350px", marginRight: "8px" }} label="Transmit For" />
                      <div style={{ color: "#dc3545", marginLeft: "0px" }}>{this.validator.message("transmitForKey", this.state.transmitForKey, "required")}{" "}</div>
                    </div>
                    <div style={{ width: "100%" }}>
                      <DatePicker label="Due Date"
                        style={{ width: '350px', marginRight: "8px" }}
                        value={this.state.dueDate}
                        hidden={this.state.hideDueDate}
                        onSelectDate={this._dueDatePickerChange}
                        minDate={this.state.dueDateForBindingApprovalLifeCycle}
                        placeholder="Select a date..."
                        ariaLabel="Select a date"
                        formatDate={this._onFormatDate}
                      />
                    </div>
                  </div>
                  <div style={{ padding: "8px 0px 0px 11px", width: "80%" }}>
                    <TextField autoComplete="off" label="Comments" multiline placeholder="" value={this.state.comments} onChange={this.onCommentChange} style={{ height: "92px", }} />
                  </div>
                  <div hidden={this.state.hideGridAddButton}>
                    <i className={styles['icon-145']} aria-hidden="true"> <IconButton iconProps={AddIcon} title="Add" ariaLabel="Delete" onClick={this._showProjectDocumentGrid} style={{ padding: "43px 0px 0px 10px", display: this.state.hideButtonAfterSubmit }} /></i>
                  </div>
                </div>
                {/* projectDocumentGrid */}
                {this.state.itemsForGrid.length > 0 &&
                  <table className={styles.tableModal} hidden={this.state.showGrid} style={{ width: "100%" }}>
                    <tr style={{ background: "#f4f4f4" }}>
                      <th style={{ padding: "5px 10px" }} >Slno</th>
                      {/* <th style={{ padding: "5px 10px" }}>Doc Id</th> */}
                      <th style={{ padding: "5px 10px" }}>Document Name</th>
                      <th style={{ padding: "5px 10px" }}>Revision No</th>
                      <th style={{ padding: "5px 10px", display: (this.state.transmitTo === "Customer") ? "" : "none" }}>Customer Document No</th>
                      <th style={{ padding: "5px 10px", display: (this.state.transmitTo === "Sub-Contractor") ? "" : "none" }}>SubContractor Document No</th>
                      <th style={{ padding: "5px 10px", display: (this.state.transmitTo === "Sub-Contractor") ? "none" : "none" }}>Acceptance Code</th>
                      <th style={{ padding: "5px 10px" }}>Size (in MB)</th>
                      <th style={{ padding: "5px 10px" }}>Transmit For</th>
                      <th style={{ padding: "5px 10px" }}>Due Date</th>
                      <th style={{ padding: "5px 10px" }}>Comments</th>
                      <th style={{ padding: "5px 10px", display: this.state.hideButtonAfterSubmit }}>Delete</th>
                    </tr>
                    {this.state.itemsForGrid.map((items, key) => {
                      return (
                        <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                          <td style={{ padding: "5px 10px" }}>{key + 1}</td>
                          <td style={{ padding: "5px 10px" }}>{items.documentName} </td>
                          <td style={{ padding: "5px 10px" }}>{items.revision} </td>
                          <td style={{ padding: "5px 10px", display: (this.state.transmitTo === "Customer") ? "" : "none" }}>{items.customerDocumentNo} </td>
                          <td style={{ padding: "5px 10px", display: (this.state.transmitTo === "Sub-Contractor") ? "" : "none" }}>{items.subcontractorDocumentNo} </td>
                          <td style={{ padding: "5px 10px", display: (this.state.transmitTo === "Sub-Contractor") ? "none" : "none" }}>{items.acceptanceCodeTitle}</td>
                          <td style={{ padding: "5px 10px" }}>{items.fileSizeInMB}</td>
                          <td style={{ padding: "5px 10px" }}>{items.transmitFor} </td>
                          <td style={{ padding: "5px 10px" }}>{items.DueDate}</td>
                          <td style={{ padding: "5px 10px" }}>{items.comments}</td>
                          <td style={{ padding: "5px 10px", display: this.state.hideButtonAfterSubmit }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key, "ProjectDocuments")} /></td>
                        </tr>
                      );
                    })}
                  </table>
                }
                <hr style={{ marginTop: "20px" }} />
                <Accordion title={''}  >
                  <Accordion title="External Documents" >
                    <div style={{ width: "100%" }}>
                      <div style={{ width: "50%" }}>
                        <input type="file" name="myFile" id="newfile" style={{ marginRight: "-13px", marginLeft: "12px" }} onChange={(e) => this._uploadadditional(e)} ref={ref => this.myfileadditional = ref} ></input>
                      </div>
                      <div style={{ color: "#dc3545", display: this.state.uploadDocumentError, marginLeft: "9px" }}>Sorry this document is unable to process due to unwanted characters.Please rename the document and try again.</div>
                      <div style={{ width: "100%", display: "flex" }}>
                        <div style={{ width: "100%", padding: "10px 7px 10px 9px" }}> <TextField required={true} value={this.state.externalComments} placeholder="" onChange={this.onCommentExternalChange} />
                        </div>
                        <div style={{ width: "5%", padding: "10px 7px 10px 9px", display: this.additionalDivHide }}>
                          <IconButton iconProps={AddIcon} title="Add External Documents" ariaLabel="Add" onClick={this._showExternalGrid} style={{ marginTop: "-4px", display: this.state.hideButtonAfterSubmit }} />
                        </div>
                      </div>
                      <div style={{ color: "#dc3545", marginLeft: "123px" }}>{this.validator.message("externalcomments", this.state.externalComments, "required")}{" "}</div>
                    </div>
                  </Accordion>
                </Accordion>
                <div hidden={this.state.showExternalGrid} >
                  <table className={styles.tableModal}  >
                    <tr style={{ background: "#f4f4f4" }}>
                      <th style={{ padding: "5px 10px" }}>Slno</th>
                      <th style={{ padding: "5px 10px" }}>Document Name</th>
                      <th style={{ padding: "5px 10px" }}>Size (in MB)</th>
                      <th style={{ padding: "5px 10px" }}>Comments</th>
                      <th style={{ padding: "5px 10px", display: this.state.hideButtonAfterSubmit }}>Delete</th>
                    </tr>
                    {this.state.itemsForExternalGrid.map((items, key) => {
                      return (
                        <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                          <td style={{ padding: "5px 10px" }}>{key + 1}</td>
                          <td style={{ padding: "5px 10px" }}>{items.documentName}</td>
                          <td style={{ padding: "5px 10px" }}>{items.fileSizeInMB}</td>
                          <td style={{ padding: "5px 10px" }}>{items.externalComments}</td>
                          <td style={{ padding: "5px 10px", display: this.state.hideButtonAfterSubmit }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key, "AdditionalDocuments")} /></td>
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
                <PrimaryButton text="Save as draft" style={{ marginLeft: "auto", marginRight: "11px", display: this.state.hideButtonAfterSubmit }} onClick={() => this._onSaveAsDraftBtnClick()} />
                <PrimaryButton text="Preview" style={{ marginRight: "11px", marginLeft: "auto" }} onClick={this._onPreviewBtnClick} />
                <PrimaryButton text="Confirm & Send" style={{ marginRight: "11px", marginLeft: "auto", display: this.state.hideButtonAfterSubmit }} onClick={this._confirmAndSendBtnClick} />
                <PrimaryButton text="Recall" style={{ marginRight: "11px", marginLeft: "auto", display: this.state.hideUnlockButton }} onClick={this._recallTransmittalConfirmation} />
                <PrimaryButton text="Cancel" style={{ marginLeft: "auto" }} onClick={this._hideGrid} />
              </div>
            </div>
            <div>
            </div>
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
                                <div className={styles.divTableCell} style={{ display: this.state.transmitTo === "Sub-Contractor" ? "none" : "none" }}>&nbsp;{items.acceptanceCodeTitle}</div>
                                <div className={styles.divTableCell}>&nbsp;{items.fileSizeInMB}</div>
                                <div className={styles.divTableCell}>&nbsp;{items.transmitFor}</div>
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
    );
  }

  //tranmit to dropdown
  public _drpdwnTransmitTo(event: React.FormEvent<HTMLDivElement>, option: IDropdownOption): void {
    let customerArray: { key: any; text: string; }[] = [];
    let customerArraySearch: { value: any; label: string; }[] = [];
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
      this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.contactListName)
        .items
        .filter("CustomerOrVendorID eq '" + this.state.customerId + "'  and  LegalEntityId eq '" + this.state.legalId + "'")
        .getAll()
        .then(contacts => {
          for (var k in contacts) {
            if (contacts[k].Active === true) {
              let transmitForItemdata = {
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
          this.setState({
            contacts: customerArray,
            contactsForSearch: customerArraySearch
          });
        });
      this._loadPublishDocuments();
      return this.setState({
        hideCustomer: "",
        subContractorItems: [],
        subContractorKey: "",
        hideSubContractor: "none",
      });

    }
    else if (option.text === "Sub-Contractor") {
      this.setState({
        documentSelect: "",
        documentSelectedDiv: true,
      });
      let subcontractorArray: { value: any; label: any; }[] = [];
      this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/SubContractorMaster")
        .items
        .filter("ProjectId eq '" + this.state.projectNumber + "' and  Title eq '" + this.state.legalId + "' ")
        .getAll()
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
          } this._loadSourceDocuments();
          return this.setState({
            contacts: [],
            hideSubContractor: "",
            hideCustomer: "none",
            subContractorItems: subcontractorArray
          });
        });
    }
  }
  public _drpdwnSubContractor(option: { text: any; value: string; label: any; }) {
    let subContractor: { key: any; text: string; }[] = [];
    let subContractorArray = [];
    let subContractorArraySearch: { value: any; label: string; }[] = [];
    console.log(option.text);
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
    this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.contactListName)
      .items
      .filter("CustomerOrVendorID eq '" + option.value + "' and  LegalEntityId eq '" + this.state.legalId + "' ")
      .getAll()
      .then(contacts => {
        for (var k in contacts) {
          if (contacts[k].Active === true) {
            let transmitForItemdata = {
              key: contacts[k].Email,
              text: contacts[k].Title + " " + (contacts[k].LastName !== null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
            };
            let transmitForItemdataSearch = {
              value: contacts[k].Email,
              label: contacts[k].Title + " " + (contacts[k].LastName !== null ? contacts[k].LastName : " ") + "<" + contacts[k].Email + ">",
            };
            subContractorArray.push(transmitForItemdata);
            subContractorArraySearch.push(transmitForItemdataSearch);
            subContractor.push(transmitForItemdata);
          }
        }
        this.setState({
          contacts: subContractor,
          contactsForSearch: subContractorArraySearch
        });
      });

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
    for (let i = 0; i < option.length; i++) {
      selectedContactsIdArray.push(option[i].value);
      selectedContactsNameArray.push(option[i].label.split("<")[0]);
    }
    this.emailsSelectedTo.push(selectedContactsIdArray);
    this.contactToDisplay.push(selectedContactsNameArray);
    checkedContacts = (this.emailsSelectedTo).toString();
    checkedContactsDisplay = (this.contactToDisplay).toString();
    let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
    this.setState({
      selectedContactsTo: checkedContactsSemicolonAttached,
      selectedContactsToDisplayName: checkedContactsDisplay,
      selectedVendor: option,
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
    let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
    this.setState({
      selectedContactsCC: checkedContactsSemicolonAttached,
      selectedContactsCCDisplayName: checkedContactsDisplay,
      selectedContactsCCName: option,
      searchContactsCC: selectedContactsCCArray
    });
  }
  // from current site
  //page load from project information list
  public projectInformation = async () => {
    await this._Service.getListItems(this.props.siteUrl, this.props.projectInformationListName)
      .then(projectInformation => {
        if (projectInformation.length > 0) {
          for (var k in projectInformation) {
            if (projectInformation[k].Key === "ProjectName") {
              this.setState({
                projectName: projectInformation[k].Title,
              });
            }
            if (projectInformation[k].Key === "Customer") {
              this.setState({
                customerName: projectInformation[k].Title,
              });
            }
            if (projectInformation[k].Key === "ContractNumber") {
              this.setState({
                contractNumber: projectInformation[k].Title,
              });
            }
            if (projectInformation[k].Key === "ApprovalCycle") {
              this.setState({
                approvalLifeCycle: projectInformation[k].Title,
              });
              const dueDate = new Date();
              let days = projectInformation[k].Title;
              console.log(Number(days));
              dueDate.setDate(dueDate.getDate() + Number(days));
              this.setState({
                hideDueDate: false,
                dueDate: dueDate,
                dueDateForBindingApprovalLifeCycle: dueDate,
              });
            }
            if (projectInformation[k].Key === "ProjectNumber") {
              this.setState({
                projectNumber: projectInformation[k].Title,
              });
            }
            if (projectInformation[k].Key === "CustomerID") {
              this.setState({
                customerId: projectInformation[k].Title,
              });
            }
            if (projectInformation[k].Key === "LegalEntityId") {
              this.setState({
                legalId: projectInformation[k].Title,
              });
            }
          }
        }
      });
  }
  //Current User
  private async _currentUser() {
    this._Service.getCurrentUserId().then(currentUser => {
      this.setState({
        currentUser: currentUser.Id,
      });
    });
  }
  //for To fields
  private _onDrpdwnCntact = async (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    let checkedContacts: string;
    let checkedContactsDisplay: string;
    let contacts: any;
    if (option.selected) {
      contacts = {
        key: option.key,
        text: option.text.split("<"),//splitting the < to split mail id to inserting ToName and CCName
      };
      this.emailsSelectedTo.push(contacts.key);
      this.contactToDisplay.push(contacts.text[0]);
      checkedContacts = (this.emailsSelectedTo).toString();
      checkedContactsDisplay = (this.contactToDisplay).toString();
      let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
      this.setState({
        selectedContactsTo: checkedContactsSemicolonAttached,
        selectedContactsToDisplayName: checkedContactsDisplay,
      });
      console.log("checkedContacts", checkedContacts);
    }
    else {
      this.emailsSelectedTo.splice(index, 1);
      this.contactToDisplay.splice(index, 1);
      let newarray = this.emailsSelectedTo.filter(element => element !== option.key);
      checkedContacts = (newarray).toString();
      let splittedName = (option.text).split("<");
      console.log(splittedName[0]);
      let newarrayContactDisplay = this.contactToDisplay.filter((element: string) => element !== splittedName[0]);
      checkedContactsDisplay = (newarrayContactDisplay).toString();
      this.contactToDisplay = newarrayContactDisplay;
      console.log("afterFilter", this.contactToDisplay);
      let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
      this.setState({
        selectedContactsTo: checkedContactsSemicolonAttached,
        selectedContactsToDisplayName: checkedContactsDisplay,
      });
      console.log("checkedContacts", checkedContacts);
    }
  }
  //for CC fields
  private _onDrpdwnCCContact = async (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number) => {
    let checkedContacts: string;
    let checkedContactsDisplay: string;
    if (option.selected) {
      let contacts = {
        key: option.key,
        text: option.text.split("<"),
      };
      this.emailsSelectedCC.push(contacts.key);
      this.contactCCDisplay.push(contacts.text[0]);
      checkedContacts = (this.emailsSelectedCC).toString();
      checkedContactsDisplay = (this.contactCCDisplay).toString();
      let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
      this.setState({
        selectedContactsCC: checkedContactsSemicolonAttached,
        selectedContactsCCDisplayName: checkedContactsDisplay,
      });
      console.log("checkedContactsCC", checkedContacts);
    }
    else {

      this.emailsSelectedCC.splice(index, 1);
      console.log("beforeFilter", this.contactCCDisplay);
      this.contactCCDisplay.splice(index, 1);
      let newarray = this.emailsSelectedCC.filter((element: string | number) => element !== option.key);
      checkedContacts = (newarray).toString();
      let splittedName = (option.text).split("<");
      console.log(splittedName[0]);
      let newarrayContactDisplay = this.contactCCDisplay.filter((element: string) => element !== splittedName[0]);
      checkedContactsDisplay = (newarrayContactDisplay).toString();
      this.contactCCDisplay = newarrayContactDisplay;
      console.log("afterFilter", this.contactCCDisplay);
      let checkedContactsSemicolonAttached = replaceString(checkedContacts, ',', ';');
      this.setState({
        selectedContactsCC: checkedContactsSemicolonAttached,
        selectedContactsCCDisplayName: checkedContactsDisplay,
      });
      console.log("checkedContactsCC", checkedContacts);
    }
  }
  public _selectedInternalCCContacts = (items: any[]) => {
    let getSelectedInternalID = [];
    let getSelectedInternalDisplayName = [];
    let getSelectedInternalEmailID = [];
    for (let item in items) {
      getSelectedInternalID.push(items[item].id);
      getSelectedInternalDisplayName.push(items[item].text);
      getSelectedInternalEmailID.push(items[item].secondaryText);
    }
    var displayInternalName = getSelectedInternalDisplayName.toString();
    var InternalEmailID = getSelectedInternalEmailID.toString();
    let InternalEmailIDSemicolonAttached = replaceString(InternalEmailID, ',', ';');
    this.setState({ internalCCContacts: getSelectedInternalID, internalCCContactsDisplayNameForPreview: displayInternalName, internalContactsEmail: InternalEmailIDSemicolonAttached });
  }
  //transmittal type 
  private _onTransmitType(ev: React.FormEvent<HTMLInputElement>, option: IChoiceGroupOption): void {
    console.dir(option);
    this.setState({
      transmittalTypekey: option.key,
      transmittalType: option.text,
      documentSelectedDiv: true,
    });
    if (option.text === 'Letter') {
      this._loadSourceDocumentsForLetter();
    }
    else if (option.text === 'Document') {
      if (this.state.transmitTo === "Sub-Contractor") {
        this._loadSourceDocuments();
      }
      else {
        this._loadPublishDocuments();
      }
    }
  }
  public async _loadSourceDocumentsForLetter() {
    let temDoc: [];
    let publishedDocumentArray: { value: any; label: any; }[] = [];
    let transmitForItemdata;
    const publishedDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.publishDocumentLibraryName;
    this._Service.getLibraryItems(publishedDocumentsDl)
      .then(async publishDocumentsItems => {
        console.log("PublishDocumentForCustomerCount", publishDocumentsItems.length);
        this.sortedArray = _.orderBy(publishDocumentsItems, 'FileLeafRef', ['asc']);
        if (publishDocumentsItems.length > 0) {
          this._Service.getDIItems(this.props.context.pageContext.web.serverRelativeUrl, "DocumentIndex")
            .then(DIndexItems => {
              console.log("PublishDocumentFormIndex", DIndexItems.length);
              const filteredIndexItems = this.sortedArray.filter((item: { DocumentIndexId: any; }) =>
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
          //alert("No documents all transmittal status is ONGOING");
          console.log("No documents for transmittal");
          this.setState({
            documentSelectedDiv: false,
            documentSelect: "No documents for transmittal "
          });
        }
      }).catch((err) => {
        console.log("Error = ", err);
      });


  }
  public _coverLetterNeeded = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ coverLetterNeeded: true, });
    }
    else if (!isChecked) { this.setState({ coverLetterNeeded: false, }); }
  }
  //for customers documents from published docs
  public async _loadPublishDocuments() {
    let publishedDocumentArray: { value: any; label: any; }[] = [];
    let transmitForItemdata;
    const publishedDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.publishDocumentLibraryName;
    this._Service.getLibraryItems(publishedDocumentsDl)
      .then(async publishDocumentsItems => {
        console.log("PublishDocumentForCustomerCount", publishDocumentsItems.length);
        this.sortedArray = _.orderBy(publishDocumentsItems, 'FileLeafRef', ['asc']);
        if (publishDocumentsItems.length > 0) {
          this._Service.getDIItems(this.props.context.pageContext.web.serverRelativeUrl, "DocumentIndex")
            .then(DIndexItems => {
              console.log("PublishDocumentForCustomerFromIndex", DIndexItems.length);
              const filteredIndexItems = this.sortedArray.filter((item: { DocumentIndexId: any; }) =>
                DIndexItems.some((pdItem: { ID: any; }) => pdItem.ID === item.DocumentIndexId)
              );
              if (filteredIndexItems.length > 0) {
                filteredIndexItems.forEach((filteredItems: any) => {
                  transmitForItemdata = {
                    value: filteredItems.ID,
                    label: filteredItems.DocumentName
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
          //alert("No documents all transmittal status is ONGOING");
          console.log("No documents for transmittal");
          this.setState({
            documentSelectedDiv: false,
            documentSelect: "No documents for transmittal "
          });
        }
      }).catch((err) => {
        console.log("Error = ", err);
      });
  }
  public async _loadSourceDocuments() {
    //for customer values from sourceDocuments
    let sourceDocumentArray: { value: any; label: any; }[] = [];
    const sourceDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.sourceDocumentLibraryName;
    this._Service.getSourceLibraryItems(sourceDocumentsDl)
      .then(sourceDocumentArrayItems => {
        console.log("SourceDocumentForCustomer", sourceDocumentArrayItems.length);
        if (sourceDocumentArrayItems.length > 0) {
          this.sortedArray = _.orderBy(sourceDocumentArrayItems, 'FileLeafRef', ['asc']);
          this.sortedArray.forEach((sourceItems: { ID: any; DocumentName: any; }) => {
            let transmitForItemdata = {
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
      }).catch((err) => {
        console.log("Error = ", err);
        this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: false, message: err, messageType: 1 }, });
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
  public _onSendAsMultipleFolder = async (ev: React.FormEvent<HTMLInputElement>, isChecked?: boolean) => {
    if (isChecked) {
      this.setState({ sendAsMultipleFolder: true, });
    }
    else if (!isChecked) { this.setState({ sendAsMultipleFolder: false, }); }
  }
  private notesOnchange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    const newMultiline = newText.length > 50;
    if (newMultiline !== this.state.toggleMultiline) {
      this.setState({
        toggleMultiline: true,
      });
    }
    this.setState({ notes: newText || '' });
  }
  private async _onDocumentClick(ID: { value: any }) {
    this.setState({ projectDocumentSelectKey: ID.value, documentSelectedDiv: true, hideGridAddButton: false, });
    this.setState({
      searchDiv: "none",
    });
    if (this.state.transmitTo === "Customer") {
      const publishDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.publishDocumentLibraryName;
      const selectItems = "ID,DocumentID,FileSizeDisplay,DocumentName,Revision,DocumentIndex/ID,CustomerDocumentNo,SubcontractorDocumentNo";
      const filterItems = "ID eq '" + ID.value + "' ";
      const expandItems = "DocumentIndex";
      this._Service.getItemForSelectInDL(publishDl, selectItems, filterItems, expandItems)
        .then((publishDocumentsItemsForGrid: any) => {
          console.log("publishDocumentsItemsForGrid", publishDocumentsItemsForGrid);
          this.setState({
            publishDocumentsItemsForGrid: publishDocumentsItemsForGrid
          });
        });
    }
    else if (this.state.transmitTo === "Sub-Contractor") {
      const sourceDocumentsDl: string = this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.sourceDocumentLibraryName;
      const SourcedocumentItem = await this._Service.getDLItemById(sourceDocumentsDl, ID.value);
      this.setState({
        sourceDocumentItem: SourcedocumentItem.AcceptanceCodeId,
      });
      const selectItems = "ID,DocumentID,FileSizeDisplay,DocumentName,Revision,AcceptanceCode/ID,AcceptanceCode/Title,DocumentIndex/ID,CustomerDocumentNo,SubcontractorDocumentNo";
      const filterItems = "ID eq '" + ID.value + "' ";
      const expandItems = "AcceptanceCode,DocumentIndex";
      this._Service.getItemForSelectInDL(sourceDocumentsDl, selectItems, filterItems, expandItems)
        .then(sourceDocumentsItemsForGrid => {
          console.log("sourceDocumentsItemsForGrid", sourceDocumentsItemsForGrid);
          this.setState({
            publishDocumentsItemsForGrid: sourceDocumentsItemsForGrid,
          });
        });
    }
  }
  public _drpdwnTransmitFor(option: { key: any; text: any }) {
    this.setState({ transmitForKey: option.key, transmitFor: option.text });
    const select = "ApprovalRequired,AcceptanceCode";
    const filter = "Title eq '" + option.text + "'";
    this._Service.getItemForSelectInLists(this.props.siteUrl, this.props.transmittalCodeSettingsListName, select, filter)
      .then(transmitfor => {
        this.setState({
          approvalRequired: transmitfor[0].ApprovalRequired,
        });
      });
  }
  private _dueDatePickerChange = (date?: Date): void => {
    this.setState({ dueDate: date, });
  }
  private onCommentChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    const newMultiline = newText.length > 50;
    if (newMultiline !== this.state.toggleMultiline) {
      this.setState({
        toggleMultiline: true,
      });
    }
    this.setState({ comments: newText || '' });
  }
  //settings for date format
  private _onFormatDate = (date: Date): string => {
    const dat = date;
    console.log(moment(date).format("DD/MM/YYYY"));
    let selectd = moment(date).format("DD/MM/YYYY");
    return selectd;
  }
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
          let sizeOfDocument;
          if (this.state.transmitTo === "Customer") {
            sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
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
            sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
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
        let sizeOfDocument;
        if (this.state.transmitTo === "Customer") {
          sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
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

            let totalSize = add(totalAdditional, totalsizeProjects);
            let convertKBtoMB = Number(totalSize).toFixed(2);
            this.setState({
              fileSize: Number(convertKBtoMB)
            });
            console.log(this.state.fileSize);
          }
        }
        else if (this.state.transmitTo === "Sub-Contractor") {
          sizeOfDocument = (((this.state.publishDocumentsItemsForGrid[0].FileSizeDisplay / 1024)).toFixed(3));
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

            let totalSize = add(totalAdditional, totalsizeProjects);
            let convertKBtoMB = Number(totalSize).toFixed(2);
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
  private modalProps = {
    isBlocking: true,
  };
  // private async _recallSubmit() {
  //   this.setState({
  //     recallConfirmMsg: true,
  //     recallConfirmMsgDiv: "none",
  //     statusMessage: { isShowMessage: true, message: "Recalled" + this.state.transmittalNo, messageType: 4 },
  //   });
  //   await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalHeaderListName)
  //   .items.getById(this.transmittalID).update({
  //     TransmittalStatus: "Recalled",
  //     RecalledDate: this.currentDate,
  //     RecalledById: this.state.currentUser,
  //   }).then((outboundDetailsList: any) => {
  //     for (let i = 0; i < this.state.currentOutboundDetailItem.length; i++) {
  //       sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName)
  //       .items.getById(this.state.currentOutboundDetailItem[i].ID).update({
  //         TransmittalStatus: "Draft",
  //       });
  //       sp.web.getList(this.props.siteUrl + "/Lists/DocumentIndex").items.getById(this.state.currentOutboundDetailItem[i].DocumentIndex.ID).update({
  //         TransmittalStatus: "Draft",
  //       });
  //       sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibraryName).items.select("ID,DocumentID").filter("DocumentIndexId eq '" + this.state.currentOutboundDetailItem[i].DocumentIndex.ID + "'").get().then((sourceDocumentID: { DocumentID: any; }[]) => {
  //         // alert("SourceDocumentID"+sourceDocumentID[0].ID);
  //         sp.web.getList(this.props.siteUrl + "/" + this.props.sourceDocumentLibraryName).items.getById(sourceDocumentID[0].ID).update({
  //           TransmittalStatus: "Draft",
  //         });
  //         sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.transmittalHistoryLogList).items.add({
  //           Title: sourceDocumentID[0].DocumentID,
  //           Status: "Recalled",
  //           DocumentIndexId: this.state.currentOutboundDetailItem[i].DocumentIndex.ID,
  //           LogDate: this.currentDate,
  //         });
  //         sp.web.getList(this.props.siteUrl + "/Lists/DocumentIndex").items.getById(this.state.currentOutboundDetailItem[i].DocumentIndex.ID).select("Owner/ID,Owner/Title,Owner/EMail,DocumentName,DocumentController/ID,DocumentController/Title,DocumentController/EMail").expand("Owner,DocumentController").get().then(async (forGettingOwner: { Owner: { EMail: any; Title: string | ReplacementFunction; }; DocumentName: string | ReplacementFunction; DocumentController: { EMail: any; Title: string | ReplacementFunction; }; }) => {
  //           //mail for owner
  //           this._sendAnEmailUsingMSGraph(forGettingOwner.Owner.EMail, "Recall", forGettingOwner.Owner.Title, forGettingOwner.DocumentName);
  //           this._sendAnEmailUsingMSGraph(forGettingOwner.DocumentController.EMail, "Recall", forGettingOwner.DocumentController.Title, forGettingOwner.DocumentName);
  //         });
  //       });
  //     }
  //   }).then((mailSend: any) => {
  //     this._LAUrlGettingForRecall();
  //     this.setState({
  //       hideButtonAfterSubmit: "",
  //       hideUnlockButton: "none",
  //       normalMsgBar: "",
  //       statusMessage: { isShowMessage: true, message: "Recalled" + this.state.transmittalNo, messageType: 4 },
  //     });
  //     setTimeout(() => {
  //       window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
  //     }, 20000);
  //   });
  // }
  // sending Email for owners
  private async _sendAnEmailUsingMSGraph(email: any, type: string, name: string | ReplacementFunction, documentName: string | ReplacementFunction): Promise<void> {
    let Subject;
    let Body;
    const emailNoficationSettings: any[] = await this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.emailNotificationSettings)
      .items.filter("Title eq '" + type + "'")();
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    //Replacing the email body with current values
    let replacedSubject1 = replaceString(Subject, '[DocumentName]', documentName);
    let replacedSubject = replaceString(replacedSubject1, '[TransmittalNo]', this.state.transmittalNo);
    let replaceRequester = replaceString(Body, '[Sir/Madam],', name);
    let replaceBody = replaceString(replaceRequester, '[DocumentName]', documentName);
    let replacelink = replaceString(replaceBody, '[TransmittalNo]', this.state.transmittalNo);
    let var1: any[] = replacelink.split('/');
    let FinalBody = replacelink;
    if (email) {
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject,
          "body": {
            "contentType": "HTML",
            "content": FinalBody
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": email
              }
            }
          ],
        }
      };
      //Send Email uisng MS Graph  
      this.props.context.msGraphClientFactory
        .getClient("3")
        .then((client: MSGraphClientV3): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody, (error: any, response: any, rawResponse?: any) => {
            });
        });
    }
    // }
  }
  private dialogContentRecallProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: "Do you want to Recall ?",
  };
  private _confirmDeleteItem = async (docID: any, items: any, key: any | number) => {
    // if (this.transmittalID === "" || this.transmittalID === null) {
    //   this.setState({
    //     confirmDeleteDialog: true,
    //     deleteConfirmation: "none"
    //   });
    //   this.validator.hideMessages();
    //   if (this.typeForDelete === "ProjectDocuments") {
    //     this.itemDeleteFromGrid(items, key);
    //   }
    //   else if (this.typeForDelete === "AdditionalDocuments") {
    //     this.itemDeleteFromExternalGrid(items, key);
    //   }

    // }
    // else {
    //   this.setState({
    //     confirmDeleteDialog: true,
    //     deleteConfirmation: "none"
    //   });
    //   this.validator.hideMessages();
    //   console.log(items[key]);

    //   if (this.typeForDelete == "ProjectDocuments") {
    //     // alert(docID);

    //     if (docID) {
    //       let list = sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName);
    //       await list.items.getById(parseInt(docID)).delete();
    //       let selectHeaderItems = "Id,DocumentIndex/ID,DocumentIndex/Title,DueDate,SentComments,Revision,Title,Size,TransmittedFor/ID,TransmittedFor/Title,Temporary,TransmittalHeader/ID,DocumentLibraryID,ID,ApprovalRequired";
    //       sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName).items.select(selectHeaderItems).expand("DocumentIndex,TransmittedFor,TransmittalHeader").filter("TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ").get().then((outboundTransmittalDetailsListName: any) => {
    //         console.log("outboundTransmittalDetailsListName", outboundTransmittalDetailsListName);
    //         this.setState({
    //           currentOutboundDetailItem: outboundTransmittalDetailsListName,
    //         });
    //       });
    //       this.setState({
    //         itemsForGrid: this.state.itemsForGrid,
    //       });
    //     }
    //     this.itemDeleteFromGrid(items, key);
    //   }
    //   else if (this.typeForDelete == "AdditionalDocuments") {
    //     if (docID) {
    //       let list = sp.web.getList(this.props.siteUrl + "/" + this.props.outboundAdditionalDocumentsListName + "/");
    //       await list.items.getById(parseInt(docID)).delete();
    //       sp.web.getList(this.props.siteUrl + "/" + this.props.outboundAdditionalDocumentsListName + "/").items.filter("TransmittalIDId eq '" + this.transmittalID + "' ").get().then((listItems: any) => {
    //         this.setState({
    //           currentOutboundAdditionalItem: listItems,
    //         });
    //       });
    //       this.setState({
    //         itemsForExternalGrid: this.state.itemsForExternalGrid,
    //       });
    //     }
    //     this.itemDeleteFromExternalGrid(items, key);
    //   }
    // }
  }
  //deleting
  public itemDeleteFromGrid(items: { fileSize: any; }, key: number) {
    console.log(items);
    this.state.itemsForGrid.splice(key, 1);
    console.log("after removal", this.state.itemsForGrid);
    console.log(items.fileSize);
    this.setState({
      itemsForGrid: this.state.itemsForGrid,
      documentSelectedDiv: true,
      projectDocumentSelectKey: "",

    });
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
  public itemDeleteFromExternalGrid(items: { fileSize: any; }, key: number) {
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
  private _closeModal = (): void => {
    this.setState({ showReviewModal: false });
  }
  private _hideGrid() {
    this.setState({
      confirmCancelDialog: false,
      cancelConfirmMsg: "",
    });
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
        this.typeForDelete = "ProjectDocuments";
        this.keyForDelete = key;
      } else if (type === "AdditionalDocuments") {
        this.typeForDelete = "AdditionalDocuments";
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
        this.typeForDelete = "ProjectDocuments";
        this.keyForDelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.outboundDetailsID,
        });
      } else if (type === "AdditionalDocuments") {
        // alert("additionalid" + items.additionalDocumentID);
        this.typeForDelete = "AdditionalDocuments";
        this.keyForDelete = key;
        this.setState({
          tempDocIndexIDForDelete: items.additionalDocumentID,
        });
      }
    }

  }
  //Save as draft 
  public _onSaveAsDraftBtnClick() {
    // alert("success");
    let selectedContactsTo = (this.state.selectedContactsTo === null) ? "" : this.state.selectedContactsTo.toString();
    let selectedContactsCC = (this.state.selectedContactsCC === null) ? "" : this.state.selectedContactsCC.toString();
    let sourceDocumentId;
    //total files
    let totalFiles: number;
    let convertKBtoMB: string;
    totalFiles = add(this.state.itemsForGrid.length, this.state.itemsForExternalGrid.length);
    //sizecalculating
    let totalsizeProjects = 0;
    let totalAdditional = 0;
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
    if (this.transmittalID === null || this.transmittalID === "") {
      console.log("Save as draft button clicked");
      //get value
      console.log(this.state.selectedContactsTo);
      console.log(selectedContactsTo);
      if (this.state.transmitTo !== "" && this.state.transmittalTypekey !== "") {
        this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "Saved Successfully", messageType: 4 }, });
        this._trannsmittalIDGeneration().then(afterIdgeneration => {
          //outboundHeaderListInsertion
          const addOTH = {
            Title: this.state.transmittalNo,
            TransmittalCategory: this.state.transmitTo,
            Customer: this.state.customerName,
            CustomerID: (this.state.transmitTo === "Customer") ? this.state.customerId : "",
            SubContractor: this.state.subContractor,
            SubContractorID: (this.state.subContractorKey).toString(),
            ToEmails: selectedContactsTo,
            CCEmails: selectedContactsCC,
            Notes: this.state.notes,
            TransmittalStatus: "Draft",
            TransmittalType: this.state.transmittalType,
            TransmittedById: this.state.currentUser,
            SendAsSharedFolder: this.state.sendAsSharedFolder,
            ReceiveInSharedFolder: this.state.recieveInSharedFolder,
            SendAsMultipleEmails: this.state.sendAsMultipleFolder,
            TransmittalSize: (this.state.itemsForGrid.length === 0) ? "" : (convertKBtoMB).toString(),
            TotalFiles: (totalFiles).toString(),
            ToName: this.state.selectedContactsToDisplayName,
            CCName: this.state.selectedContactsCCDisplayName,
            CoverLetter: this.state.coverLetterNeeded,
            InternalCCId: { results: this.state.internalCCContacts, }
          }
          this._Service.addToList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, addOTH)
            .then(async (outboundTransmittalHeader: { data: { ID: string; }; }) => {
              this.setState({ outboundTransmittalHeaderId: outboundTransmittalHeader.data.ID });
              const UpdateLink = {
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

                }
              }
              await this._Service.updateList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, UpdateLink, Number(outboundTransmittalHeader.data.ID))
                .then(async (outboundTransmittalDetails: any) => {
                  if (this.state.itemsForGrid.length > 0) {
                    for (var i in this.state.itemsForGrid) {
                      const itemsToAdd = {
                        Title: this.state.itemsForGrid[i].documentName,
                        TransmittalHeaderId: outboundTransmittalHeader.data.ID,
                        DocumentIndexId: this.state.itemsForGrid[i].documentIndexId,
                        Revision: this.state.itemsForGrid[i].revision,
                        TransmittalRevision: this.state.itemsForGrid[i].revision,
                        DueDate: this.state.itemsForGrid[i].dueDate,
                        Size: this.state.itemsForGrid[i].fileSizeInMB,
                        SentComments: this.state.itemsForGrid[i].comments,
                        CustomerAcceptanceCodeId: (this.state.sourceDocumentItem !== null) ? this.state.itemsForGrid[i].acceptanceCode : null,
                        TransmittedForId: this.state.itemsForGrid[i].transmitForKey,
                        TransmitFor: this.state.itemsForGrid[i].transmitFor,
                        ApprovalRequired: this.state.itemsForGrid[i].approvalRequired,
                        TransmittalStatus: "Draft",
                        DocumentLibraryID: this.state.itemsForGrid[i].publishDoumentlibraryID,
                        Slno: (Number(i) + Number(1)).toString(),
                        CustomerDocumentNo: this.state.itemsForGrid[i].customerDocumentNo,
                        SubcontractorDocumentNo: this.state.itemsForGrid[i].subcontractorDocumentNo
                      }
                      await this._Service.addToList(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, itemsToAdd);
                    }
                  }

                  if (this.state.itemsForExternalGrid.length > 0) {
                    for (var k in this.state.itemsForExternalGrid) {
                      var splitted = this.state.itemsForExternalGrid[k].documentName.split(".");
                      console.log(splitted.length);
                      console.log(splitted[splitted.length - 1]);
                      let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + this.state.transmittalNo + '.' + splitted[splitted.length - 1];
                      let docName = documentNameExtension;
                      const additionalMetadataUpdate = {
                        Title: this.state.itemsForExternalGrid[k].documentName,
                        TransmittalIDId: outboundTransmittalHeader.data.ID,
                        Size: this.state.itemsForExternalGrid[k].fileSizeInMB,
                        Comments: this.state.itemsForExternalGrid[k].externalComments,
                        TransmittalStatus: "Draft",
                        Slno: (Number(k) + Number(1)).toString(),
                      }
                      this._Service.uploadDocument(docName, this.state.itemsForExternalGrid[k].content, this.props.outboundAdditionalDocumentsListName, additionalMetadataUpdate);

                    }
                  }
                });
            });
        }).then(transmittalLogEntry => {
          for (let i = 0; i < this.state.itemsForGrid.length; i++) {
            const transLogHistory = {
              Title: this.state.itemsForGrid[i].documentID,
              Status: "Registered for transmittal with " + this.state.itemsForGrid[i].transmitFor + " " + this.state.transmitTo,
              DocumentIndexId: this.state.itemsForGrid[i].documentIndexId,
              LogDate: this.currentDate,
            }
            this._Service.addToList(this.props.siteUrl, this.props.transmittalHistoryLogList, transLogHistory);
          }
        }).then(masgDisplay => {
          // alert("success");
          this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "Saved Successfully", messageType: 4 }, });
          setTimeout(() => {
            window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
          }, 10000);
        });
        this.validator.hideMessages();
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
    else {
      let selectedContactsTo = (this.state.selectedContactsTo === null) ? "" : this.state.selectedContactsTo.toString();
      let selectedContactsCC = (this.state.selectedContactsCC === null) ? "" : this.state.selectedContactsCC.toString();
      if (this.state.transmitTo !== "" && this.state.transmittalTypekey !== "") {
        this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "Saved Successfully", messageType: 4 }, });
        const OHLUpdate = {
          ToEmails: selectedContactsTo,
          CCEmails: selectedContactsCC,
          ToName: this.state.selectedContactsToDisplayName,
          CCName: this.state.selectedContactsCCDisplayName,
          Notes: this.state.notes,
          TransmittalType: this.state.transmittalType,
          TransmittedById: this.state.currentUser,
          SendAsSharedFolder: this.state.sendAsSharedFolder,
          ReceiveInSharedFolder: this.state.recieveInSharedFolder,
          SendAsMultipleEmails: this.state.sendAsMultipleFolder,
          TotalFiles: (totalFiles).toString(),
          TransmittalSize: (this.state.itemsForGrid.length === 0) ? "" : (convertKBtoMB).toString(),
          CoverLetter: this.state.coverLetterNeeded,
          InternalCCId: { results: this.state.internalCCContacts, }
        }
        this._Service.updateList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, OHLUpdate, Number(this.transmittalID))
          .then((afterOutboundUpdtation: any) => {
            if (this.state.itemsForGrid.length > this.state.currentOutboundDetailItem.length) {
              for (var k = this.state.currentOutboundDetailItem.length; k < this.state.itemsForGrid.length; k++) {
                const OTD = {
                  Title: this.state.itemsForGrid[k].documentName,
                  TransmittalHeaderId: this.transmittalID,
                  DocumentIndexId: this.state.itemsForGrid[k].documentIndexId,
                  Revision: this.state.itemsForGrid[k].revision,
                  TransmittalRevision: this.state.itemsForGrid[k].revision,
                  DueDate: this.state.itemsForGrid[k].dueDate,
                  Size: this.state.itemsForGrid[k].fileSizeInMB,
                  SentComments: this.state.itemsForGrid[k].comments,
                  CustomerAcceptanceCodeId: (this.state.sourceDocumentItem !== null) ? this.state.itemsForGrid[k].acceptanceCode : null,
                  TransmittedForId: this.state.itemsForGrid[k].transmitForKey,
                  TransmitFor: this.state.itemsForGrid[k].transmitFor,
                  ApprovalRequired: this.state.itemsForGrid[k].approvalRequired,
                  DocumentLibraryID: this.state.itemsForGrid[k].publishDoumentlibraryID,
                  Slno: (Number(k) + Number(1)).toString(),
                  CustomerDocumentNo: this.state.itemsForGrid[k].customerDocumentNo,
                  SubcontractorDocumentNo: this.state.itemsForGrid[k].subcontractorDocumentNo,
                }
                this._Service.addToList(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, OTD)
                  .then((transmittallog: any) => {
                    const TLH = {
                      Title: this.state.itemsForGrid[k].documentID,
                      Status: "Registered for transmittal with" + this.state.itemsForGrid[k].transmitFor + this.state.transmitTo,
                      DocumentIndexId: this.state.itemsForGrid[k].documentIndexId,
                      LogDate: this.currentDate,
                    }
                    this._Service.addToList(this.props.siteUrl, this.props.transmittalHistoryLogList, TLH);
                  });
              }
            }
          }).then(async (outboundAdditionaldocumentUpdate: any) => {
            if (this.state.itemsForExternalGrid.length > this.state.currentOutboundAdditionalItem.length) {
              for (let k = this.state.currentOutboundAdditionalItem.length; k < this.state.itemsForExternalGrid.length; k++) {
                var splitted = this.state.itemsForExternalGrid[k].documentName.split(".");
                let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + this.state.transmittalNo + '.' + splitted[splitted.length - 1];
                let docName = documentNameExtension;
                console.log(this.state.itemsForExternalGrid);
                const additionalMetadataUpdate = {
                  Title: this.state.itemsForExternalGrid[k].documentName,
                  TransmittalIDId: Number(this.transmittalID),
                  Size: this.state.itemsForExternalGrid[k].fileSizeInMB,
                  Comments: this.state.itemsForExternalGrid[k].externalComments,
                  TransmittalStatus: "Draft",
                  Slno: (Number(k) + Number(1)).toString(),
                }
                this._Service.uploadDocument(docName, this.state.itemsForExternalGrid[k].content, this.props.outboundAdditionalDocumentsListName, additionalMetadataUpdate);

              }
            }
          }).then((msgDisplay: any) => {
            this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "Saved Successfully", messageType: 4 }, });
            setTimeout(() => {
              window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
            }, 9000);
            //window.location.replace(window.location.protocol+"//"+window.location.hostname+"/"+this.props.siteUrl); 
          });
        this.validator.hideMessages();
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }

  }
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

      let totalSize = add(totalAdditional, totalsizeProjects);
      let convertKBtoMB = Number(totalSize).toFixed(2);
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
  private _confirmAndSendBtnClick() {
    let selectedContactsTo = (this.state.selectedContactsTo === null) ? "" : this.state.selectedContactsTo.toString();
    let selectedContactsCC = (this.state.selectedContactsCC === null) ? "" : this.state.selectedContactsCC.toString();
    console.log(this.state.selectedContactsToDisplayName);
    let sourceDocumentId;
    let hidden = 1;
    let statusCount = 0;
    let forTransmittalStatus: string;
    //total files
    let totalFiles: number;
    let convertKBtoMB: string;
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
    if (Number(convertKBtoMB) > 25 && (this.state.sendAsSharedFolder === false)) {
      this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "File size is greater than 25 MB.Please select the checkbox Send and Receive as shared folder", messageType: 1 }, });
    }
    else {
      if (this.transmittalID === null || this.transmittalID === "") {
        let selectedContactsTo = this.state.selectedContactsTo.toString();
        let selectedContactsCC = this.state.selectedContactsCC.toString();
        //when add butten not clicked
        if (this.state.itemsForGrid.length === 0 && this.state.transmitTo !== "" && this.state.transmittalTypekey !== "" && this.state.selectedContactsTo !== null && this.validator.fieldValid("selectedContactsTo")) {
          this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "Please click the project add button", messageType: 1 }, });
        }
        else {
          this.setState({ normalMsgBar: "none", statusMessage: { isShowMessage: false, message: "Please click the project add button", messageType: 1 }, });
        }
        if (this.state.transmitTo !== "" && this.state.itemsForGrid.length !== 0 && this.state.transmittalTypekey !== "" && this.state.publishDocumentsItemsForGrid.length !== 0 && this.state.selectedContactsTo !== null && this.validator.fieldValid("selectedContactsTo")) {
          this.setState({
            spinnerDiv: "",
            hideButtonAfterSubmit: "none",
            hideUnlockButton: "none",
          });
          this._trannsmittalIDGeneration().then(afterIdgeneration => {
            //header list
            const items = {
              Title: this.state.transmittalNo,
              TransmittalCategory: this.state.transmitTo,
              Customer: this.state.customerName,
              CustomerID: (this.state.transmitTo === "Customer") ? this.state.customerId : "",
              SubContractor: this.state.subContractor,
              SubContractorID: (this.state.subContractorKey).toString(),
              ToEmails: selectedContactsTo,
              CCEmails: selectedContactsCC,
              Notes: this.state.notes,
              TransmittalStatus: (this.state.transmitTo !== "Customer") ? "Completed" : "Ongoing",
              TransmittalType: this.state.transmittalType,
              TransmittedById: this.state.currentUser,
              SendAsSharedFolder: this.state.sendAsSharedFolder,
              ReceiveInSharedFolder: this.state.recieveInSharedFolder,
              SendAsMultipleEmails: this.state.sendAsMultipleFolder,
              TransmittalSize: (convertKBtoMB).toString(),
              TransmittalDate: this.currentDate,
              TotalFiles: (totalFiles).toString(),
              ToName: this.state.selectedContactsToDisplayName,
              CCName: this.state.selectedContactsCCDisplayName,
              CoverLetter: this.state.coverLetterNeeded,
              InternalCCId: { results: this.state.internalCCContacts, }
            }
            this._Service.addToList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, items)
              .then(async (outboundTransmittalHeader: { data: { ID: string; }; }) => {
                this.setState({ outboundTransmittalHeaderId: outboundTransmittalHeader.data.ID });
                for (let i = 0; i < this.state.itemsForGrid.length; i++) {
                  if (this.state.itemsForGrid[i].approvalRequired === true) {
                    forTransmittalStatus = "true";
                  }
                }
                const UpdateLinks = {
                  TransmittalLink: {
                    Description: "Project workspace",
                    Url: this.props.siteUrl + "/SitePages/" + this.props.outBoundTransmittalSitePage + ".aspx?trid=" + outboundTransmittalHeader.data.ID + ""
                  },
                  TransmittalDetails: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Transmittal Details",
                    Url: this.props.siteUrl + "/Lists/" + this.props.outboundTransmittalDetailsListName + "/AllItems.aspx?FilterField1=TransmittalHeader&FilterValue1=" + outboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=6da3a1b3%2D0155%2D48d9%2Da7c7%2Dd2e862c07db5"
                  },
                  OutboundAdditionalDetails: {
                    "__metadata": { type: "SP.FieldUrlValue" },
                    Description: "Outbound Additional Details",
                    Url: this.props.siteUrl + "/" + this.props.outboundAdditionalDocumentsListName + "/Forms/AllItems.aspx?FilterField1=TransmittalID&FilterValue1=" + outboundTransmittalHeader.data.ID + "&FilterType1=Lookup&viewid=bcc64a99-0907-4416-b9f6-8001acf1e000"
                  },

                  TransmittalStatus: forTransmittalStatus !== "true" ? "Completed" : "Ongoing",
                }
                await this._Service.updateList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, UpdateLinks, Number(outboundTransmittalHeader.data.ID));

                if (this.state.transmitTo != "Customer") {
                  const tSUpdate = {
                    TransmittalStatus: "Completed",
                  }
                  this._Service.updateList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, tSUpdate, Number(outboundTransmittalHeader.data.ID));
                }
                //outbound Details
                if (this.state.itemsForGrid.length > 0) {
                  for (var i in this.state.itemsForGrid) {
                    const detailItems = {
                      Title: this.state.itemsForGrid[i].documentName,
                      TransmittalHeaderId: outboundTransmittalHeader.data.ID,
                      DocumentIndexId: this.state.itemsForGrid[i].documentIndexId,
                      Revision: this.state.itemsForGrid[i].revision,
                      TransmittalRevision: this.state.itemsForGrid[i].revision,
                      DueDate: this.state.itemsForGrid[i].dueDate,
                      Size: this.state.itemsForGrid[i].fileSizeInMB,
                      SentComments: this.state.itemsForGrid[i].comments,
                      CustomerAcceptanceCodeId: this.state.itemsForGrid[i].acceptanceCode,
                      TransmittedForId: this.state.itemsForGrid[i].transmitForKey,
                      TransmitFor: this.state.itemsForGrid[i].transmitFor,
                      ApprovalRequired: this.state.itemsForGrid[i].approvalRequired,
                      TransmittalStatus: (this.state.itemsForGrid[i].approvalRequired === true && this.state.transmitTo == "Customer") ? "Ongoing" : "Completed",
                      DocumentLibraryID: this.state.itemsForGrid[i].publishDoumentlibraryID,
                      Slno: (Number(i) + Number(1)).toString(),
                      CustomerDocumentNo: this.state.itemsForGrid[i].customerDocumentNo,
                      SubcontractorDocumentNo: this.state.itemsForGrid[i].subcontractorDocumentNo,
                    }
                    this._Service.addToList(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, detailItems);
                  }
                }
                //outbound additional
                if (this.state.itemsForExternalGrid.length > 0) {
                  for (var i in this.state.itemsForExternalGrid) {
                    var splitted = this.state.itemsForExternalGrid[i].documentName.split(".");
                    let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + this.state.transmittalNo + '.' + splitted[splitted.length - 1];
                    let docName = documentNameExtension;

                    console.log(this.state.itemsForExternalGrid);
                    const additionalMetadataUpdate = {
                      Title: this.state.itemsForExternalGrid[i].documentName,
                      TransmittalIDId: this.state.outboundTransmittalHeaderId,
                      Size: this.state.itemsForExternalGrid[i].fileSizeInMB,
                      Comments: this.state.itemsForExternalGrid[i].externalComments,
                      SentDate: this.currentDate,
                      TransmittalStatus: "Ongoing",
                      Slno: (Number(i) + Number(1)).toString(),
                    }
                    this._Service.uploadDocument(docName, this.state.itemsForExternalGrid[i].content, this.props.outboundAdditionalDocumentsListName, additionalMetadataUpdate);


                  }
                }
                for (let i = 0; i < this.state.itemsForGrid.length; i++) {
                  if (this.state.transmitTo === "Customer") {
                    this._Service.getItembyID(this.props.siteUrl, this.props.documentIndex, this.state.itemsForGrid[i].documentIndexId)
                      .then((CurrentTransmittalItems: { hiddenFieldForOutboundTransmitta: number; CurrentTransmittalId: any; CurrentActualSubmitedDate: any; CurrentTransmittalRevision: any; }) => {
                        console.log("CurrentTransmittalItems", CurrentTransmittalItems);
                        if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === null || CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 5) {
                          const updateLists = {
                            TransmittalStatus: this.state.itemsForGrid[i].approvalRequired === true ? "Ongoing" : "Completed",
                            TransmittalLocation: "Out to " + this.state.transmitTo,
                            Workflow: "Transmittal",
                            TransmittalRevision: this.state.itemsForGrid[i].revision,
                            CurrentTransmittalId: this.state.transmittalNo,
                            CurrentActualSubmitedDate: this.currentDate,
                            CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                            TransmittalDueDate: this.state.itemsForGrid[i].dueDate,
                            hiddenFieldForOutboundTransmitta: Number(hidden),
                          }
                          this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateLists, this.state.itemsForGrid[i].documentIndexId);
                        }
                        else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta == 1) {
                          const UpdateDLOne = {
                            TransmittalStatus: this.state.itemsForGrid[i].approvalRequired === true ? "Ongoing" : "Completed",
                            TransmittalLocation: "Out to " + this.state.transmitTo,
                            Workflow: "Transmittal",
                            TransmittalDueDate: this.state.itemsForGrid[i].dueDate,
                            TransmittalRevision: this.state.itemsForGrid[i].revision,
                            TransmittalId1: CurrentTransmittalItems.CurrentTransmittalId,
                            ActualSubmitedDate1: CurrentTransmittalItems.CurrentActualSubmitedDate,
                            TransmittalRevision1: CurrentTransmittalItems.CurrentTransmittalRevision,
                            CurrentTransmittalId: this.state.transmittalNo,
                            CurrentActualSubmitedDate: this.currentDate,
                            CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                            hiddenFieldForOutboundTransmitta: Number(CurrentTransmittalItems.hiddenFieldForOutboundTransmitta + 1),
                          }
                          this._Service.updateList(this.props.siteUrl, this.props.documentIndex, UpdateDLOne, this.state.itemsForGrid[i].documentIndexId);
                        }
                        else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 2) {
                          const updateList3 = {
                            TransmittalStatus: this.state.itemsForGrid[i].approvalRequired === true ? "Ongoing" : "Completed",
                            TransmittalLocation: "Out to " + this.state.transmitTo,
                            Workflow: "Transmittal",
                            TransmittalDueDate: this.state.itemsForGrid[i].dueDate,
                            TransmittalRevision: this.state.itemsForGrid[i].revision,
                            TransmittalId2: CurrentTransmittalItems.CurrentTransmittalId,
                            ActualSubmitedDate2: CurrentTransmittalItems.CurrentActualSubmitedDate,
                            TransmittalRevision2: CurrentTransmittalItems.CurrentTransmittalRevision,
                            CurrentTransmittalId: this.state.transmittalNo,
                            CurrentActualSubmitedDate: this.currentDate,
                            CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                            hiddenFieldForOutboundTransmitta: Number(CurrentTransmittalItems.hiddenFieldForOutboundTransmitta + 1),
                          }
                          this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateList3, this.state.itemsForGrid[i].documentIndexId);
                        }
                        else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 3) {
                          const updateList4 = {
                            TransmittalStatus: this.state.itemsForGrid[i].approvalRequired === true ? "Ongoing" : "Completed",
                            TransmittalLocation: "Out to " + this.state.transmitTo,
                            Workflow: "Transmittal",
                            TransmittalDueDate: this.state.itemsForGrid[i].dueDate,
                            TransmittalRevision: this.state.itemsForGrid[i].revision,
                            TransmittalId3: CurrentTransmittalItems.CurrentTransmittalId,
                            ActualSubmitedDate3: CurrentTransmittalItems.CurrentActualSubmitedDate,
                            TransmittalRevision3: CurrentTransmittalItems.CurrentTransmittalRevision,
                            CurrentTransmittalId: this.state.transmittalNo,
                            CurrentActualSubmitedDate: this.currentDate,
                            CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                            hiddenFieldForOutboundTransmitta: Number(CurrentTransmittalItems.hiddenFieldForOutboundTransmitta + 1),
                          }
                          this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateList4, this.state.itemsForGrid[i].documentIndexId);
                        } else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 4) {
                          const updateList5 = {
                            TransmittalStatus: this.state.itemsForGrid[i].approvalRequired === true ? "Ongoing" : "Completed",
                            TransmittalLocation: "Out to " + this.state.transmitTo,
                            Workflow: "Transmittal",
                            TransmittalDueDate: this.state.itemsForGrid[i].dueDate,
                            TransmittalRevision: this.state.itemsForGrid[i].revision,
                            TransmittalId4: CurrentTransmittalItems.CurrentTransmittalId,
                            ActualSubmitedDate4: CurrentTransmittalItems.CurrentActualSubmitedDate,
                            TransmittalRevision4: CurrentTransmittalItems.CurrentTransmittalRevision,
                            CurrentTransmittalId: this.state.transmittalNo,
                            CurrentActualSubmitedDate: this.currentDate,
                            CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                            hiddenFieldForOutboundTransmitta: Number(CurrentTransmittalItems.hiddenFieldForOutboundTransmitta + 1),
                          }
                          this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateList5, this.state.itemsForGrid[i].documentIndexId);
                        }
                        else {
                          const updateList6 = {
                            TransmittalStatus: this.state.itemsForGrid[i].approvalRequired === true ? "Ongoing" : "Completed",
                            TransmittalLocation: "Out to " + this.state.transmitTo,
                            Workflow: "Transmittal",
                            TransmittalDueDate: this.state.itemsForGrid[i].dueDate,
                            TransmittalRevision: this.state.itemsForGrid[i].revision,
                            CurrentTransmittalId: this.state.transmittalNo,
                            CurrentActualSubmitedDate: this.currentDate,
                            CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                            hiddenFieldForOutboundTransmitta: Number(hidden),
                          }
                          this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateList6, this.state.itemsForGrid[i].documentIndexId);

                        }
                      });
                  }
                  const updateList7 = {
                    TransmittalStatus: (this.state.itemsForGrid[i].approvalRequired === true && this.state.transmitTo === "Customer") ? "Ongoing" : "Completed",
                    TransmittalLocation: "Out to " + this.state.transmitTo,
                    Workflow: "Transmittal",
                    TransmittalRevision: this.state.itemsForGrid[i].revision,
                    TransmittalDueDate: this.state.itemsForGrid[i].dueDate,
                    CurrentTransmittalId: this.state.transmittalNo,
                    CurrentActualSubmitedDate: this.currentDate,
                    CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                  }
                  this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateList7, this.state.itemsForGrid[i].documentIndexId);
                  this._Service.getItemWithSelectAndExpandWithId(this.props.siteUrl, this.props.documentIndex, "Owner/ID,Owner/Title,Owner/EMail,DocumentName,DocumentID", "Owner", this.state.itemsForGrid[i].documentIndexId)
                    .then((forGettingOwner: { DocumentID: any; Owner: { EMail: any; Title: string | ReplacementFunction; }; DocumentName: string | ReplacementFunction; }) => {
                      const historyLog = {
                        Title: forGettingOwner.DocumentID,
                        Status: "Registered for transmittal with " + this.state.itemsForGrid[i].transmitFor + " " + this.state.transmitTo,
                        DocumentIndexId: this.state.itemsForGrid[i].documentIndexId,
                        LogDate: this.currentDate,
                      }
                      this._Service.addToList(this.props.siteUrl, this.props.transmittalHistoryLogList, historyLog);
                      const historyLogOutTo = {
                        Title: forGettingOwner.DocumentID,
                        Status: "OUT to " + this.state.transmitTo,
                        DocumentIndexId: this.state.itemsForGrid[i].documentIndexId,
                        LogDate: this.currentDate,
                      }
                      this._Service.addToList(this.props.siteUrl, this.props.transmittalHistoryLogList, historyLogOutTo);
                      //mail for owner
                      this._sendAnEmailUsingMSGraph(forGettingOwner.Owner.EMail, "OutboundTransmittalSending", forGettingOwner.Owner.Title, forGettingOwner.DocumentName);
                    });
                }//forloop from transmittal details 
                let selectHeaderItems = "DocumentIndex/ID,DocumentIndex/Title,DueDate,SentComments,Revision,Title,Size,TransmittedFor/ID,TransmittedFor/Title,Temporary,TransmittalHeader/ID,DocumentLibraryID,ID,ApprovalRequired";
                let expand = "DocumentIndex,TransmittedFor,TransmittalHeader";
                let filter = "TransmittalHeader/ID eq '" + Number(this.state.outboundTransmittalHeaderId) + "' ";
                this._Service.getItemForSelectExpandInListsWithFilter(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, selectHeaderItems, filter, expand)
                  .then((outboundTransmittalDetailsListName: any) => {
                    this.setState({
                      currentOutboundDetailItem: outboundTransmittalDetailsListName,
                    });
                    if (this.state.transmitTo === "Customer") {
                    }
                  });
              }).then((after: any) => {
                this.triggerOutboundTransmittal(Number(this.state.outboundTransmittalHeaderId));
                let selectHeaderItems = "DocumentIndex/ID,DocumentIndex/Title,DueDate,SentComments,Revision,Title,Size,TransmittedFor/ID,TransmittedFor/Title,Temporary,TransmittalHeader/ID,DocumentLibraryID,ID,ApprovalRequired";
                let expand = "DocumentIndex,TransmittedFor,TransmittalHeader";
                let filter = "TransmittalHeader/ID eq '" + Number(this.state.outboundTransmittalHeaderId) + "' ";
                this._Service.getItemForSelectExpandInListsWithFilter(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, selectHeaderItems, filter, expand)
                  .then((outboundTransmittalDetailsListName: any) => {
                    this.setState({
                      currentOutboundDetailItem: outboundTransmittalDetailsListName,
                    });
                  }).then((aftermail: any) => {
                    this.setState({
                      hideButtonAfterSubmit: "none",
                      hideUnlockButton: "none",
                      spinnerDiv: "",
                    });
                  });
              });
          });
          this.validator.hideMessages();
        }
        else {
          this.validator.showMessages();
          this.forceUpdate();
        }
      }
      else {
        let selectedContactsTo = (this.state.selectedContactsTo === null) ? "" : this.state.selectedContactsTo.toString();
        let selectedContactsCC = (this.state.selectedContactsCC === null) ? "" : this.state.selectedContactsCC.toString();
        if (this.state.itemsForGrid.length === 0 && this.state.transmitTo !== "" && this.state.transmittalTypekey !== "" && this.state.selectedContactsTo !== null && this.validator.fieldValid("selectedContactsTo")) {
          this.setState({ normalMsgBar: "", statusMessage: { isShowMessage: true, message: "Please click the project add button", messageType: 1 }, });
        }
        else {
          this.setState({ normalMsgBar: "none", statusMessage: { isShowMessage: false, message: "Please click the project add button", messageType: 1 }, });
        }
        if (this.state.transmitTo !== "" && this.state.transmittalTypekey !== "" && this.state.selectedContactsTo !== null && this.state.itemsForGrid.length !== 0 && this.validator.fieldValid("selectedContactsTo")) {
          this.setState({
            spinnerDiv: "",
            hideButtonAfterSubmit: "none",
            hideUnlockButton: "none",
          });
          this.validator.hideMessages();
          const updateLists1 = {
            ToEmails: selectedContactsTo,
            CCEmails: selectedContactsCC,
            ToName: this.state.selectedContactsToDisplayName,
            CCName: this.state.selectedContactsCCDisplayName,
            Notes: this.state.notes,
            TransmittalType: this.state.transmittalType,
            TransmittedById: this.state.currentUser,
            SendAsSharedFolder: this.state.sendAsSharedFolder,
            ReceiveInSharedFolder: this.state.recieveInSharedFolder,
            SendAsMultipleEmails: this.state.sendAsMultipleFolder,
            TransmittalSize: (convertKBtoMB).toString(),
            TransmittalDate: this.currentDate,
            TransmittalStatus: (this.state.transmitTo !== "Customer") ? "Completed" : "Ongoing",
            TotalFiles: (totalFiles).toString(),
            CoverLetter: this.state.coverLetterNeeded,
            InternalCCId: { results: this.state.internalCCContacts, }
          }
          //header list
          this._Service.updateList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, updateLists1, Number(this.transmittalID))
            .then((afterOutboundUpdtation: any) => {
              if (this.state.itemsForGrid.length > this.state.currentOutboundDetailItem.length) {
                //detail list
                const updateDetailsLists = {
                  Title: this.state.itemsForGrid[k].documentName,
                  TransmittalHeaderId: this.transmittalID,
                  DocumentIndexId: this.state.itemsForGrid[k].documentIndexId,
                  Revision: this.state.itemsForGrid[k].revision,
                  TransmittalRevision: this.state.itemsForGrid[k].revision,
                  DueDate: this.state.itemsForGrid[k].dueDate,
                  Size: this.state.itemsForGrid[k].fileSizeInMB,
                  SentComments: this.state.itemsForGrid[k].comments,
                  CustomerAcceptanceCodeId: (this.state.sourceDocumentItem !== null) ? this.state.itemsForGrid[k].acceptanceCode : null,
                  TransmittedForId: this.state.itemsForGrid[k].transmitForKey,
                  TransmitFor: this.state.itemsForGrid[k].transmitFor,
                  ApprovalRequired: this.state.itemsForGrid[k].approvalRequired,
                  TransmittalStatus: (this.state.itemsForGrid[k].approvalRequired === true && this.state.transmitTo == "Customer") ? "Ongoing" : "Completed",
                  DocumentLibraryID: this.state.itemsForGrid[k].publishDoumentlibraryID,
                  Slno: (Number(k) + Number(1)).toString(),
                  CustomerDocumentNo: this.state.itemsForGrid[k].customerDocumentNo,
                  SubcontractorDocumentNo: this.state.itemsForGrid[k].subcontractorDocumentNo,
                }
                for (var k = this.state.currentOutboundDetailItem.length; k < this.state.itemsForGrid.length; k++) {
                  this._Service.addToList(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, updateDetailsLists);

                }
              }
            }).then(async (outboundAdditionaldocumentUpdate: any) => {
              let docName;
              if (this.state.itemsForExternalGrid.length > this.state.currentOutboundAdditionalItem.length) {
                for (let k = this.state.currentOutboundAdditionalItem.length; k < this.state.itemsForExternalGrid.length; k++) {
                  var splitted = this.state.itemsForExternalGrid[k].documentName.split(".");
                  let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + this.state.transmittalNo + '.' + splitted[splitted.length - 1];
                  let docName = documentNameExtension;
                  const additionalMetadataUpdate = {
                    Title: this.state.itemsForExternalGrid[k].documentName,
                    TransmittalIDId: Number(this.transmittalID),
                    Size: this.state.itemsForExternalGrid[k].fileSizeInMB,
                    Comments: this.state.itemsForExternalGrid[k].externalComments,
                    SentDate: this.currentDate,
                    TransmittalStatus: "Ongoing",
                    Slno: (Number(k) + Number(1)).toString(),
                  }
                  this._Service.uploadDocument(docName, this.state.itemsForExternalGrid[k].content, this.props.outboundAdditionalDocumentsListName, additionalMetadataUpdate);
                }
              }
            }).then((result: any) => {
              let selectHeaderItems = "DocumentIndex/ID,DocumentIndex/Title,DueDate,SentComments,Revision,Title,Size,TransmittedFor/ID,TransmittedFor/Title,Temporary,TransmittalHeader/ID,DocumentLibraryID,ID,ApprovalRequired";
              let expand = "DocumentIndex,TransmittedFor,TransmittalHeader";
              let filter = "TransmittalHeader/ID eq '" + Number(this.transmittalID) + "' ";
              this._Service.getItemForSelectExpandInListsWithFilter(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, selectHeaderItems, filter, expand)
                .then((outboundTransmittalDetailsListName: any) => {
                  console.log("outboundTransmittalDetailsListName", outboundTransmittalDetailsListName);
                  this.setState({
                    currentOutboundDetailItem: outboundTransmittalDetailsListName,
                  });
                  for (let i = 0; i < this.state.currentOutboundDetailItem.length; i++) {
                    const StatusUpdate = {
                      TransmittalStatus: (this.state.currentOutboundDetailItem[i].ApprovalRequired === true && this.state.transmitTo === "Customer") ? "Ongoing" : "Completed",
                    }
                    this._Service.updateList(this.props.siteUrl, this.props.outboundTransmittalDetailsListName, StatusUpdate, this.state.currentOutboundDetailItem[i].ID)

                    if (this.state.transmitTo === "Customer") {
                      this._Service.getItembyID(this.props.siteUrl, this.props.documentIndex, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)
                        .then((CurrentTransmittalItems: { hiddenFieldForOutboundTransmitta: number; CurrentTransmittalId: any; CurrentActualSubmitedDate: any; CurrentTransmittalRevision: any; }) => {
                          console.log("CurrentTransmittalItems", CurrentTransmittalItems);
                          const updatingIndex = {
                            TransmittalStatus: this.state.currentOutboundDetailItem[i].ApprovalRequired === true ? "Ongoing" : "Completed",
                            TransmittalLocation: "Out to " + this.state.transmitTo,
                            Workflow: "Transmittal",
                            TransmittalDueDate: this.state.currentOutboundDetailItem[i].DueDate,
                            TransmittalRevision: this.state.currentOutboundDetailItem[i].Revision,
                            TransmittalId1: CurrentTransmittalItems.CurrentTransmittalId,
                            ActualSubmitedDate1: CurrentTransmittalItems.CurrentActualSubmitedDate,
                            TransmittalRevision1: CurrentTransmittalItems.CurrentTransmittalRevision,
                            CurrentTransmittalId: this.state.transmittalNo,
                            CurrentActualSubmitedDate: this.currentDate,
                            CurrentTransmittalRevision: this.state.currentOutboundDetailItem[i].Revision,
                            hiddenFieldForOutboundTransmitta: Number(CurrentTransmittalItems.hiddenFieldForOutboundTransmitta + 1),
                          }
                          if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === null || CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 5) {
                            const UpdateIndexWhennull = {
                              TransmittalStatus: this.state.currentOutboundDetailItem[i].ApprovalRequired === true ? "Ongoing" : "Completed",
                              TransmittalLocation: "Out to " + this.state.transmitTo,
                              Workflow: "Transmittal",
                              TransmittalDueDate: this.state.currentOutboundDetailItem[i].DueDate,
                              TransmittalRevision: this.state.currentOutboundDetailItem[i].Revision,
                              CurrentTransmittalId: this.state.transmittalNo,
                              CurrentActualSubmitedDate: this.currentDate,
                              CurrentTransmittalRevision: this.state.itemsForGrid[i].revision,
                              hiddenFieldForOutboundTransmitta: Number(hidden),
                            }
                            this._Service.updateList(this.props.siteUrl, this.props.documentIndex, UpdateIndexWhennull, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)

                          }
                          else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 1) {
                            this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updatingIndex, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)

                          }
                          else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 2) {
                            this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updatingIndex, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)

                          }
                          else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 3) {
                            this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updatingIndex, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)
                          } else if (CurrentTransmittalItems.hiddenFieldForOutboundTransmitta === 4) {
                            this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updatingIndex, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)
                          }
                          else {
                            const updateIn = {
                              TransmittalStatus: this.state.currentOutboundDetailItem[i].ApprovalRequired == true ? "Ongoing" : "Completed",
                              TransmittalLocation: "Out to " + this.state.transmitTo,
                              Workflow: "Transmittal",
                              TransmittalDueDate: this.state.currentOutboundDetailItem[i].DueDate,
                              TransmittalRevision: this.state.currentOutboundDetailItem[i].Revision,
                              CurrentTransmittalId: this.state.transmittalNo,
                              CurrentActualSubmitedDate: this.currentDate,
                              CurrentTransmittalRevision: this.state.currentOutboundDetailItem[i].Revision,
                              hiddenFieldForOutboundTransmitta: Number(hidden),
                            }
                            this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateIn, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)
                          }
                        });
                    }
                    const updateOne = {
                      TransmittalStatus: (this.state.currentOutboundDetailItem[i].ApprovalRequired === true && this.state.transmitTo === "Customer") ? "Ongoing" : "Completed",
                      TransmittalLocation: "Out to " + this.state.transmitTo,
                      Workflow: "Transmittal",
                      TransmittalDueDate: this.state.currentOutboundDetailItem[i].DueDate,
                      TransmittalRevision: this.state.currentOutboundDetailItem[i].Revision,
                      CurrentTransmittalId: this.state.transmittalNo,
                      CurrentActualSubmitedDate: this.currentDate,
                      CurrentTransmittalRevision: this.state.currentOutboundDetailItem[i].Revision,
                    }
                    this._Service.updateList(this.props.siteUrl, this.props.documentIndex, updateOne, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)

                    this._Service.getlistItemById(this.props.siteUrl, this.props.documentIndex, this.state.currentOutboundDetailItem[i].DocumentIndex.ID)
                      .then(async (sourceDocumentID: { DocumentID: any; }) => {
                        const THLogList = {
                          Title: sourceDocumentID.DocumentID,
                          Status: "OUT to " + this.state.transmitTo,
                          DocumentIndexId: this.state.currentOutboundDetailItem[i].DocumentIndex.ID,
                          LogDate: this.currentDate,
                        }
                        await this._Service.addToList(this.props.siteUrl, this.props.transmittalHistoryLogList, THLogList);
                        const select = "Owner/ID,Owner/Title,Owner/EMail,DocumentName";
                        const expand = "Owner";
                        this._Service.getItemWithSelectAndExpand(this.props.siteUrl, this.props.documentIndex, select, expand)
                          .then((forGettingOwner: { Owner: { EMail: any; Title: string | ReplacementFunction; }; DocumentName: string | ReplacementFunction; }) => {
                            this._sendAnEmailUsingMSGraph(forGettingOwner.Owner.EMail, "OutboundTransmittalSending", forGettingOwner.Owner.Title, forGettingOwner.DocumentName);
                          });
                      });
                  }//forloop from transmittal details           
                  if (this.state.transmitTo === "Customer") {
                    for (let i = 0; i < this.state.currentOutboundDetailItem.length; i++) {
                      if (this.state.currentOutboundDetailItem[i].ApprovalRequired !== true) {
                        statusCount++;
                      }
                    }
                    if (Number(statusCount) === this.state.currentOutboundDetailItem.length) {
                      const OTUpdate = {
                        TransmittalStatus: "Completed",
                      }
                      this._Service.updateList(this.props.siteUrl, this.props.outboundTransmittalHeaderListName, OTUpdate, Number(this.transmittalID));

                    }
                  }
                  this.triggerOutboundTransmittal(Number(this.transmittalID));
                }).then((after: any) => {
                }).then((aftergettingDetails: any) => {
                }).then((aftermail: any) => {
                  this.setState({
                    hideButtonAfterSubmit: "none",
                    hideUnlockButton: "none",
                    spinnerDiv: "",
                  });
                });
            });
        }
        else {
          this.validator.showMessages();
          this.forceUpdate();
        }
      }
    }
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
    if (this.state.transmitTo === "Customer") { transmitTo = "Outbound Customer"; }
    else if (this.state.transmitTo === "Sub-Contractor") { transmitTo = "Outbound Sub-contractor"; }
    await this._Service.getItemForSelectInListsWithFilter(this.props.siteUrl, this.props.transmittalIdSettingsListName, "TransmittalCategory eq '" + transmitTo + "' and(TransmittalType eq '" + this.state.transmittalType + "')")
      .then((transmittalIdSettingsItems: {
        Counter: any;
        Title: any;
        SequenceNumber: any;
        Separator: any;
        Prefix: any; ID: any;
      }[]) => {
        console.log("transmittalIdSettingsItems", transmittalIdSettingsItems);
        prefix = transmittalIdSettingsItems[0].Prefix;
        separator = transmittalIdSettingsItems[0].Separator;
        sequenceNumber = transmittalIdSettingsItems[0].SequenceNumber;
        title = transmittalIdSettingsItems[0].Title;
        counter = transmittalIdSettingsItems[0].Counter;
        let increment = counter + 1;
        var incrementValue = increment.toString();
        this._transmittalSequenceNumber(incrementValue, sequenceNumber);
        transmittalID = prefix + separator + title + separator + this.state.projectNumber + separator + this.state.incrementSequenceNumber;
        console.log("transmittalID", transmittalID);
        this.setState({
          transmittalNo: transmittalID,
        });
        //counter updation
        const idSettings = {
          Counter: increment
        }
        this._Service.updateList(this.props.siteUrl, this.props.transmittalIdSettingsListName, idSettings, transmittalIdSettingsItems[0].ID);
      });
  }
  protected async triggerProjectPermissionFlow(PostUrl: any) {
    //alert("triggerProjectPermissionFlow")
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    // alert("In function");
    // alert(transmittalID);
    const postURL = PostUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'PermissionTitle': 'Project_UnlockTransmittal',
      'SiteUrl': siteUrl,
      'CurrentUserEmail': this.props.context.pageContext.user.email
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
      if (responseJSON['Status'] === "Valid") {
        // this.setState({
        //   loaderDisplay: "none",
        //   webpartView: "",
        // });
        //this._queryParamGetting();
        this.permissionForRecall = "Yes";
        this.setState({
          hideButtonAfterSubmit: "none",
          hideUnlockButton: "",
        });
      }
      else {
        // this.setState({
        //   webpartView: "none",
        //   loaderDisplay: "none",
        //   accessDeniedMsgBar: "",
        //   statusMessage: { isShowMessage: true, message: "You are not permitted to perform this operations", messageType: 1 },
        // });
        // setTimeout(() => {
        //   window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
        // }, 20000);
      }

    }
    else { }

  }
  private _recallTransmittalConfirmation() {
    this.setState({
      recallConfirmMsgDiv: "",
      recallConfirmMsg: false,
    });
  }

  private _LAUrlGettingForRecall = async () => {
    const laUrl = await this.reqWeb.getList("/sites/" + this.props.hubSite + "/Lists/" + this.props.masterListName)
      .items.filter("Title eq 'EMEC_RecallForOutBound'")();
    this.postUrlForRecall = laUrl[0].PostUrl;
    this.triggerProjectRecall();
  }
  protected async triggerProjectRecall() {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrlForRecall;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'TransmittalNo': this.transmittalID,
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
  }
  public _onPreviewBtnClick() {
    let totalFiles;
    totalFiles = add(this.state.itemsForGrid.length, this.state.itemsForExternalGrid.length);
    this.setState({
      totalNoOfFiles: totalFiles.toString(),
      previewDiv: false,
      showReviewModal: true,
    });
  }
  public _uploadadditional(e: React.ChangeEvent<HTMLInputElement>) {
    this.myfileadditional = e.target.value;
    let documentNameExtension;
    console.log(this.myfileadditional);
    console.log(e.target.value);
    console.log(e.currentTarget.value);
    let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
    var splitted = myfile.name.split(".");
    console.log(splitted);
    console.log(splitted.length);
    console.log(splitted[splitted.length - 1]);
    for (let r = 0; r < splitted.length - 1; r++) {
      documentNameExtension = splitted.slice(0, -1).join('.') + "_TR00011" + '.' + splitted[splitted.length - 1];
    }
    // documentNameExtension = splitted[0] + "_TR00011" + '.' + splitted[splitted.length - 1];
    console.log(documentNameExtension);
    let docName = documentNameExtension;
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
      if ((document.querySelector("#newfile") as HTMLInputElement).files[0] != null) {
        let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
        if (myfile.size) {
          this.state.tempArrayForExternalDocumentGrid.push({
            documentName: myfile.name,
            fileSize: (((myfile.size / 1024)).toFixed(2)),
            fileSizeInMB: (((myfile.size / 1024) * 0.0009765625).toFixed(2)),
            externalComments: this.state.externalComments,
            content: myfile,
          });
          this.setState({
            showExternalGrid: false,
            fileSizeDiv: false,
            itemsForExternalGrid: this.state.tempArrayForExternalDocumentGrid,
          });
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
  private onCommentExternalChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    const newMultiline = newText.length > 50;
    if (newMultiline !== this.state.toggleMultiline) {
      this.setState({
        toggleMultiline: true,
      });
    }
    this.setState({ externalComments: newText || '' });
  }
  protected async triggerOutboundTransmittal(transmittalID: number) {
    let siteUrl = window.location.protocol + "//" + window.location.hostname + this.props.siteUrl;
    const postURL = this.postUrl;
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    const body: string = JSON.stringify({
      'SiteURL': siteUrl,
      'TransmittalNo': transmittalID,
      'ProjectName': this.state.projectName,
      'ContractNumber': this.state.contractNumber,
      'ProjectNumber': this.state.projectNumber,
      'CoverLetterNeeded': (this.state.coverLetterNeeded === true ? "Yes" : "NO"),
      'InternalContactsEmails': this.state.internalContactsEmail,
      'InternalContactsDisplayNames': this.state.internalCCContactsDisplayNameForPreview,
      'OutboundTransmittalDetails': this.state.itemsForGrid
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
      // alert(response.text);
      if (responseJSON['Status'] === "MailSend") {
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
  //incrementing transmittal id sequence number
  private _transmittalSequenceNumber(incrementValue: any, sequenceNumber: number) {
    var incrementSequenceNumber = incrementValue;
    while (incrementSequenceNumber.length < sequenceNumber)
      incrementSequenceNumber = "0" + incrementSequenceNumber;
    console.log(incrementSequenceNumber);
    this.setState({
      incrementSequenceNumber: incrementSequenceNumber,
    });
  }
}
