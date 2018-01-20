/** Fanric */
import { NormalPeoplePicker, TagPicker, ITag } from "office-ui-fabric-react/lib/Pickers";
import { PrimaryButton, ButtonType, Button, DefaultButton, ActionButton } from "office-ui-fabric-react/lib/Button";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Checkbox } from "office-ui-fabric-react/lib/Checkbox";
import { Label } from "office-ui-fabric-react/lib/Label";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from "office-ui-fabric-react/lib/DetailsList";
import { DatePicker, } from "office-ui-fabric-react/lib/DatePicker";
import { IPersonaProps } from "office-ui-fabric-react/lib/Persona";
/** SPFX Stuff */
import { SPComponentLoader } from "@microsoft/sp-loader";
import * as React from "react";

/** Other utilities */
import * as moment from "moment";
import { find, clone, remove, filter, map, orderBy, countBy, findIndex, startsWith } from "lodash";
import Dropzone from 'react-dropzone';
// switch to fabric pivot on text update
//import * as tabs from "react-tabs";
import { Pivot, PivotItem, PivotLinkFormat, PivotLinkSize } from "office-ui-fabric-react/lib/Pivot";

/**  Custom Stuff */
import { DocumentIframe } from "./DocumentIframe";
import { TRDocument, TR, modes, Pigment, Test, PropertyTest, DisplayPropertyTest, Customer } from "../dataModel";
import { ITrFormProps } from "./ITrFormProps";
import * as md from "./MessageDisplay";
import MessageDisplay from "./MessageDisplay";
import TRPicker from "./TRPicker";
import { ITRFormState } from "./ITRFormState";
import styles from "./TrForm.module.scss";

/**
 * Renders the new and edit form for technical requests
 * 
 * @export
 * @class TrForm
 * @extends {React.Component<ITrFormProps, ITRFormState>}
 */
export default class TrForm extends React.Component<ITrFormProps, ITRFormState> {
  private ckeditor: any;
  private originalAssignees: Array<number> = [];
  private originalStatus: string = "";
  private originalRequiredDate: string = "";
  private validBrandIcons = " accdb csv docx dotx mpp mpt odp ods odt one onepkg onetoc potx ppsx pptx pub vsdx vssx vstx xls xlsx xltx xsn ";


  constructor(props: ITrFormProps) {
    super(props);
    this.state = props.initialState;
    this.originalAssignees = clone(this.state.tr.TRAssignedToId);// sasve original so we can email new assignees
    this.originalStatus = this.state.tr.TRStatus;// sasve original so we can email if it gets closed
    this.originalRequiredDate = this.state.tr.RequiredDate;// sasve original so we can email if it gets closed
    this.SaveButton = this.SaveButton.bind(this);
    this.save = this.save.bind(this);
    this.cancel = this.cancel.bind(this);

    this.selectChildTR = this.selectChildTR.bind(this);
    this.cancelTrSearch = this.cancelTrSearch.bind(this);
    this.parentTRSelected = this.parentTRSelected.bind(this);
    this.editParentTR = this.editParentTR.bind(this);
    this.uploadFile = this.uploadFile.bind(this);
    this.resolveCustomerSuggestions = this.resolveCustomerSuggestions.bind(this);

  }


  /**
   * The ckeditor is not a react component so it is handled outside of the react lifecycle.
   * After the component mounts. We need toload ck editoe and replace wjat is the first tan with a ck-editor control.
   * The first tab is the Title tab.  So that is the one that is initially displayed and needs the ckeditor.
   * Later as the tabs changes we will add a ckeditor under the other tabks
   * 
   * 
   * @memberof TrForm
   */
  public componentDidMount() {
    // see https://github.com/SharePoint/sp-de//cdn.ckeditor.com/4.6.2/full/ckeditor.jsv-docs/issues/374
    // var ckEditorCdn: string = "//cdn.ckeditor.com/4.6.2/full/ckeditor.js";
    var ckEditorCdn: string = this.props.ckeditorUrl;
    SPComponentLoader.loadScript(ckEditorCdn, { globalExportsName: "CKEDITOR" }).then((CKEDITOR: any): void => {
      this.ckeditor = CKEDITOR;
      // replaces the title with a ckeditor. the other textareas are not visible yet. They will be replaced when the tab becomes active
      this.ckeditor.replace("tronoxtrtextarea-title", this.props.ckeditorConfig);

    });

  }

  /**
   * When the user changes the tabs, if the new tab contains a ckeditor control, replcae the textarea that was rendered by default 
   * with the ckeditor control. We need a breief delay before doing so so that we can be sure the control has been rendered
   * @param {any} newTabID 
   * @param {any} oldTabID 
   * 
   * @memberof TrForm
   */
  public tabChanged(item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) {
    switch (item.props.linkText) {
      case "Title":
        if (this.ckeditor.instances["tronoxtrtextarea-title"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-title", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-title");
          });
        }
        break;
      case "Description":
        if (this.ckeditor.instances["tronoxtrtextarea-description"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-description", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-description");
          });
        }
        break;
      case "Summary":

        if (this.ckeditor.instances["tronoxtrtextarea-summary"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-summary", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-summary");
          });
        }
        break;
      case "Test Params":
        if (this.ckeditor.instances["tronoxtrtextarea-testparams"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-testparams", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-testparams");
          });
        }
        break;
      case "Formulae":
        if (this.ckeditor.instances["tronoxtrtextarea-formulae"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-formulae", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-formulae");
          });
        }
        break;
      default:

    }



  }

  /**
   * Determines if the values entered in the TR are valid
   * 
   * @returns {boolean} 
   * 
   * @memberof TrForm
   */
  public getErrors(): md.Message[] {

    let errorMessages: md.Message[] = [];
    if (!this.state.tr.Title) {
      errorMessages.push(new md.Message("Request #  is required"));
    }
    if (!this.state.tr.WorkTypeId) {
      errorMessages.push(new md.Message("Work Type is required"));
    }
    if (!this.state.tr.ApplicationTypeId) {
      errorMessages.push(new md.Message("Application Type is required"));
    }
    if (!this.state.tr.RequestDate) {
      errorMessages.push(new md.Message("Initiation Date   is required"));
    }
    if (!this.state.tr.RequiredDate) {
      errorMessages.push(new md.Message("Due Date  is required"));
    }
    if (this.state.tr.RequestDate && this.state.tr.ActualStartDate &&
      moment(this.state.tr.ActualStartDate).format("YYYYMMDD") < moment(this.state.tr.RequestDate).format("YYYYMMDD")) {
      errorMessages.push(new md.Message("Actual Start Date must be on or after Initiation Date"));
    }
    if (this.state.tr.ActualCompletionDate && this.state.tr.ActualStartDate &&
      moment(this.state.tr.ActualStartDate).format("YYYYMMDD") > moment(this.state.tr.ActualCompletionDate).format("YYYYMMDD")) {
      errorMessages.push(new md.Message("Actual Completion Date must be on or after Actual Start Date"));
    }
    if (this.state.tr.TRStatus === "Completed" && !this.state.tr.ActualStartDate) {
      errorMessages.push(new md.Message("Actual Start Date is required to complete a request"));
    }
    if (this.state.tr.TRStatus === "Completed" && !this.state.tr.ActualCompletionDate) {
      errorMessages.push(new md.Message("Actual Completion Date is required to complete a request"));
    }
    if (this.state.tr.TRStatus === "Completed" && !this.state.tr.ActualManHours) {
      errorMessages.push(new md.Message("Actual Hours is required to complete a request"));
    }
    debugger;
    if (this.state.tr.RequiredDate && this.state.tr.RequestDate && this.state.tr.RequestDate > this.state.tr.RequiredDate) {
      errorMessages.push(new md.Message("Due Date  must be after Initiation Date"));
    }
    if (!this.state.tr.Site) {
      errorMessages.push(new md.Message("Site is required"));
    }
    if (!this.state.tr.TRPriority) {
      errorMessages.push(new md.Message("Proiority  is required"));
    }
    if (!this.state.tr.TRStatus) {
      errorMessages.push(new md.Message("Status  is required"));
    }
    // if (!this.state.tr.RequestorId) {
    //   errorMessages.push(new md.Message("Requestor is required"));
    // }
    return errorMessages;
  }

  /**
   * Saves the TR Back to sharepoint
   * Gets the data out of the ckeditor controls and adds it to the TR, validates all the fields on the TR, and then calls the
   * save method on the parent webpart.
   * @returns 
   * 
   * @memberof TrForm
   */
  public save() {

    let tr: TR = { ...this.state.tr };
    // tslint:disable-next-line:forin
    for (const instanceName in this.ckeditor.instances) {

      let instance = this.ckeditor.instances[instanceName];
      let data = instance.getData();
      switch (instanceName) {
        case "tronoxtrtextarea-title":
          tr.RequestTitle = data;
          break;
        case "tronoxtrtextarea-description":
          tr.Description = data;
          break;
        case "tronoxtrtextarea-summary":
          tr.SummaryNew = data; // summaryNew gets appended to summary when we save
          break;
        case "tronoxtrtextarea-testparams":
          tr.TestingParameters = data;
          break;
        case "tronoxtrtextarea-formulae":
          tr.Formulae = data;
          break;
        default:
          console.log("Text area missing in save " + instanceName);

      }
    }
    const errors: md.Message[] = this.getErrors();
    if (errors.length === 0) {
      tr.FileCount=this.state.documents.length; // update the file count to be however many files are here
      tr.ActualManHours = this.props.hoursSpent; // updaye hours spent, Thi # we put in the prop, is the total accumulated so fat in the time spent list
      this.props.save(tr, this.originalAssignees, this.originalStatus, this.originalRequiredDate)
        .then((result: TR) => {
          tr.Id = result.Id;
          this.setState((current) => ({ ...current, errorMessages: [], isDirty: false, tr: tr }));
        })
        .catch((response) => {
          let errormessages = this.state.errorMessages;
          errormessages.push(new md.Message(response.data.responseBody["odata.error"].message.value));
          this.setState((current) => ({ ...current, errorMessages: errormessages, tr: tr }));
        });
    } else {
      this.setState((current) => ({ ...current, errorMessages: errors, tr: tr })); // show errors
    }
    return false; // stop postback
  }

  /**
   * Cancels editing the TR. Returs to the previos page (done in parent webpart.)
   * 
   * @returns 
   * 
   * @memberof TrForm
   */
  public cancel() {
    this.props.cancel();
    return false; // stop postback
  }

  /**
   * Updates the text in all the existingck editors after we loaded a new TR (parent or child)
   * 
   * @param {TR} tr 
   * 
   * @memberof TrForm
   */
  public updateCKEditorText(tr: TR) {
    for (const instanceName of this.ckeditor.instances) {
      let instance: any = this.ckeditor.instances[instanceName];
      switch (instanceName) {
        case "tronoxtrtextarea-title":
          instance.setData(tr.RequestTitle);
          break;
        case "tronoxtrtextarea-description":
          instance.setData(tr.Description);
          break;
        case "tronoxtrtextarea-summary":
          instance.setData(tr.SummaryNew);
          break;
        case "tronoxtrtextarea-formulae":
          instance.setData(tr.Formulae);
          break;
        default:

      }
    }
  }

  /**
   * Removes a message from the Message List . (called when the user clicks to remove a message)
   * 
   * @param {Array<md.Message>} messageList The message list to remove the message from
   * @param {string} messageId The message to remove
   * 
   * @memberof TrForm
   */
  public removeMessage(messageList: Array<md.Message>, messageId: string) {
    remove(messageList, {
      Id: messageId
    });
    this.setState((current) => ({ ...current }));
  }

  /**
   * Renders the save buton with appropriate handler
   * 
   * @returns {JSX.Element} 
   * 
   * @memberof TrForm
   */
  public SaveButton(): JSX.Element {
    if (this.props.mode === modes.DISPLAY) {
      return <div />;
    } else {
      return (
        <PrimaryButton buttonType={ButtonType.primary} onClick={this.save} icon="ms-Icon--Save">
          <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
          Save
      </PrimaryButton>

      );
    }
  }

  /**
   * Determines if the TR contains the selected Pigment
   * 
   * @param {TR} tr  The TR to check.
   * @param {number} PigmentId  the ID of the pigment to check.
   * @returns {boolean} 
   * 
   * @memberof TrForm
   */
  public trContainsPigment(tr: TR, PigmentId: number): boolean {
    if (tr.PigmentsId) {
      return (tr.PigmentsId.indexOf(PigmentId) != -1);
    } else {
      return false;
    }
  }
  /**
   * Determines if the TR contains the selected Test
   * 
   * @param {TR} tr  The TR to check.
   * @param {number} TestId  the ID of the Test to check.
   * @returns {boolean} 
   * 
   * @memberof TrForm
   */
  public trContainsTest(tr: TR, TestId: number): boolean {
    if (tr.TestsId) {
      return (tr.TestsId.indexOf(TestId) !== -1);
    } else {
      return false;
    }
  }

  /**
   * Return pigments Not on the Tr being edited and are Active. Theses are the pigments that can be selected./
   * 
   * @returns {Array<Pigment>} 
   * 
   * @memberof TrForm
   */
  public getAvailablePigments(): Array<Pigment> {
    var tempPigments = filter(this.props.pigments, (pigment: Pigment) => {
      return !this.trContainsPigment(this.state.tr, pigment.id);
    });
    var pigments = map(tempPigments, (pigment) => {
      return {
        title: pigment.title,
        manufacturer: (pigment.manufacturer) ? pigment.manufacturer : "(none)",
        id: pigment.id,
        isActive: pigment.isActive
      };
    }).filter((p) => { return p.isActive === "Yes"; });
    return orderBy(pigments, ["manufacturer"], ["asc"]);
  }


  /*
 * Gets the Groups used to display the available pigments.
 * (See https://dev.office.com/fabric#/components/groupedlist)
 * The group contains the starting index of the group and the number of elements
 * 
 * @returns {Array<IGroup>} 
 * 
 * @memberof TrForm
 */
  public getAvailablePigmentGroups(): Array<IGroup> {
    var pigs: Array<Pigment> = this.getAvailablePigments();
    var pigmentManufactureres = countBy(pigs, (p1: Pigment) => { return p1.manufacturer; });
    var groups: Array<IGroup> = [];

    // tslint:disable-next-line:forin
    for (const pm in pigmentManufactureres) {
      groups.push({
        isCollapsed: (pm !== this.state.expandedPigmentManufacturer),
        name: pm,
        key: pm,
        startIndex: findIndex(pigs, (pig) => { return pig.manufacturer === pm; }),
        count: pigmentManufactureres[pm],

      });
    }
    return groups;
  }

  /**
   *  Gets all the Pigments on the TR being edited
   * 
   * @returns {Array<Pigment>} 
   * 
   * @memberof TrForm
   */
  public getSelectedPigments(): Array<Pigment> {

    var tempPigments = filter(this.props.pigments, (pigment: Pigment) => {
      return this.trContainsPigment(this.state.tr, pigment.id);
    });
    var selectedPigments = map(tempPigments, (pigment) => {
      return {
        title: pigment.title,
        manufacturer: (pigment.manufacturer) ? pigment.manufacturer : "(none)",
        id: pigment.id,
        isActive: pigment.isActive
      };
    });
    return orderBy(selectedPigments, ["title"], ["asc"]);
  }



  /**
   * return tests available from the PropetyTest table and not on the tr. These are the tests which can be added to the TR
   * 
   * @returns {Array<DisplayPropertyTest>} 
   * 
   * @memberof TrForm
   */
  public getAvailableTests(): Array<DisplayPropertyTest> {
    // select the propertyTest available based on the applicationType and EndUse of the tr
    debugger;
    var temppropertyTest: Array<PropertyTest> = filter(this.props.propertyTests, (pt: PropertyTest) => {
      return (pt.applicationTypeid === this.state.tr.ApplicationTypeId
        && pt.endUseIds.indexOf(this.state.tr.EndUseId) !== -1
      );
    });
    // get all the the tests in those propertyTests and output them as DisplayPropertyTest
    var tempDisplayTests: Array<DisplayPropertyTest> = [];
    for (const pt of temppropertyTest) {
      for (const testid of pt.testIds) {
        const test: Test = find(this.props.tests, (t) => { return t.id === testid; });
        if (test) {
          tempDisplayTests.push({
            property: pt.property,
            testid: testid,
            test: test.title
          });
        } else {
          console.log(` test ${testid} NOT FOUND`);
        }

      }
    }
    // now remove those that are already on the tr
    const displayTests = filter(tempDisplayTests, (dt) => { return !this.trContainsTest(this.state.tr, dt.testid); });

    return orderBy(displayTests, ["property", "test"], ["asc", "asc"]);


  }

  /**
   * Gets the Groups used to display the available tests.
   * (See https://dev.office.com/fabric#/components/groupedlist)
   * The group contains the starting index of the group and the number of elements
   * 
   * @returns {Array<IGroup>} 
   * 
   * @memberof TrForm
   */
  public getAvailableTestGroups(): Array<IGroup> {
    debugger;
    var displayPropertyTests: Array<DisplayPropertyTest> = this.getAvailableTests(); // all the avalable tests with their Property
    var properties = countBy(displayPropertyTests, (dpt: DisplayPropertyTest) => {
      return dpt.property;
    });// an object with an element for each propert, the value of the elemnt is the count of tsts with that property
    var groups: Array<IGroup> = [];
    // tslint:disable-next-line:forin
    for (const property in properties) {
      groups.push({
        isCollapsed: (property !== this.state.expandedProperty),
        name: property,
        key: property,
        startIndex: findIndex(displayPropertyTests, (dpt) => { return dpt.property === property; }),
        count: properties[property],

      });
    }
    return groups;
  }
  public getPropertyName(testId: number, applicationTypeid: number, endUseId: number): string {
    /**
     * A property can be valid for many enduses, application  types amd tests
     * Find the proprty for the  selected enduse, application  type amd test
     */

    var property: PropertyTest = find(this.props.propertyTests, (pt) => {
      return (
        pt.applicationTypeid === applicationTypeid
        &&
        pt.testIds.indexOf(testId) !== -1
        &&
        pt.endUseIds.indexOf(endUseId) !== -1

      );
    });
    return (property) ? property.property : "";
  }
  /**
   * return Tests on the tr being edited
   * 
   * @returns {Array<Test>} 
   * 
   * @memberof TrForm
   */
  public getSelectedTests(): Array<Test> {
    var tempTests = filter(this.props.tests, (test: Test) => {
      return this.trContainsTest(this.state.tr, test.id);
    });
    var selectedTests = map(tempTests, (test) => {
      return {
        title: test.title,
        id: test.id,
        propertyName: this.getPropertyName(test.id, this.state.tr.ApplicationTypeId, this.state.tr.EndUseId)
      };
    });
    return orderBy(selectedTests, ["propertyName", "title"], ["asc", "asc"]);
  }
  /**
   * Gets the technical Specialists, including an indicator if they are selected or not and if its the Primary Tech Spec
   * 
   * @returns 
   * 
   * @memberof TrForm
   */
  public getTechSpecs() {
    var techSpecs = map(this.props.techSpecs, (techSpec) => {
      return {
        title: techSpec.title,
        selected: ((this.state.tr.TRAssignedToId) ? this.state.tr.TRAssignedToId.indexOf(techSpec.id) != -1 : null),
        id: techSpec.id,
        primary: (techSpec.id === this.state.tr.TRPrimaryAssignedToId)
      };
    });
    return orderBy(techSpecs, ["selected", "title"], ["desc", "asc"]);
  }

  /**
   * Adds or removes a preson from the TechnicalSpecialts on the TR being edited
   * 
   * @param {boolean} isSelected 
   * @param {number} id 
   * 
   * @memberof TrForm
   */
  public toggleTechSpec(isSelected: boolean, id: number) {
    if (isSelected) {
      if (this.state.tr.TRAssignedToId) {
        this.state.tr.TRAssignedToId.push(id);// addit
      } else {
        this.state.tr.TRAssignedToId = [id];
      }
    } else {
      // only remove if its not the primary
      if (this.state.tr.TRPrimaryAssignedToId !== id) {
        this.state.tr.TRAssignedToId = filter(this.state.tr.TRAssignedToId, (x) => { return x !== id; });// remove it
      }
    }
    this.setState((current) => ({ ...current, isDirty: true }));
  }
  /**
   * Flags the selected techSpec as the PrimaryTechspec. A TR can have many tech specs . Only one primary.
   * 
   * @param {boolean} isPrimary 
   * @param {number} id 
   * 
   * @memberof TrForm
   */
  public togglePrimaryTechSpec(isPrimary: boolean, id: number) {
    if (isPrimary) {
      this.state.tr.TRPrimaryAssignedToId = id;// addit
    }
    if (!(this.state.tr.TRAssignedToId) || this.state.tr.TRAssignedToId.indexOf(id) === -1) { // if i dont have tech specs yet, or this one is not selected
      this.toggleTechSpec(true, id);
    }
    this.setState((current) => ({ ...current, isDirty: true }));
  }


  public staffCCChanged(items?: Array<IPersonaProps>): void {
    this.state.tr.StaffCC = items;
    this.props.ensureUsersInPersonas(this.state.tr.StaffCC);
  }

  /******** TEST Toggles , this is two lists, toggling adds from one , removes from the other*/
  public addTest(id: number): void {
    let tr: TR = this.state.tr;
    if (tr.TestsId) {
      tr.TestsId.push(id);// addit
    } else {
      tr.TestsId = [id];
    }
    this.setState((current) => ({ ...current, isDirty: true, tr: tr }));
  }
  public removeTest(id: number) {
    //   this.setDirty(true);
    if (this.state.tr.TestsId) {
      this.state.tr.TestsId = filter(this.state.tr.TestsId, (x) => { return x !== id; });// remove it
    }
    this.setState((current) => ({ ...current, isDirty: true }));
  }

  /**
   * Adds the selected Pigment to the TR being edited
   * In the UI Pigmemts is two lists (selected and unselected pigments, toggling adds from one , removes from the other
   * 
   * @param {number} id The ID if the pigment to add or remove
   * 
   * @memberof TrForm
   */
  public addPigment(id: number) {
    //  this.setDirty(true);
    if (this.state.tr.PigmentsId) {
      this.state.tr.PigmentsId.push(id);// add it
    } else {
      this.state.tr.PigmentsId = [id];
    }
    //  this.setState(this.state);
    this.setState((current) => ({ ...current, isDirty: true }));
  }
  /*
  * Removes  the selected Pigment to the TR being edited
  * In the UI Pigmemts is two lists (selected and unselected pigments, toggling adds from one , removes from the other
  * 
  * @param {number} id The ID if the pigment to add or remove
  * 
  * @memberof TrForm
  */
  public removePigment(id: number) {
    // this.setDirty(true);
    if (this.state.tr.PigmentsId) {
      this.state.tr.PigmentsId = filter(this.state.tr.PigmentsId, (x) => { return x !== id; });// remove it
    }
    // this.setState(this.state);
    this.setState((current) => ({ ...current, isDirty: true }));
  }




  /**
   * Renders a formatted date in the UI
   * 
   * @param {*} [item]  The item the field resides in
   * @param {number} [index] The index of the item in the list of items
   * @param {IColumn} [column] The column that contains the date to dispplay
   * @returns {*} 
   * 
   * @memberof TrForm
   */
  public renderDate(item?: any, index?: number, column?: IColumn): any {

    return moment(item[column.fieldName]).format("MMM Do YYYY");
  }

  /**
   * Hides the TR Search modal(TRPicker.tsx).
   * The seatch modal is used to select a parent tr
   * 
   * @memberof TrForm
   */
  public cancelTrSearch(): void {
    this.setState((current) => ({ ...current, showTRSearch: false }));
  }

  /**
   * Makes the selected Child TR the current TR. Called when the user clicks the edit icon in the child TR List.
   * 
   * @param {number} trId The ID of the Child TR to edit.
   * @returns {*} 
   * 
   * @memberof TrForm
   */
  public selectChildTR(trId: number): any {
    const childTr = find(this.state.childTRs, (tr) => { return tr.Id === trId; });

    if (childTr) {
      console.log("switching to tr " + trId);
      // delete this.state.tr;
      // this.state.tr = childTr;
      this.originalAssignees = clone(childTr.TRAssignedToId);
      this.originalStatus = childTr.TRStatus;
      this.originalRequiredDate = childTr.RequiredDate;

      this.updateCKEditorText(this.state.tr);
      // this.state.childTRs = [];
      this.setState((current) => ({ ...current, tr: childTr, childTRs: [] }));
      // this.setState(this.state);
      // now get its children, need to move children to state
      this.props.fetchChildTr(this.state.tr.Id).then((trs) => {
        // this.state.childTRs = trs;
        this.setState((current) => ({ ...current, childTRs: trs }));
      });
    }
    return false;
  }

  /**
   * Opens the selected TR Document in a new window.
   * 
   * @param {TRDocument} trdocument 
   * 
   * @memberof TrForm
   */
  public editDocument(trdocument: TRDocument): void {

    // mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(trdocument.id, 1).then(url => {

      if (!url || url === "") {
        window.open(trdocument.serverRalativeUrl, "_blank");
      } else {
        window.open(url, "_blank");
      }
      //    this.state.wopiFrameUrl=url;
      //  this.setState(this.state);
      // window.location.href=url;

    });

  }
   /**
   * Deletes the selectd file from the TR Document library.
   * 
   * @param {TRDocument} trdocument 
   * 
   * @memberof TrForm
   */
  public deleteFile(trdocument: TRDocument): void {

    debugger;
    this.props.deleteFile(trdocument.id).then((results)=>{
      let remainingdocs=filter(this.state.documents,(doc)=>{
        return (doc.id !== trdocument.id);
      });
      this.setState((current) => ({ ...current, documents: remainingdocs }));
    }).catch(error=>{
      console.error(error);
      alert("there was an error deleting tis file");
      
    });
    
  }

  public parentTRSelected(id: number, title: string) {
    // this.state.tr.ParentTR = title;
    // this.state.tr.ParentTRId = id;
    // this.setDirty(true);
    const tr: TR = { ...this.state.tr, ParentTR: title, ParentTRId: id };
    this.setState((current) => ({ ...current, isDirty: true, tr: tr }));
    this.cancelTrSearch();
  }
  public uploadFile(e: any) {

    let file: any = e.target["files"][0];
    this.props.uploadFile(file, this.state.tr.Id,this.state.tr.Title).then((response) => {
      this.props.getDocuments(this.state.tr.Id).then((dox) => {
        // this.state.documents = dox;
        this.setState((current) => ({ ...current, documents: dox }));
      });

    }).catch((error) => {
      console.log("an error occurred uploading the file");
      console.log(error);
    });
  }
  public editParentTR(): void {

    if (this.state.tr.ParentTRId) {
      const parentId = this.state.tr.ParentTRId;
      this.props.fetchTR(parentId).then((parentTR) => {

        // this.state.tr = parentTR;
        this.originalAssignees = clone(parentTR.TRAssignedToId);
        this.originalStatus = parentTR.TRStatus;
        this.originalRequiredDate = parentTR.RequiredDate;
        // this.state.childTRs = [];
        this.setState((current) => ({ ...current, tr: parentTR, childTRs: [] }));
        this.updateCKEditorText(this.state.tr);
        this.props.fetchChildTr(parentId).then((subTRs) => {
          // this.state.childTRs = subTRs;
          this.setState((current) => ({ ...current, childTRs: subTRs }));
        });
      });
    }


  }
  public documentRowMouseEnter(trdocument: TRDocument, e: any) {

    // mode passed to fetchDocumentWopiFrameURL: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(trdocument.id, 3).then(url => {
      if (!url || url === "") {
        url = trdocument.serverRalativeUrl;
      }
      // this.state.documentCalloutIframeUrl = url;
      // this.state.documentCalloutTarget = e.target;
      // this.state.documentCalloutVisible = true;
      this.setState((current) => ({
        ...current,
        documentCalloutIframeUrl: url,
        documentCalloutTarget: e.target,
        documentCalloutVisible: true
      }));

    });
  }
  public documentRowMouseOut(item: TRDocument, e: any) {

    // this.state.documentCalloutTarget = null;
    // this.state.documentCalloutVisible = false;
    this.setState((current) => ({ ...current, documentCalloutTarget: null, documentCalloutVisible: false }));
    console.log("mouse exit for " + item.title);
  }

  public createSummaryMarkup(tr: TR) {
    return { __html: tr.Summary };
  }

  public renderDocumentTitle(item?: any, index?: number, column?: IColumn): any {
    let extension = item.title.split('.').pop();
    let classname = "";
    if (this.validBrandIcons.indexOf(" " + extension + " ") !== -1) {
      classname += " ms-Icon ms-BrandIcon--" + extension + " ms-BrandIcon--icon16 ";
    }
    else {
      //classname += " ms-Icon ms-Icon--TextDocument " + styles.themecolor;
      classname += " ms-Icon ms-Icon--TextDocument ";
    }


    return (
      <div>
        <div className={classname} /> &nbsp;
        <a href="#"
          onClickCapture={(e) => {

            e.preventDefault();
            this.editDocument(item); return false;
          }}><span className={styles.documentTitle} > {item.fileName}</span></a>
      </div>);
  }

  private resolveCustomerSuggestions(filterString: string, slectedItems?: ITag[]): ITag[] {

    const upperfilter = filterString.toUpperCase();
    const matches: Customer[] = filter(this.props.customers, (c) => { return startsWith(c.title.toUpperCase(), upperfilter); });
    const results: ITag[] = map(matches, (c) => {
      return { key: c.id.toString(), name: c.title };
    });
    return results;

  }
  // public setDirty(isDirty: boolean) {

  //   if (!this.state.isDirty && isDirty) { //wasnt dirty now it is 
  //     window.onbeforeunload = function (e) {
  //       var dialogText = "You have unsaved changes, are you sure you want to leave?";
  //       e.returnValue = dialogText;
  //       return dialogText;
  //     };
  //   }
  //   if (this.state.isDirty && !isDirty) { //was dirty now it is not
  //     window.onbeforeunload = null;
  //   }
  //   //this.state.isDirty = isDirty;
  //   this.setState((this.state)=>{ ...this.state, isDirty: isDirty });
  // }
    /**
   * Called when a user drops files into the DropZone. It calls 
   * the uploadFile method on the props to upload the files to sharepoint and then adds them to state.
   * 
   * @private
   * @param {any} acceptedFiles 
   * @param {any} rejectedFiles 
   * @memberof EfrApp
   */
  private onDrop(acceptedFiles, rejectedFiles) {
    console.log("in onDrop");
    let promises: Array<Promise<any>> = [];
    acceptedFiles.forEach(file => {
      promises.push(this.props.uploadFile(file,this.state.tr.Id,this.state.tr.Title));
    });
    Promise.all(promises).then((x) => {
      this.props.getDocuments(this.state.tr.Id).then((dox) => {
        this.setState((current) => ({ ...current, documents: dox }));
      });

    });

  }
  public render(): React.ReactElement<ITrFormProps> {

    let worktypeDropDoownoptions = map(this.props.workTypes, (wt) => {
      return {
        key: wt.id, text: wt.workType
      };
    });

    let applicationtypeDropDoownoptions =
      filter(this.props.applicationTypes, (at) => {
        // show if its valid for the selected Worktype, OR if its already on the tr
        return (at.workTypeIds.indexOf(this.state.tr.WorkTypeId) !== -1
          || at.id === this.state.tr.ApplicationTypeId);

      })
        .map((at) => {
          return {
            key: at.id, text: at.applicationType
          };
        });


    let enduseDropDoownoptions =
      filter(this.props.endUses, (eu) => {
        // show if its valid for the selected ApplicationType, OR if its already on the tr
        return (eu.applicationTypeId === this.state.tr.ApplicationTypeId
          || eu.id === this.state.tr.EndUseId);
      })
        .map((eu) => {
          return {
            key: eu.id, text: eu.endUse
          };
        });
    let selectedCustomer: ITag[] = [];
    if (this.state.tr.CustomerId) {
      let cust: Customer = find(this.props.customers, (c) => { return c.id === this.state.tr.CustomerId; });
      selectedCustomer.push({
        key: cust.id.toString(),
        name: cust.title
      });
    }
    debugger;
    return (
      <div>

        <MessageDisplay messages={this.state.errorMessages}
          hideMessage={this.removeMessage.bind(this)} />
        <div style={{ float: "left" }} className={styles.secondarybuttonrwg}> <Label className={styles.secondarybuttonrwg}>MODE : {modes[this.props.mode]}</Label></div>
        <div style={{ float: "right" }}>  <Label>Status : {(this.state.isDirty) ? "Unsaved" : "Saved"}</Label></div>
        <div style={{ clear: "both" }}></div>
        <table>

          <tr>
            <td>
              <Label >Request #</Label>
            </td>
            <td>
              <TextField disabled={true} value={this.state.tr.Title} onChanged={e => {
                // this.setDirty(true);
                this.state.tr.Title = e;
                // this.setState(this.state);
                this.setState((current) => ({ ...current, isDirty: true }));
              }} />
            </td>
            <td>
              <Label >Work Type</Label>
            </td>
            <td>
              <Dropdown label=""
                selectedKey={this.state.tr.WorkTypeId}
                options={worktypeDropDoownoptions}
                onChanged={e => {
                  // this.setDirty(true);
                  this.state.tr.WorkTypeId = e.key as number;
                  //  this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }} />
            </td>

            <td>
              <Label >Site</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Site} onChanged={e => {
                //  this.setDirty(true);
                this.state.tr.Site = e;
                //  this.setState(this.state);
                this.setState((current) => ({ ...current, isDirty: true }));
              }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label  >Parent TR</Label>
            </td>
            <td>
              <div>
                <Label style={{ "display": "inline" }}>
                  {this.state.tr.ParentTR}
                </Label>
                <i onClick={this.editParentTR}
                  className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                <i onClick={(e) => {
                  // this.state.showTRSearch = true;
                  // this.setState({ ...this.state, showTRSearch: true });
                  this.setState((current) => ({ ...current, showTRSearch: true }));
                }}
                  className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>

              </div>

            </td>
            <td>
              <Label >Application Type</Label>
            </td>
            <td>
              <Dropdown label=""
                selectedKey={this.state.tr.ApplicationTypeId}
                options={applicationtypeDropDoownoptions}
                onChanged={e => {
                  //  this.setDirty(true);
                  this.state.tr.ApplicationTypeId = e.key as number;
                  // this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));

                }}
              />
            </td>
            <td>
              <Label >Priority</Label>
            </td>
            <td>
              <Dropdown
                label=""
                options={[
                  { key: "High", text: "High" },
                  { key: "Normal", text: "Normal" },
                  { key: "Routine", text: "Routine" },

                ]}
                onChanged={e => {
                  //  this.setDirty(true);
                  this.state.tr.TRPriority = e.text;
                  //  this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }}
                selectedKey={this.state.tr.TRPriority} />
            </td>

          </tr>
          <tr>
            <td>
              <Label>CER #</Label>
            </td>
            <td>
              <TextField value={this.state.tr.CER} onChanged={e => {
                // this.setDirty(true);
                this.state.tr.CER = e;
                this.setState((current) => ({ ...current, isDirty: true }));
              }} />
            </td>
            <td>
              <Label>Requestor</Label>
            </td>
            <td>
              {/* <Dropdown
                label=""
                options={this.props.requestors.map((r) => { return { key: r.id, text: r.title }; })}
                onChanged={e => {
                  this.setDirty(true);
                  this.state.tr.RequestorId = e.key as number;
                  this.setState(this.state);
                }}
                selectedKey={this.state.tr.RequestorId}
              /> */}
              <Label>{this.state.tr.RequestorName}</Label>
            </td>
            <td>
              <Label>Customer</Label>
            </td>
            <td>
              {/* <ComboBox options={customerSelectOptions}
                autoComplete="on"
                autoCorrect="on"
                value={find(this.props.customers, (c) => { return c.id === this.state.tr.CustomerId; }).title}
                onChanged={(newValue: IComboBoxOption) => {
                  const custId: number = newValue.key as number;
                  const tr: TR = { ...this.state.tr, CustomerId: custId };
                  //   this.setState({ ...this.state, tr: tr, isDirty: true });
                  this.setState((current) => ({ ...current, tr: tr, isDirty: true }));
                  // this.state.tr.CustomerId = newValue.key;
                  // this.setDirty(true);
                  // this.setState(this.state);
                }}
              /> */}
              {/* <Select
                simpleValue
                placeholder="+ Add a Customer"
                options={customerSelectOptions}
                value={this.state.tr.CustomerId}
                matchPos={"start"}
                onChange={(newValue) => {
                  this.state.tr.CustomerId = newValue;
                  this.setDirty(true);
                  this.setState(this.state);
                }}
                    onResolveSuggestions: (filter: string, selectedItems?: T[]) => T[] | PromiseLike<T[]>;
              /> */}
              <TagPicker
                onResolveSuggestions={this.resolveCustomerSuggestions}
                itemLimit={1}
                onChange={(newValue) => {
                  let custId: number = null;
                  if (newValue.length !== 0) {
                    custId = parseInt(newValue[0].key, 10);
                  }
                  const tr: TR = { ...this.state.tr, CustomerId: custId };
                  this.setState((current) => ({ ...current, tr: tr, isDirty: true }));
                }}
                defaultSelectedItems={selectedCustomer}
              />
              {/* <Dropdown
                label=""
                options={this.props.customers.map((r) => { return { key: r.id, text: r.title }; })}
                onChanged={e => {
                  this.setDirty(true);
                  this.state.tr.CustomerId = e.key as number;
                  this.setState(this.state);
                }}
                selectedKey={this.state.tr.CustomerId}
              /> */}

            </td>

          </tr>
          <tr>
            <td>
              <Label  >Initiation Date</Label>
            </td>
            <td>

              <DatePicker
                value={(this.state.tr.RequestDate) ? moment(this.state.tr.RequestDate).toDate() : null}

                onSelectDate={e => {
                  //                  this.setDirty(true);
                  this.state.tr.RequestDate = moment(e).toISOString();
                  //                this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }} />
            </td>
            <td>
              <Label >End Use</Label>
            </td>
            <td>
              <Dropdown label=""
                selectedKey={this.state.tr.EndUseId}
                options={enduseDropDoownoptions}
                onChanged={e => {
                  //     this.setDirty(true);
                  this.state.tr.EndUseId = e.key as number;
                  //    this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }} />
            </td>
            <td>
              <Label >Status</Label>
            </td>
            <td>
              <Dropdown
                label=""
                options={[
                  { key: "Pending", text: "Pending" },
                  { key: "In Progress", text: "In Progress" },
                  { key: "Completed", text: "Completed" },
                  { key: "Canceled", text: "Canceled" },
                ]}
                onChanged={e => {
                  //       this.setDirty(true);
                  this.state.tr.TRStatus = e.text;
                  //       this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }}
                selectedKey={this.state.tr.TRStatus} />

            </td>

          </tr>
          <tr>
            <td>
              <Label  >Due Date</Label>
            </td>
            <td>

              <DatePicker
                value={(this.state.tr.RequiredDate) ? moment(this.state.tr.RequiredDate).toDate() : null}
                onSelectDate={e => {
                  //    this.setDirty(true);
                  this.state.tr.RequiredDate = moment(e).toISOString();
                  //   this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }} />
            </td>
            <td>
              <Label >Actual Start Date</Label>
            </td>
            <td>
              <DatePicker
                value={(this.state.tr.ActualStartDate) ? moment(this.state.tr.ActualStartDate).toDate() : null}
                onSelectDate={e => {
                  //       this.setDirty(true);
                  this.state.tr.ActualStartDate = moment(e).toISOString();
                  //      this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }} />
            </td>
            <td>
              <Label >Actual Completion Date</Label>
            </td>
            <td>
              <DatePicker
                value={(this.state.tr.ActualCompletionDate) ? moment(this.state.tr.ActualCompletionDate).toDate() : null}
                onSelectDate={e => {
                  //      this.setDirty(true);
                  this.state.tr.ActualCompletionDate = moment(e).toISOString();
                  //       this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }} />
            </td>

          </tr>
          <tr>
            <td>
              <Label  >Estimated Hours</Label>
            </td>
            <td>

              <TextField datatype="number"
                value={(this.state.tr.EstManHours) ? this.state.tr.EstManHours.toString() : null}
                onChanged={e => {
                  //        this.setDirty(true);
                  this.state.tr.EstManHours = parseInt(e, 10);
                  //        this.setState(this.state);
                  this.setState((current) => ({ ...current, isDirty: true }));
                }} />
            </td>
            <td>
              <Label >Actual Hours</Label>
            </td>
            <TextField datatype="number"
              value={(this.state.tr.ActualManHours) ? this.state.tr.ActualManHours.toString() : null}
              onChanged={e => {
                //      this.setDirty(true);
                this.state.tr.ActualManHours = parseInt(e, 10);
                //    this.setState(this.state);
                this.setState((current) => ({ ...current, isDirty: true }));
              }} />


            <td>
              <Label>Accumulated Hours</Label>
            </td>
            <td>
              {this.props.hoursSpent}
            </td>

          </tr>


        </table>
        <Pivot onLinkClick={this.tabChanged.bind(this)}
          linkFormat={PivotLinkFormat.tabs}
          linkSize={PivotLinkSize.normal}>
          <PivotItem linkText='Title' onClick={(e) => { debugger; }}  >

            <textarea name="tronoxtrtextarea-title" id="tronoxtrtextarea-title" style={{ display: "none" }}>
              {this.state.tr.RequestTitle}
            </textarea>
          </PivotItem>
          <PivotItem linkText='Description' >
            <textarea name="tronoxtrtextarea-description" id="tronoxtrtextarea-description" style={{ display: "none" }}>
              {this.state.tr.Description}
            </textarea>
          </PivotItem>
          <PivotItem linkText='Summary' >
            <div dangerouslySetInnerHTML={this.createSummaryMarkup(this.state.tr)} />
            <textarea name="tronoxtrtextarea-summary" id="tronoxtrtextarea-summary" style={{ display: "none" }}>
              {this.state.tr.SummaryNew}
            </textarea>
          </PivotItem>
          <PivotItem linkText='Test Params' >
            <textarea name="tronoxtrtextarea-testparams" id="tronoxtrtextarea-testparams" style={{ display: "none" }}>
              {this.state.tr.TestingParameters}
            </textarea>
          </PivotItem>

          <PivotItem linkText={`Assigned To(${((this.state.tr.TRAssignedToId === null) ? "0" : this.state.tr.TRAssignedToId.length.toString())})`}>

            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.none}
              items={this.getTechSpecs()}
              setKey="id"
              columns={[
                { key: "title", name: "Technical Specialist", fieldName: "title", minWidth: 20, maxWidth: 200 },
                {
                  key: "selected", name: "Assigned?", fieldName: "selected", minWidth: 200, onRender: (item) =>
                    <Checkbox

                      checked={item.selected}
                      autoFocus={true}
                      onChange={(element, value) => {
                        this.toggleTechSpec(value, item.id);
                      }}

                    />
                },
                {
                  key: "primary", name: "Primary?", fieldName: "primary", minWidth: 200, onRender: (item) =>
                    <Checkbox

                      checked={item.primary}
                      autoFocus={true}
                      onChange={(element, value) => {
                        this.togglePrimaryTechSpec(value, item.id);
                      }}

                    />
                }
              ]}
            />
          </PivotItem>
          <PivotItem linkText={`Staff cc(${(this.state.tr.StaffCC === null) ? "0" : this.state.tr.StaffCC.length})`} >
            <NormalPeoplePicker
              defaultSelectedItems={this.state.tr.StaffCC}
              onChange={this.staffCCChanged.bind(this)}
              onResolveSuggestions={this.props.peopleSearch}
            >
            </NormalPeoplePicker>
          </PivotItem>
          <PivotItem linkText={`Pigments(${(this.state.tr.PigmentsId === null) ? "0" : this.state.tr.PigmentsId.length})`} >

            <div style={{ float: "left" }}>
              <Label> Available Pigments</Label>
              <DetailsList

                onDidUpdate={(dl: DetailsList) => {
                  // save expanded group in state;

                  var expandedGroup: IGroup = find(dl.props.groups, (group) => {
                    // its an expanded group that want expanded before
                    return !(group.isCollapsed) && group.key !== this.state.expandedPigmentManufacturer;
                  });
                  if (expandedGroup) {
                    // this.state.expandedPigmentManufacturer = expandedGroup.key;
                    this.setState((current) => ({ ...current, expandedPigmentManufacturer: expandedGroup.key }));
                  }
                }}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                groups={this.getAvailablePigmentGroups()}
                items={this.getAvailablePigments()}
                setKey="id"
                columns={[
                  { key: "title", name: "Pigment Name", fieldName: "title", minWidth: 20, maxWidth: 100 },
                  {
                    key: "select", name: "Select", fieldName: "selected", minWidth: 80, onRender: (item) =>
                      <Checkbox

                        checked={false}
                        onChange={(element, value) => {
                          this.addPigment(item.id);
                        }}
                      />
                  }
                ]}
              />

            </div>
            <div style={{ float: "right" }}>
              <Label> Selected Pigments</Label>
              <DetailsList

                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                items={this.getSelectedPigments()}
                setKey="id"
                columns={[
                  { key: "title", name: "Pigment Name", fieldName: "title", minWidth: 20, maxWidth: 100 },
                  { key: "manufacturer", name: "Manufacturer", fieldName: "manufacturer", minWidth: 20, maxWidth: 100 },
                  {
                    key: "select", name: "Select", fieldName: "selected", minWidth: 80, onRender: (item) =>
                      <Checkbox autoFocus={true}
                        checked={true}
                        onChange={(element, value) => {
                          this.removePigment(item.id);
                        }}
                      />
                  }
                ]}
              />
            </div>
            <div style={{ clear: "both" }}></div>
          </PivotItem>
          <PivotItem linkText={`Tests(${(this.state.tr.TestsId === null) ? "0" : this.state.tr.TestsId.length})`} >

            <div style={{ float: "left" }}>
              <Label> Available Tests</Label>
              <DetailsList
                onDidUpdate={(dl: DetailsList) => {
                  // save expanded group in state;

                  var expandedGroup = find(dl.props.groups, (group) => {
                    // its an expanded group that want expanded before
                    return !(group.isCollapsed) && group.key !== this.state.expandedProperty;
                  });
                  if (expandedGroup) {
                    // this.state.expandedProperty = expandedGroup.key;
                    this.setState((current) => ({ ...current, expandedProperty: expandedGroup.key }));
                  }
                }}
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                groups={this.getAvailableTestGroups()}
                items={this.getAvailableTests()}

                setKey="id"
                columns={[
                  { key: "title", name: "test", fieldName: "test", minWidth: 20, maxWidth: 150 },
                  {
                    key: "select", name: "Select", fieldName: "selected", minWidth: 70, onRender: (item) =>
                      <Checkbox
                        checked={false}
                        autoFocus={true}
                        onChange={(element, value) => { this.addTest(item.testid); }}
                      />
                  }
                ]}
              />

            </div>
            <div style={{ float: "right" }}>
              <Label> Selected Tests</Label>
              <DetailsList
                layoutMode={DetailsListLayoutMode.fixedColumns}
                selectionMode={SelectionMode.none}
                items={this.getSelectedTests()}
                setKey="id"
                columns={[
                  { key: "propertyName", name: "Property", fieldName: "propertyName", minWidth: 20, maxWidth: 80 },
                  { key: "title", name: "Test Name", fieldName: "title", minWidth: 20, maxWidth: 150 },
                  {
                    key: "selected", name: "Selected?", fieldName: "selected", minWidth: 70, onRender: (item) =>
                      <Checkbox
                        autoFocus={true}
                        checked={true}
                        onChange={(element, value) => { this.removeTest(item.id); }}
                      />
                  }
                ]}
              />

            </div>
            <div style={{ clear: "both" }}></div>



          </PivotItem>
          <PivotItem hidden={(this.state.tr.Id===null) ? true : false} linkText={`Documents(${(this.state.documents === null) ? "0" : this.state.documents.length})`}  >
          <Dropzone className={styles.dropzone} onDrop={this.onDrop.bind(this)} disableClick={true} >
        
            <div style={{ float: "left" }}>
              <DetailsList
                layoutMode={DetailsListLayoutMode.fixedColumns}
                items={this.state.documents}
                onRenderRow={(props, defaultRender) => <div
                  onMouseEnter={(event) => this.documentRowMouseEnter(props.item, event)}
                  onMouseOut={(evemt) => this.documentRowMouseOut(props.item, event)}>
                  {defaultRender(props)}
                </div>}
                setKey="id"
                selectionMode={SelectionMode.none}
                columns={[
                  { key: "title", name: "Request #", fieldName: "title", minWidth: 1,
                   maxWidth: 200, onRender: this.renderDocumentTitle.bind(this) },
                  {
                    key: "Delete", name: "", fieldName: "Delete", minWidth: 20,
                    onRender: (item) => <div>
                      <i onClick={(e) => { this.deleteFile(item); }}
                        className="ms-Icon ms-Icon--Delete" aria-hidden="true"></i>
                    </div>
                  },


                ]}
              />
              <input type="file" id="uploadfile" onChange={e => { this.uploadFile(e); }} />
            </div>
            <div style={{ float: "right" }}>
              <DocumentIframe src={this.state.documentCalloutIframeUrl} height={this.props.documentIframeHeight}
                width={this.props.documentIframeWidth} />
            </div>
            <div style={{ clear: "both" }}></div>
            </Dropzone>

          </PivotItem>
          <PivotItem linkText='Formulae' >

            <textarea name="tronoxtrtextarea-formulae" id="tronoxtrtextarea-formulae" style={{ display: "none" }}>
              {this.state.tr.Formulae}
            </textarea>
          </PivotItem>
          <PivotItem linkText={`Child TRs(${(this.state.childTRs === null) ? "0" : this.state.childTRs.length})`} >


            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              items={this.state.childTRs}
              setKey="id"
              selectionMode={SelectionMode.none}
              columns={[
                {
                  key: "Edit", name: "", fieldName: "Title", minWidth: 20,
                  onRender: (item?: any, index?: number, column?: IColumn) => <div>
                    <i onClick={(e) => {

                      this.selectChildTR(item.Id);
                    }}
                      className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                  </div>
                },
                { key: "Title", name: "Request #", fieldName: "Title", minWidth: 80, },
                { key: "Status", name: "Status", fieldName: "Status", minWidth: 90 },
                { key: "InitiationDate", onRender: this.renderDate, name: "Initiation Date", fieldName: "InitiationDate", minWidth: 80 },
                { key: "TRDueDate", onRender: this.renderDate, name: "Due Date", fieldName: "TRDueDate", minWidth: 80 },
                { key: "ActualStartDate", onRender: this.renderDate, name: "Actual Start Date", fieldName: "ActualStartDate", minWidth: 90 },
                { key: "ActualCompetionDate", onRender: this.renderDate, name: "Actual Competion<br />Date", fieldName: "ActualCompetionDate", minWidth: 80 },
              ]}
            />
          </PivotItem>
        

        </Pivot>

        <this.SaveButton />
        <span style={{ margin: 20 }}>
          <DefaultButton href="#" onClick={this.cancel} icon="ms-Icon--Cancel" className={styles.primarybutton} >
            <i className={`ms-Icon ms-Icon--Cancel ${styles.primarybutton}`} aria-hidden="true"></i>
            Cancel
        </DefaultButton>
        </span>
        <br />
        version 3
      < TRPicker
          isOpen={this.state.showTRSearch}
          callSearch={this.props.TRsearch}
          cancel={this.cancelTrSearch}
          select={this.parentTRSelected}
        />
      </div>
    );
  }
}
