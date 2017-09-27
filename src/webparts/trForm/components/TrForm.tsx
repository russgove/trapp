
import {
  NormalPeoplePicker, CompactPeoplePicker, IBasePickerSuggestionsProps,
} from 'office-ui-fabric-react/lib/Pickers';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { PrimaryButton, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType, } from 'office-ui-fabric-react/lib/MessageBar';
import { Dropdown, IDropdownProps, } from 'office-ui-fabric-react/lib/Dropdown';
// switch to fabric  ComboBox on next upgrade
let Select = require("react-select") as any;
import 'react-select/dist/react-select.css';
import { TagItem } from 'office-ui-fabric-react/lib/components/pickers/TagPicker/TagItem';
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from 'office-ui-fabric-react/lib/DetailsList';
import { DatePicker, } from 'office-ui-fabric-react/lib/DatePicker';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { SPComponentLoader } from '@microsoft/sp-loader';


/** SPFX Stuff */
import * as React from 'react';
//import styles from './TrForm.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

/** Other utilities */
import * as moment from 'moment';
import * as _ from "lodash";
import * as tabs from "react-tabs";

/**  Custom Stuff */
import { DocumentIframe } from "./DocumentIframe";
import { TRDocument, TR, modes, Pigment, Test, PropertyTest, DisplayPropertyTest, Customer } from "../dataModel";
import { ITrFormProps } from './ITrFormProps';
import * as md from "./MessageDisplay";
import MessageDisplay from "./MessageDisplay";
import TRPicker from "./TRPicker";
import { ITRFormState } from "./ITRFormState";


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
  private resultsPersonas: Array<IPersonaProps> = new Array<IPersonaProps>();

  constructor(props: ITrFormProps) {
    super(props);
    this.state = props.initialState;
    this.originalAssignees = _.clone(this.state.tr.TRAssignedToId);// sasve original so we can email new assignees
    this.originalStatus = this.state.tr.TRStatus;// sasve original so we can email if it gets closed
    this.SaveButton = this.SaveButton.bind(this);
    this.save = this.save.bind(this);
    this.cancel = this.cancel.bind(this);

    this.selectChildTR = this.selectChildTR.bind(this);
    this.cancelTrSearch = this.cancelTrSearch.bind(this);
    this.parentTRSelected = this.parentTRSelected.bind(this);
    this.editParentTR = this.editParentTR.bind(this);
    this.uploadFile = this.uploadFile.bind(this);

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
    //see https://github.com/SharePoint/sp-de//cdn.ckeditor.com/4.6.2/full/ckeditor.jsv-docs/issues/374
    //var ckEditorCdn: string = '//cdn.ckeditor.com/4.6.2/full/ckeditor.js';
    var ckEditorCdn: string = this.props.ckeditorUrl;
    SPComponentLoader.loadScript(ckEditorCdn, { globalExportsName: 'CKEDITOR' }).then((CKEDITOR: any): void => {
      this.ckeditor = CKEDITOR;
      this.ckeditor.replace("tronoxtrtextarea-title", this.props.ckeditorConfig); // replaces the title with a ckeditor. the other textareas are not visible yet. They will be replaced when the tab becomes active

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
  public tabChanged(newTabID, oldTabID) {


    switch (newTabID) {
      case 0:
        if (this.ckeditor.instances["tronoxtrtextarea-title"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-title", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-title");
          });
        }
        break;
      case 1:
        if (this.ckeditor.instances["tronoxtrtextarea-description"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-description", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-description");
          });
        }
        break;
      case 2:

        if (this.ckeditor.instances["tronoxtrtextarea-summary"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-summary", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-summary");
          });
        }
        break;
      case 3:
        if (this.ckeditor.instances["tronoxtrtextarea-testparams"] === undefined) {
          new Promise(resolve => setTimeout(resolve, this.props.delayPriorToSettingCKEditor)).then((xx) => {
            this.ckeditor.replace("tronoxtrtextarea-testparams", this.props.ckeditorConfig);
            console.log("created tronoxtrtextarea-testparams");
          });
        }
        break;
      case 8:
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
  public isValid(): boolean {

    this.state.errorMessages = [];
    let errorsFound = false;
    if (!this.state.tr.Title) {
      this.state.errorMessages.push(new md.Message("Request #  is required"));
      errorsFound = true;
    }
    if (!this.state.tr.WorkTypeId) {
      this.state.errorMessages.push(new md.Message("Work Type is required"));
      errorsFound = true;
    }
    if (!this.state.tr.ApplicationTypeId) {
      this.state.errorMessages.push(new md.Message("Application Type is required"));
      errorsFound = true;
    }
    if (!this.state.tr.RequestDate) {
      this.state.errorMessages.push(new md.Message("Initiation Date   is required"));
      errorsFound = true;
    }
    if (!this.state.tr.RequiredDate) {
      this.state.errorMessages.push(new md.Message("Due Date  is required"));
      errorsFound = true;
    }
    debugger;
    if (this.state.tr.RequestDate && this.state.tr.ActualStartDate && this.state.tr.ActualStartDate <  this.state.tr.RequestDate) {
      this.state.errorMessages.push(new md.Message("Actual Start Date must be on or after Initiation Date"));
      errorsFound = true;
    }
    if (this.state.tr.ActualCompletionDate && this.state.tr.ActualStartDate && this.state.tr.ActualStartDate >  this.state.tr.ActualCompletionDate) {
      this.state.errorMessages.push(new md.Message("Actual Completion Date must be on or after Actual Start Date"));
      errorsFound = true;
    }
    if (this.state.tr.TRStatus==="Completed" && ! this.state.tr.ActualStartDate ){
      this.state.errorMessages.push(new md.Message("Actual Start Date is required to complete a request"));
      errorsFound = true;
    }
    if (this.state.tr.TRStatus==="Completed" && ! this.state.tr.ActualCompletionDate ){
      this.state.errorMessages.push(new md.Message("Actual Completion Date is required to complete a request"));
      errorsFound = true;
    }
      
    if (this.state.tr.RequiredDate && this.state.tr.RequestDate && this.state.tr.RequestDate > this.state.tr.RequiredDate) {
      this.state.errorMessages.push(new md.Message("Due Date  must be after Initiation Date"));
      errorsFound = true;

    }
    if (!this.state.tr.Site) {
      this.state.errorMessages.push(new md.Message("Site is required"));
      errorsFound = true;
    }
    if (!this.state.tr.TRPriority) {
      this.state.errorMessages.push(new md.Message("Proiority  is required"));
      errorsFound = true;
    }
    if (!this.state.tr.TRStatus) {
      this.state.errorMessages.push(new md.Message("Status  is required"));
      errorsFound = true;
    }
    if (!this.state.tr.RequestorId) {
      this.state.errorMessages.push(new md.Message("Requestor is required"));
      errorsFound = true;
    }


    return !errorsFound;
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

    for (let instanceName in this.ckeditor.instances) {

      let instance = this.ckeditor.instances[instanceName];
      let data = instance.getData();
      switch (instanceName) {
        case "tronoxtrtextarea-title":
          this.state.tr.RequestTitle = data;
          break;
        case "tronoxtrtextarea-description":
          this.state.tr.Description = data;
          break;
        case "tronoxtrtextarea-summary":
          this.state.tr.SummaryNew = data; // summaryNew gets appended to summary when we save
          break;
        case "tronoxtrtextarea-testparams":
          this.state.tr.TestingParameters = data;
          break;
        case "tronoxtrtextarea-formulae":
          this.state.tr.Formulae = data;
          break;
        default:
          alert("Text area missing in save");

      }
    }
    if (this.isValid()) {
      this.props.save(this.state.tr, this.originalAssignees, this.originalStatus)
        .then((result: TR) => {
          this.state.tr.Id = result.Id;
          this.setDirty(false);
          this.setState(this.state);
        })
        .catch((response) => {
          this.state.errorMessages.push(new md.Message(response.data.responseBody['odata.error'].message.value));
          this.setState(this.state);
        });
    } else {
      this.setState(this.state); // show errors
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
    for (let instanceName in this.ckeditor.instances) {
      let instance = this.ckeditor.instances[instanceName];
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
    _.remove(messageList, {
      Id: messageId
    });
    this.setState(this.state);
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
    } else return (
      <PrimaryButton buttonType={ButtonType.primary} onClick={this.save} icon="ms-Icon--Save">
        <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
        Save
      </PrimaryButton>

    );
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
    }
    else return false;
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
      return (tr.TestsId.indexOf(TestId) != -1);
    }
    else {
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
    var tempPigments = _.filter(this.props.pigments, (pigment: Pigment) => {
      return !this.trContainsPigment(this.state.tr, pigment.id);
    });
    var pigments = _.map(tempPigments, (pigment) => {
      return {
        title: pigment.title,
        manufacturer: (pigment.manufacturer) ? pigment.manufacturer : "(none)",
        id: pigment.id,
        isActive: pigment.isActive
      };
    }).filter((p) => { return p.isActive === "Yes"; });
    return _.orderBy(pigments, ["manufacturer"], ["asc"]);
  }


  /**
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
    //var pigmentTypes=_.uniqWith(pigs,(p1:Pigment,p2:Pigment)=>{return p1.type === p2.type});
    var pigmentManufactureres = _.countBy(pigs, (p1: Pigment) => { return p1.manufacturer; });
    var groups: Array<IGroup> = [];
    debugger;
    for (const pm in pigmentManufactureres) {
      groups.push({
        isCollapsed: (pm !== this.state.expandedPigmentManufacturer),
        name: pm,
        key: pm,
        startIndex: _.findIndex(pigs, (pig) => { return pig.manufacturer === pm; }),
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
    var tempPigments = _.filter(this.props.pigments, (pigment: Pigment) => {
      return this.trContainsPigment(this.state.tr, pigment.id);
    });
    var selectedPigments = _.map(tempPigments, (pigment) => {
      return {
        title: pigment.title,
        manufacturer: (pigment.manufacturer) ? pigment.manufacturer : "(none)",
        id: pigment.id,
        isActive: pigment.isActive
      };
    });
    return _.orderBy(selectedPigments, ["title"], ["asc"]);
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
    var temppropertyTest: Array<PropertyTest> = _.filter(this.props.propertyTests, (pt: PropertyTest) => {
      return (pt.applicationTypeid === this.state.tr.ApplicationTypeId
        && pt.endUseIds.indexOf(this.state.tr.EndUseId) != -1
      );
    });
    // get all the the tests in those propertyTests and output them as DisplayPropertyTest
    var tempDisplayTests: Array<DisplayPropertyTest> = [];
    for (const pt of temppropertyTest) {
      console.log(`checking PropertyTest ${pt.id} with property ${pt.property} which has ${pt.testIds.length} tests`);
      for (const testid of pt.testIds) {
        console.log(`Looking for  testid ${testid}`);
        const test: Test = _.find(this.props.tests, (t) => { return t.id === testid; });
        if (test) {
          console.log(`adding test ${test.title}`);
          tempDisplayTests.push({
            property: pt.property,
            testid: testid,
            test: test.title
          });
        }
        else {
          console.log(` test ${testid} NOT FOUND`);
        }

      }
    }
    // now remove those that are already on the tr
    const displayTests = _.filter(tempDisplayTests, (dt) => { return !this.trContainsTest(this.state.tr, dt.testid); });
    return _.orderBy(displayTests, ["type"], ["asc"]);


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
    var displayPropertyTests: Array<DisplayPropertyTest> = this.getAvailableTests(); // all the avalable tests with their Property
    var properties = _.countBy(displayPropertyTests, (dpt: DisplayPropertyTest) => { return dpt.property; });// an object with an element for each propert, the value of the elemnt is the count of tsts with that property
    var groups: Array<IGroup> = [];
    for (const property in properties) {
      groups.push({
        isCollapsed: (property !== this.state.expandedProperty),
        name: property,
        key: property,
        startIndex: _.findIndex(displayPropertyTests, (dpt) => { return dpt.property === property; }),
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

    var property = _.find(this.props.propertyTests, (pt) => {
      return (
        pt.applicationTypeid === applicationTypeid
        &&
        pt.testIds.indexOf(testId) !== -1
        &&
        pt.endUseIds.indexOf(endUseId) !== -1

      );
    });
    return (property) ? property.property : '';
  }
  /**
   * return Tests on the tr being edited
   * 
   * @returns {Array<Test>} 
   * 
   * @memberof TrForm
   */
  public getSelectedTests(): Array<Test> {
    var tempTests = _.filter(this.props.tests, (test: Test) => {
      return this.trContainsTest(this.state.tr, test.id);
    });
    var selectedTests = _.map(tempTests, (test) => {
      return {
        title: test.title,
        id: test.id,
        propertyName: this.getPropertyName(test.id, this.state.tr.ApplicationTypeId, this.state.tr.EndUseId)
      };
    });
    return _.orderBy(selectedTests, ["propertyName", "title"], ["asc", "asc"]);
  }







  /**
   * Gets the technical Specialists, including an indicator if they are selected or not
   * 
   * @returns 
   * 
   * @memberof TrForm
   */
  public getTechSpecs() {
    var techSpecs = _.map(this.props.techSpecs, (techSpec) => {
      return {
        title: techSpec.title,
        selected: ((this.state.tr.TRAssignedToId) ? this.state.tr.TRAssignedToId.indexOf(techSpec.id) != -1 : null),
        id: techSpec.id
      };
    });
    return _.orderBy(techSpecs, ["selected", "title"], ["desc", "asc"]);
  }

  /**
   *  Gets the Staff CC (Same group as technical Specialists), including an indicator if they are selected or not
   * 
   * @returns 
   * 
   * @memberof TrForm
   */
  // public getStaffCC() {
  //   var staffCC = _.map(this.props.techSpecs, (techSpec) => {
  //     return {
  //       title: techSpec.title,
  //       selected: ((this.state.tr.StaffCCId) ? this.state.tr.StaffCCId.indexOf(techSpec.id) != -1 : null),
  //       id: techSpec.id
  //     };
  //   });
  //   return _.orderBy(staffCC, ["selected", "title"], ["desc", "asc"]);
  // }
  public staffCCChanged(items?: Array<IPersonaProps>): void {
    this.state.tr.StaffCC = items;
    this.props.ensureUsersInPersonas(this.state.tr.StaffCC);
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
    this.setDirty(true);
    if (isSelected) {
      if (this.state.tr.TRAssignedToId) {
        this.state.tr.TRAssignedToId.push(id);//addit
      }
      else {
        this.state.tr.TRAssignedToId = [id];
      }
    }
    else {
      this.state.tr.TRAssignedToId = _.filter(this.state.tr.TRAssignedToId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
  }



  /**
   * Adds or removes a preson from the StaffCC on the TR being edited
   * 
   * @param {boolean} isSelected 
   * @param {number} id 
   * 
   * @memberof TrForm
   */
  // public toggleStaffCC(isSelected: boolean, id: number) {
  //   this.setDirty(true);
  //   if (isSelected) {
  //     if (this.state.tr.StaffCCId) {
  //       this.state.tr.StaffCCId.push(id);//addit
  //     }
  //     else {
  //       this.state.tr.StaffCCId = [id];
  //     }
  //   }
  //   else {
  //     this.state.tr.StaffCCId = _.filter(this.state.tr.StaffCCId, (x) => { return x != id; });//remove it
  //   }
  //   this.setState(this.state);
  // }

  /******** TEST Toggles , this is two lists, toggling adds from one , removes from the other*/
  public addTest(id: number) {
    this.setDirty(true);
    if (this.state.tr.TestsId) {
      this.state.tr.TestsId.push(id);//addit
    }
    else {
      this.state.tr.TestsId = [id];
    }
    this.setState(this.state);
  }
  public removeTest(id: number) {
    this.setDirty(true);
    if (this.state.tr.TestsId) {
      this.state.tr.TestsId = _.filter(this.state.tr.TestsId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
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
    this.setDirty(true);
    if (this.state.tr.PigmentsId) {
      this.state.tr.PigmentsId.push(id);//addit
    }
    else {
      this.state.tr.PigmentsId = [id];
    }
    this.setState(this.state);
  }
  /**
  * Removes  the selected Pigment to the TR being edited
  * In the UI Pigmemts is two lists (selected and unselected pigments, toggling adds from one , removes from the other
  * 
  * @param {number} id The ID if the pigment to add or remove
  * 
  * @memberof TrForm
  */
  public removePigment(id: number) {
    this.setDirty(true);
    if (this.state.tr.PigmentsId) {
      this.state.tr.PigmentsId = _.filter(this.state.tr.PigmentsId, (x) => { return x != id; });//remove it
    }
    this.setState(this.state);
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
    this.state.showTRSearch = false;
    this.setState(this.state);
  }
  //make the child tr the currently selected tr


  /**
   * Makes the selected Child TR the current TR. Called when the user clicks the edit icon in the child TR List.
   * 
   * @param {number} trId The ID of the Child TR to edit.
   * @returns {*} 
   * 
   * @memberof TrForm
   */
  public selectChildTR(trId: number): any {
    const childTr = _.find(this.state.childTRs, (tr) => { return tr.Id === trId; });

    if (childTr) {
      console.log("switching to tr " + trId);
      delete this.state.tr;
      this.state.tr = childTr;
      this.originalAssignees = _.clone(this.state.tr.TRAssignedToId);
      this.originalStatus = this.state.tr.TRStatus;
      this.updateCKEditorText(this.state.tr);
      this.state.childTRs = [];

      this.setState(this.state);
      // now get its children, need to move children to state
      this.props.fetchChildTr(this.state.tr.Id).then((trs) => {
        this.state.childTRs = trs;
        this.setState(this.state);
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

    //mode: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(trdocument.id, 1).then(url => {

      if (!url || url === "") {
        window.open(trdocument.serverRalativeUrl, '_blank');
      }
      else {
        window.open(url, '_blank');
      }
      //    this.state.wopiFrameUrl=url;
      //  this.setState(this.state);
      // window.location.href=url;

    });

  }

  public parentTRSelected(id: number, title: string) {
    this.state.tr.ParentTR = title;
    this.state.tr.ParentTRId = id;
    this.setDirty(true);
    this.cancelTrSearch();
  }
  public uploadFile(e: any) {

    let target: any = e.target as any;
    let file = e.target["files"][0];
    this.props.uploadFile(file, this.state.tr.Id).then((response) => {
      this.props.getDocuments(this.state.tr.Id).then((dox) => {
        this.state.documents = dox;
        this.setState(this.state);
      });

    }).catch((error) => {

    });
  }
  public editParentTR() {

    if (this.state.tr.ParentTRId) {
      const parentId = this.state.tr.ParentTRId;
      this.props.fetchTR(parentId).then((parentTR) => {

        this.state.tr = parentTR;
        this.originalAssignees = _.clone(this.state.tr.TRAssignedToId);
        this.originalStatus = this.state.tr.TRStatus;
        this.state.childTRs = [];
        this.setState(this.state);
        this.updateCKEditorText(this.state.tr);
        this.props.fetchChildTr(parentId).then((subTRs) => {
          this.state.childTRs = subTRs;
          this.setState(this.state);
        });
      });
    }


  }
  public documentRowMouseEnter(trdocument: TRDocument, e: any) {

    //mode passed to fetchDocumentWopiFrameURL: 0: view, 1: edit, 2: mobileView, 3: interactivePreview
    this.props.fetchDocumentWopiFrameURL(trdocument.id, 0).then(url => {
      if (!url || url === "") {
        url = trdocument.serverRalativeUrl;
      }
      this.state.documentCalloutIframeUrl = url;
      this.state.documentCalloutTarget = e.target;
      this.state.documentCalloutVisible = true;
      this.setState(this.state);

    });
  }
  public documentRowMouseOut(item: TRDocument, e: any) {

    this.state.documentCalloutTarget = null;
    this.state.documentCalloutVisible = false;
    this.setState(this.state);
    console.log("mouse exit for " + item.title);
  }

  public createSummaryMarkup(tr: TR) {
    return { __html: tr.Summary };
  }

  public setDirty(isDirty: boolean) {
    debugger;
    if (!this.state.isDirty && isDirty) { //wasnt dirty now it is 
      window.onbeforeunload = function (e) {
        var dialogText = "You have unsaved changes, are you sure you want to leave?";
        e.returnValue = dialogText;
        return dialogText;
      };
    }
    if (this.state.isDirty && !isDirty) { //was dirty now it is not
      window.onbeforeunload = null;
    };
    this.state.isDirty = isDirty;
  }
  public render(): React.ReactElement<ITrFormProps> {

    let worktypeDropDoownoptions = _.map(this.props.workTypes, (wt) => {
      return {
        key: wt.id, text: wt.workType
      };
    });

    let applicationtypeDropDoownoptions =
      _.filter(this.props.applicationTypes, (at) => {
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
      _.filter(this.props.endUses, (eu) => {
        // show if its valid for the selected ApplicationType, OR if its already on the tr
        return (eu.applicationTypeId === this.state.tr.ApplicationTypeId
          || eu.id === this.state.tr.EndUseId);
      })
        .map((eu) => {
          return {
            key: eu.id, text: eu.endUse
          };
        });

    let customerSelectOptions = _.map(this.props.customers, (c) => {
      return {
        value: c.id, label: c.title
      };
    });
    return (
      <div >

        <MessageDisplay messages={this.state.errorMessages}
          hideMessage={this.removeMessage.bind(this)} />
        <div style={{ float: "left" }}> <Label>MODE : {modes[this.props.mode]}</Label></div>
        <div style={{ float: "right" }}>  <Label>Status : {(this.state.isDirty) ? "Unsaved" : "Saved"}</Label></div>
        <div style={{ clear: "both" }}></div>
        <table>

          <tr>
            <td>
              <Label >Request #</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Title} onChanged={e => {
                this.setDirty(true);
                this.state.tr.Title = e; this.setState(this.state);
              }} />
            </td>
            <td>
              <Label >Work Type</Label>
            </td>
            <td>
              <Dropdown label=''
                selectedKey={this.state.tr.WorkTypeId}
                options={worktypeDropDoownoptions}
                onChanged={e => {
                  this.setDirty(true);
                  this.state.tr.WorkTypeId = e.key as number;
                  this.setState(this.state);
                }} />
            </td>

            <td>
              <Label >Site</Label>
            </td>
            <td>
              <TextField value={this.state.tr.Site} onChanged={e => {
                this.setDirty(true);
                this.state.tr.Site = e;
                this.setState(this.state);
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
                <i onClick={(e) => { this.state.showTRSearch = true; this.setState(this.state); }}
                  className="ms-Icon ms-Icon--Search" aria-hidden="true"></i>

              </div>

            </td>
            <td>
              <Label >Application Type</Label>
            </td>
            <td>
              <Dropdown label=''
                selectedKey={this.state.tr.ApplicationTypeId}
                options={applicationtypeDropDoownoptions}
                onChanged={e => {
                  this.setDirty(true);
                  this.state.tr.ApplicationTypeId = e.key as number;
                  this.setState(this.state);
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
                  { key: 'High', text: 'High' },
                  { key: 'Normal', text: 'Normal' },
                  { key: 'Routine', text: 'Routine' },

                ]}
                onChanged={e => {
                  this.setDirty(true);
                  this.state.tr.TRPriority = e.text;
                  this.setState(this.state);
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
                this.setDirty(true);
                this.state.tr.CER = e;
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
              <Select
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
              />
              {/* <TagPicker ref="customerPicker"
                onResolveSuggestions={this.onCustomerResolveSuggestions.bind(this)}
                onChange={this.onCustomerChanged.bind(this)}
                pickerSuggestionsProps={
                  {
                    noResultsFoundText: 'No Matches found'
                  }
                }
              /> */}
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
                  this.setDirty(true);
                  this.state.tr.RequestDate = moment(e).toISOString();
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >End Use</Label>
            </td>
            <td>
              <Dropdown label=''
                selectedKey={this.state.tr.EndUseId}
                options={enduseDropDoownoptions}
                onChanged={e => {
                  this.setDirty(true);
                  this.state.tr.EndUseId = e.key as number;
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >Status</Label>
            </td>
            <td>
              <Dropdown
                label=""
                options={[
                  { key: 'Pending', text: 'Pending' },
                  { key: 'In Progress', text: 'In Progress' },
                  { key: 'Completed', text: 'Completed' },
                  { key: 'Canceled', text: 'Canceled' },
                ]}
                onChanged={e => {
                  this.setDirty(true);
                  this.state.tr.TRStatus = e.text;
                  this.setState(this.state);
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
                  this.setDirty(true);
                  this.state.tr.RequiredDate = moment(e).toISOString();
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >Actual Start Date</Label>
            </td>
            <td>
              <DatePicker
                value={(this.state.tr.ActualStartDate) ? moment(this.state.tr.ActualStartDate).toDate() : null}
                onSelectDate={e => {
                  this.setDirty(true);
                  this.state.tr.ActualStartDate = moment(e).toISOString();
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label >Actual Completion Date</Label>
            </td>
            <td>
              <DatePicker
                value={(this.state.tr.ActualCompletionDate) ? moment(this.state.tr.ActualCompletionDate).toDate() : null}
                onSelectDate={e => {
                  this.setDirty(true);
                  this.state.tr.ActualCompletionDate = moment(e).toISOString();
                  this.setState(this.state);
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
                  this.setDirty(true);
                  this.state.tr.EstManHours = parseInt(e);
                  this.setState(this.state);
                }} />
            </td>
            <td>
              <Label ></Label>
            </td>


            <td>
              <Label ></Label>
            </td>
            <td>

            </td>

          </tr>


        </table>
        <tabs.Tabs onSelect={this.tabChanged.bind(this)}>
          <tabs.TabList>
            <tabs.Tab>
              Title
             </tabs.Tab>
            <tabs.Tab>
              Description
             </tabs.Tab>
            <tabs.Tab>
              Summary
             </tabs.Tab>
            <tabs.Tab>
              Test Params
             </tabs.Tab>
            <tabs.Tab>
              Assigned To({(this.state.tr.TRAssignedToId) ? this.state.tr.TRAssignedToId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Staff cc({(this.state.tr.StaffCC) ? this.state.tr.StaffCC.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Pigments({(this.state.tr.PigmentsId) ? this.state.tr.PigmentsId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Tests({(this.state.tr.TestsId) ? this.state.tr.TestsId.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Formulae
             </tabs.Tab>
            <tabs.Tab>
              Child TRs({(this.state.childTRs) ? this.state.childTRs.length : 0})
             </tabs.Tab>
            <tabs.Tab>
              Documents
             </tabs.Tab>
          </tabs.TabList>
          <tabs.TabPanel >

            <textarea name="tronoxtrtextarea-title" id="tronoxtrtextarea-title" style={{ display: "none" }}>
              {this.state.tr.RequestTitle}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-description" id="tronoxtrtextarea-description" style={{ display: "none" }}>
              {this.state.tr.Description}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <div dangerouslySetInnerHTML={this.createSummaryMarkup(this.state.tr)} />
            <textarea name="tronoxtrtextarea-summary" id="tronoxtrtextarea-summary" style={{ display: "none" }}>
              {this.state.tr.SummaryNew}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-testparams" id="tronoxtrtextarea-testparams" style={{ display: "none" }}>
              {this.state.tr.TestingParameters}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              selectionMode={SelectionMode.none}
              items={this.getTechSpecs()}
              setKey="id"
              columns={[
                { key: "title", name: "Technical Specialist", fieldName: "title", minWidth: 20, maxWidth: 200 },
                {
                  key: "selected", name: "Assigned?", fieldName: "selected", minWidth: 200, onRender: (item) => <Checkbox
                    checked={item.selected}
                    onChange={(element, value) => { this.toggleTechSpec(value, item.id); }}
                  />
                }
              ]}
            />
          </tabs.TabPanel>
          <tabs.TabPanel>
            <NormalPeoplePicker
              defaultSelectedItems={this.state.tr.StaffCC}
              onChange={this.staffCCChanged.bind(this)}
              onResolveSuggestions={this.props.peopleSearch}
            >
            </NormalPeoplePicker>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <div style={{ float: "left" }}>
              <Label> Available Pigments</Label>
              <DetailsList
                onDidUpdate={(dl: DetailsList) => {
                  // save expanded group in state;
                  debugger;
                  var expandedGroup = _.find(dl.props.groups, (group) => {
                    return !(group.isCollapsed) && group.key !== this.state.expandedPigmentManufacturer;// its an expanded group that want expanded before
                  });
                  if (expandedGroup) {
                    this.state.expandedPigmentManufacturer = expandedGroup.key;
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
                    key: "select", name: "Select", fieldName: "selected", minWidth: 80, onRender: (item) => <Checkbox
                      checked={false}
                      onChange={(element, value) => { this.addPigment(item.id); }}
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
                    key: "select", name: "Select", fieldName: "selected", minWidth: 80, onRender: (item) => <Checkbox
                      checked={true}
                      onChange={(element, value) => { this.removePigment(item.id); }}
                    />
                  }
                ]}
              />
            </div>
            <div style={{ clear: "both" }}></div>
          </tabs.TabPanel>
          <tabs.TabPanel>
            <div style={{ float: "left" }}>
              <Label> Available Tests</Label>
              <DetailsList
                onDidUpdate={(dl: DetailsList) => {
                  // save expanded group in state;
                  debugger;
                  var expandedGroup = _.find(dl.props.groups, (group) => {
                    return !(group.isCollapsed) && group.key !== this.state.expandedProperty; // its an expanded group that want expanded before
                  });
                  if (expandedGroup) {
                    this.state.expandedProperty = expandedGroup.key;
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
                    key: "select", name: "Select", fieldName: "selected", minWidth: 70, onRender: (item) => <Checkbox
                      checked={false}

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
                    key: "selected", name: "Selected?", fieldName: "selected", minWidth: 70, onRender: (item) => <Checkbox
                      checked={true}
                      onChange={(element, value) => { this.removeTest(item.id); }}
                    />
                  }
                ]}
              />

            </div>
            <div style={{ clear: "both" }}></div>



          </tabs.TabPanel>
          <tabs.TabPanel>
            <textarea name="tronoxtrtextarea-formulae" id="tronoxtrtextarea-formulae" style={{ display: "none" }}>
              {this.state.tr.Formulae}
            </textarea>
          </tabs.TabPanel>
          <tabs.TabPanel>

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
          </tabs.TabPanel>
          <tabs.TabPanel>
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
                  {
                    key: "Edit", name: "", fieldName: "Title", minWidth: 20,
                    onRender: (item) => <div>
                      <i onClick={(e) => { this.editDocument(item); }}
                        className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>
                    </div>
                  },
                  { key: "title", name: "Request #", fieldName: "title", minWidth: 80, },

                ]}
              />
              <input type='file' id='uploadfile' onChange={e => { this.uploadFile(e); }} />
            </div>
            <div style={{ float: "right" }}>
              <DocumentIframe src={this.state.documentCalloutIframeUrl} />
            </div>
            <div style={{ clear: "both" }}></div>


          </tabs.TabPanel>
        </tabs.Tabs>

        <this.SaveButton />
        <span style={{ margin: 20 }}>
          <PrimaryButton href="#" onClick={this.cancel} icon="ms-Icon--Cancel">
            <i className="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
            Cancel
        </PrimaryButton>
        </span>
        <br />
        version 2
      < TRPicker
          isOpen={this.state.showTRSearch}
          callSearch={this.props.TRsearch}
          cancel={this.cancelTrSearch}
          select={this.parentTRSelected}
        />
      </div >
    );
  }
}
