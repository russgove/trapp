import * as React from "react";
import { ITrTimeCardProps } from "./ITrTimeCardProps";
import { ITrTimeCardState } from "./ITrTimeCardState";
//import { escape } from "@microsoft/sp-lodash-subset";
import {
} from "office-ui-fabric-react/lib/Pickers";
import { PrimaryButton, ButtonType } from "office-ui-fabric-react/lib/Button";
import { Link } from "office-ui-fabric-react/lib/Link";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import { Label } from "office-ui-fabric-react/lib/Label";
import { MessageBar, MessageBarType, } from "office-ui-fabric-react/lib/MessageBar";
import { Dropdown, IDropdownProps, } from "office-ui-fabric-react/lib/Dropdown";
import { SpinButton, ISpinButtonProps, ISpinButtonStyles, KeyboardSpinDirection } from "office-ui-fabric-react/lib/SpinButton";
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from "office-ui-fabric-react/lib/DetailsList";
import { DatePicker, } from "office-ui-fabric-react/lib/DatePicker";
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from "office-ui-fabric-react/lib/Persona";
import { IPersonaWithMenu } from "office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props";
import * as moment from "moment";
import { reduce, find } from "lodash";
import { TimeSpent } from "../dataModel";

export default class TrTimeCard extends React.Component<ITrTimeCardProps, ITrTimeCardState> {
  constructor(props: ITrTimeCardProps) {
    super(props);
    this.state = props.initialState;
    this.setState(this.state);
    this.updateNewHoursSpent = this.updateNewHoursSpent.bind(this);
    this.renderNewHoursSpent = this.renderNewHoursSpent.bind(this);
    this._getErrorMessage = this._getErrorMessage.bind(this);
    this.save = this.save.bind(this);
  }

  /**
   * Validator for the Hours worked field. Must be nimeric
   * 
   * @private
   * @param {string} value Value entered into the field.
   * @returns {string} An error message, or an empty string
   * 
   * @memberof TrTimeCard
   */
  private _getErrorMessage(value: string): string {
    var test = Number(value);
    return isNaN(test)
      ? ` ${value} is invalid.`
      : "";
  }


  /**
   * Called when the user enter info into the textbox, this method updates the
   * newhours spent on the TR stored in state. When we save back to sharepoint we 
   * add newHoursSpent to hoursSpent
   * 
   * @param {number} trId  The ID of the TR the hours were entered for
   * @param {*} newValue  Th evalue entered
   * 
   * @memberof TrTimeCard
   */
  public updateNewHoursSpent(trId: number, newValue: any) {
    let timeSpent = find(this.state.timeSpents, (ts) => { return ts.trId === trId; });
    //this.state.message = "";
    if (timeSpent) {
      timeSpent.newHoursSpent = parseFloat(newValue);
    } else {
      console.log(`Cannot find timespent record with a TR id of ${trId}`);
    }
    this.setState((current) => ({ ...current, message: "" }));
  }
  private _hasSuffix(string: string, suffix: string): Boolean {
    let subString = string.substr(string.length - suffix.length);
    return subString === suffix;
  }

  private _removeSuffix(string: string, suffix: string): string {
    if (!this._hasSuffix(string, suffix)) {
      return string;
    }

    return string.substr(0, string.length - suffix.length);
  }
  /**
   * Called by the Details list to reneder the HoursWorked textbox with appropriate handlers
   * 
   * @param {*} [item] The timeSpent record
   * @param {number} [index]  The index of the TimeSpent record in the array//not used
   * @param {IColumn} [column]  The column tyo display//not used
   * @returns 
   * 
   * @memberof TrTimeCard
   */
  public renderNewHoursSpent(item?: TimeSpent, index?: number, column?: IColumn) {
    let suffix = ' hours   ';
    return (<SpinButton
      label=""

      key={item.trId.toString() + item.weekEndingDate}
      value={item.newHoursSpent + suffix}
      min={-168}
      max={168}
      step={0.25}


      onValidate={(value: string) => {
        console.log("in on validate");
        value = this._removeSuffix(value, suffix);
        if (isNaN(+value)) {
          console.log("value set to 0");
          item["newHoursSpent"] = 0;
          return '0' + suffix;

        }
        item["newHoursSpent"] = parseFloat(value);
        this.setState({});
        console.log("value set to " + value);
        return String(value) + suffix;
      }}
      onIncrement={(value: string) => {
        console.log("in on increment");
        let newValue = parseFloat(this._removeSuffix(value, suffix)) + .25;
        
        
        item["newHoursSpent"] = newValue;
        console.log("value set to " + newValue);
        this.setState({});
        return String(newValue) + suffix;
      }}
      onDecrement={(value: string) => {
        console.log("in on decrement");
        let newValue = parseFloat(this._removeSuffix(value, suffix)) - .25;

        item["newHoursSpent"] = newValue;
        this.setState({});
        console.log("value set to " + newValue);
        return String(newValue) + suffix;
      }}
      onBlur={(e) => {
        console.log("in on blir");
        let value = e.currentTarget.value;
        let newValue = parseFloat(this._removeSuffix(value, suffix));
        this.setState({});
        item["newHoursSpent"] = newValue;
        console.log("value set to " + newValue);

      }
      }


    />);
    // return (<TextField
    //   value={item.newHoursSpent}
    //   onGetErrorMessage={this._getErrorMessage}
    //   validateOnFocusIn
    //   validateOnFocusOut
    //   onChanged={(newValue) => { this.updateNewHoursSpent(item.trId, newValue); }}
    // />);
  }

  /**
   * Saves all data entered back to sharepoint. Then Updates the state with the info retuended. Note that When 
   * new TimeSpent rows are created , the ID of the new row is returned, so we need to update our state to have the 
   * new row id.
   * @returns 
   * 
   * @memberof TrTimeCard
   */
  public save() {
    debugger;
    console.log("in Save");
    this.props.save(this.state.timeSpents)
      .then((timespents) => {
        console.log("did update");

        this.props.getTimeSpent(this.state.weekEndingDate).then((timeSpents) => {
          console.log("fetched new");
          this.setState((current) => ({ ...current, timeSpents: [] }));  // i shouldnt need to do this, but neccesary bacuse spinner does not see new value
          this.setState((current) => ({ ...current, timeSpents: timeSpents }));
        });

      })
      .catch((error) => {
        //this.state.message = error;
        this.setState((current) => ({ ...current, message: error }));

      });
    return false; // stop postback
  }

  /**
   * Renders the display.
   * @returns {React.ReactElement<ITrTimeCardProps>} 
   * 
   * @memberof TrTimeCard
   */
  public render(): React.ReactElement<ITrTimeCardProps> {

    return (
      <div>
        <Label>Time spent by Technical Specialist {this.props.userName} for the week ending {this.state.weekEndingDate.toDateString()}   </Label>
        <DatePicker
          value={moment(this.state.weekEndingDate).toDate()}
          onSelectDate={e => {
            const weekEndingDate = moment(e).utc().endOf("isoWeek").startOf("day").toDate();
            this.props.getTimeSpent(weekEndingDate).then((timeSpents) => {
              // this.state.weekEndingDate = weekEndingDate;
              // this.state.timeSpents = timeSpents;
              this.setState((current) => ({ ...current, weekEndingDate: weekEndingDate, timeSpents: timeSpents }));
            });

          }} />
        <Label>(if you do not see a TR you are working on displayed here, please ask you adminstrator to assign it to you, or to reopen it) </Label>
        <Label>{this.state.message}</Label>
        <DetailsList
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionMode={SelectionMode.none}
          items={this.state.timeSpents}
          setKey="Id"

          columns={[
            {
              key: "TR", name: "Request #", fieldName: "trTitle", minWidth: 20, maxWidth: 75,
              onRender: (item) => <a href={this.props.editFormUrlFormat.replace("{1}", item.trId).replace("{2}", window.location.href).replace("{3}", this.props.webUrl)}>{item.trTitle}</a>
            },
            {
              key: "Description", name: "Title", fieldName: "trRequestTitle", minWidth: 20, maxWidth: 150,
              onRender: (item) => <div style={{ "whiteSpace": "normal" }} dangerouslySetInnerHTML={{ __html: item.trRequestTitle }} />
            },
            { key: "Priority", name: "Priority", fieldName: "trPriority", minWidth: 20, maxWidth: 50 },

            { key: "Status", name: "Status", fieldName: "trStatus", minWidth: 20, maxWidth: 50 },
            {
              key: "Required", name: "Required Date", fieldName: "trRequiredDate", minWidth: 20, maxWidth: 70,
              onRender: (item) => <div>{moment(item.trRequiredDate).format("DD-MMM-YYYY")}</div>
            },
            { key: "WorkTyoe", name: "Work Type", fieldName: "trWorkType", minWidth: 20, maxWidth: 90 },
            { key: "AssignedTo", name: "Assigned To", fieldName: "trAssignedTo", minWidth: 20, maxWidth: 70 },

            {
              key: "hoursSpent", name: "Previous Hours", fieldName: "hoursSpent",
              minWidth: 100
            },
            {
              key: "newHoursSpent", name: "New Hours", fieldName: "newHoursSpent",
              minWidth: 130, onRender: this.renderNewHoursSpent
            }
          ]}
        />
        <table>
          <tr>

            <td>
              <Label>Previous Hours</Label>
            </td>
            <td>
              {reduce(this.state.timeSpents, (sum, ts) => {
                return sum + ts.hoursSpent;
              }, 0)}
            </td>
            <td>
              <Label>New Hours</Label>
            </td>
            <td>
              {reduce(this.state.timeSpents, (sum, ts) => {
                return sum + ts.newHoursSpent;
              }, 0)}
            </td>
          </tr>
        </table>
        <span style={{ margin: 20 }}>
          <PrimaryButton href="#" onClick={this.save} icon="ms-Icon--Save">
            <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
            Save
        </PrimaryButton>
        </span>
      </div>

    );
  }
}
