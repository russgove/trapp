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
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from "office-ui-fabric-react/lib/DetailsList";
import { DatePicker, } from "office-ui-fabric-react/lib/DatePicker";
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from "office-ui-fabric-react/lib/Persona";
import { IPersonaWithMenu } from "office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props";
import * as moment from "moment";
import * as _ from "lodash";
import { TimeSpent } from "../dataModel";
export default class TrTimeCard extends React.Component<ITrTimeCardProps, ITrTimeCardState> {
  constructor(props: ITrTimeCardProps) {
    super(props);
    this.state = props.initialState;
    this.setState(this.state);
    this.updateHoursSpent = this.updateHoursSpent.bind(this);
    this.renderHoursSpent = this.renderHoursSpent.bind(this);
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
   * Called when the user enter info into the textbox, this method updates the hours spent on the TR stored in state
   * 
   * @param {number} trId  The ID of the TR the hours were entered for
   * @param {*} newValue  Th evalue entered
   * 
   * @memberof TrTimeCard
   */
  public updateHoursSpent(trId: number, newValue: any) {
    let timeSpent = _.find(this.state.timeSpents, (ts) => { return ts.trId === trId; });
    //this.state.message = "";
    if (timeSpent) {
      timeSpent.hoursSpent = newValue;
    } else {
      console.log(`Cannot find timespent record with a TR id of ${trId}`);
    }
    this.setState({ ...this.state, message: "" });
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
  public renderHoursSpent(item?: any, index?: number, column?: IColumn) {
    const hoursSpent: number = item.hoursSpent;
    return (<TextField
      value={item.hoursSpent}
      onGetErrorMessage={this._getErrorMessage}
      validateOnFocusIn
      validateOnFocusOut
      onChanged={(newValue) => { this.updateHoursSpent(item.trId, newValue); }}
    />);
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
    this.props.save(this.state.timeSpents)
      .then((timespents) => {
        //this.state.timeSpents = timespents;
        //this.state.message = "Saved";
        this.setState({ ...this.state, timeSpents: timespents, message : "Saved" });
      })
      .catch((error) => {
        //this.state.message = error;
        this.setState({ ...this.state, message: error });
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
              this.setState({ ...this.state, weekEndingDate: weekEndingDate, timeSpents: timeSpents });
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
              key: "TR", name: "TR", fieldName: "trTitle", minWidth: 20, maxWidth: 100,
              onRender: (item) => <a href={this.props.editFormUrlFormat.replace("{1}", item.trId).replace("{2}", window.location.href).replace("{3}",this.props.webUrl)}>{item.trTitle}</a>
            },
            {
              key: "Description", name: "Title", fieldName: "trRequestTitle", minWidth: 20, maxWidth: 150,
              onRender: (item) => <div style={{"whiteSpace":"normal"}} dangerouslySetInnerHTML={{__html: item.trRequestTitle}} /> 
            },
            { key: "Status", name: "Status", fieldName: "trStatus", minWidth: 20, maxWidth: 70 },
            {
              key: "Required", name: "Required", fieldName: "trRequiredDate", minWidth: 20, maxWidth: 70,
              onRender: (item) => <div>{moment(item.trRequiredDate).format("DD-MMM-YYYY")}</div>
            },
            { key: "hoursSpent", name: "hoursSpent", fieldName: "hoursSpent", minWidth: 100, onRender: this.renderHoursSpent }
          ]}
        />
        <span style={{ margin: 20 }}>
          <PrimaryButton  href="#" onClick={this.save} icon="ms-Icon--Save">
            <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
            Save
        </PrimaryButton>
        </span>
      </div>

    );
  }
}
