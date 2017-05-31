import * as React from 'react';
import { ITrTimeCardProps } from './ITrTimeCardProps';
import { ITrTimeCardState } from './ITrTimeCardState';
import { escape } from '@microsoft/sp-lodash-subset';
///// Add tab for testinParameters text block


import {
  NormalPeoplePicker, CompactPeoplePicker, IBasePickerSuggestionsProps,
} from 'office-ui-fabric-react/lib/Pickers';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType, } from 'office-ui-fabric-react/lib/MessageBar';
import { Dropdown, IDropdownProps, } from 'office-ui-fabric-react/lib/Dropdown';
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn, SelectionMode, IGroup } from 'office-ui-fabric-react/lib/DetailsList';
import { DatePicker, } from 'office-ui-fabric-react/lib/DatePicker';
import { IPersonaProps, PersonaPresence, PersonaInitialsColor, Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { IPersonaWithMenu } from 'office-ui-fabric-react/lib/components/pickers/PeoplePicker/PeoplePickerItems/PeoplePickerItem.Props';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as moment from 'moment';
import * as _ from "lodash";
import { TimeSpent } from "../dataModel";
export default class TrTimeCard extends React.Component<ITrTimeCardProps, ITrTimeCardState> {
  constructor(props: ITrTimeCardProps) {

    super(props);
    this.state = props.initialState;
    this.setState(this.state);
    this.getDisplayTRs = this.getDisplayTRs.bind(this);
    this.updateHoursSpent = this.updateHoursSpent.bind(this);
    this.renderHoursSpent = this.renderHoursSpent.bind(this);
    this.save = this.save.bind(this);
  }
  public getDisplayTRs(): Array<TimeSpent> {
    return this.state.timeSpents;
  }

  public updateHoursSpent(trId: number, newValue: number) {

    let timeSpent = _.find(this.state.timeSpents, (ts) => { return ts.trId === trId; });
    if (timeSpent) {
      timeSpent.hoursSpent = newValue;
    } else {
      console.log(`Cannot find timespent record with a TR id of ${trId}`);
    }

  }

  public renderHoursSpent(item?: any, index?: number, column?: IColumn) {

    return (<TextField
      value={item.hoursSpent}
      onChanged={(newValue) => { debugger; this.updateHoursSpent(item.trId, newValue) }}
    />);
  }
  public save() {
    this.props.save(this.state.timeSpents)
      .then((response) => {
     
        ;
      })
      .catch((error) => {
     
      })
    return false; // stop postback
  }
  public render(): React.ReactElement<ITrTimeCardProps> {
    debugger;



    return (
      <div>
        <Label>Time spent by Technical Specialist {this.props.userName} for the week ending {this.state.weekEndingDate.toDateString()}   </Label>
        <Label>(if you do not see a TR you are working on displayed here, please ask you adminstrator to assign it to you, or to reopen it) </Label>

        <DetailsList
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selectionMode={SelectionMode.none}
          items={this.state.timeSpents}
          setKey="Id"
          columns={[
            { key: "trId", name: "trId", fieldName: "trId", minWidth: 20, maxWidth: 20 },
            { key: "tsId", name: "tsId", fieldName: "tsId", minWidth: 20, maxWidth: 20 },
            { key: "TR", name: "TR", fieldName: "trTitle", minWidth: 20, maxWidth: 100 },
            { key: "Status", name: "Status", fieldName: "trStatus", minWidth: 20, maxWidth: 100 },
            { key: "Required", name: "Required", fieldName: "trRequired", minWidth: 20, maxWidth: 100 },
            { key: "hoursSpent", name: "hoursSpent", fieldName: "hoursSpent", minWidth: 100, onRender: this.renderHoursSpent }
          ]}
        />
        <span style={{ margin: 20 }}>
          <Button href="#" onClick={this.save} icon="ms-Icon--Save">
            <i className="ms-Icon ms-Icon--Save" aria-hidden="true"></i>
            Save
        </Button>
        </span>
      </div>

    );
  }
}
