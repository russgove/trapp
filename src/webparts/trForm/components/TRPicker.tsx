import * as React from 'react';
import styles from './TrForm.module.scss';
import { TR, modes } from "../dataModel";
import {
  Modal, IModalProps
} from 'office-ui-fabric-react/lib/Modal';
import {
  SearchBox,ISearchBoxProps
} from 'office-ui-fabric-react/lib/SearchBox';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import * as _ from "lodash";
export interface iTrPickerState {
    searchText:string;
    searchRusults:Array<TR>;
}
import * as moment from 'moment';
export interface iTrPickerProps {
    
  
}

export default class TRPicker extends React.Component<iTrPickerProps, iTrPickerState> {
  
constructor(props: iTrPickerProps) {
    super(props);
    this.state = {
        searchText:null,
        searchRusults:[]
    };

  }


  public save() {

   
    }

  public cancel() {

    return false; // stop postback

  }
 public selectTR(trId: number): any {
 debugger;
    return false;
  }
   public renderSelect(item?: any, index?: number, column?: IColumn): JSX.Element {
    debugger;
    return (<i
      onClick={(e) => { debugger; this.selectTR(item.Id) }}
      className="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>);
    /*return (<a href="#" onClick={(e) => { debugger; this.selectChildTR(item.Id) }}>
      {item[column.fieldName]}
    </a>);*/
  }
    public renderDate(item?: any, index?: number, column?: IColumn): any {

    return moment(item[column.fieldName]).format("MMM Do YYYY");
  }
  public doSearch(newValue: any) : void{
      
  }
  public render(): React.ReactElement<iTrPickerProps> {
  
    return (
      <div>
<SearchBox  onSearch={this.doSearch.bind(this)}/>
            <DetailsList
              layoutMode={DetailsListLayoutMode.fixedColumns}
              items={this.state.searchRusults}
              setKey="id"
              columns={[
                { key: "Select", onRender: this.renderSelect, name: "", fieldName: "Title", minWidth: 20, },

                { key: "Title", name: "Request #", fieldName: "Title", minWidth: 80, },
                { key: "Status", name: "Status", fieldName: "Status", minWidth: 90 },
                { key: "InitiationDate", onRender: this.renderDate, name: "Initiation Date", fieldName: "InitiationDate", minWidth: 80 },
                { key: "TRDueDate", onRender: this.renderDate, name: "Due Date", fieldName: "TRDueDate", minWidth: 80 },
                { key: "ActualStartDate", onRender: this.renderDate, name: "Actual Start Date", fieldName: "ActualStartDate", minWidth: 90 },
                { key: "ActualCompetionDate", onRender: this.renderDate, name: "Actual Competion<br />Date", fieldName: "ActualCompetionDate", minWidth: 80 },

              ]}
            />
        </div>
    );
  }
}
