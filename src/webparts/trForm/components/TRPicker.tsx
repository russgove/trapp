import * as React from 'react';
import styles from './TrForm.module.scss';
import { TR, modes } from "../dataModel";
import {
    Modal, IModalProps
} from 'office-ui-fabric-react/lib/Modal';
import {
    SearchBox, ISearchBoxProps
} from 'office-ui-fabric-react/lib/SearchBox';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { Button, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn,SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import * as _ from "lodash";
export interface iTrPickerState {
    searchText: string;
    searchRusults: Array<TR>;
}
import * as moment from 'moment';
export interface iTrPickerProps {
    isOpen: boolean;
    callSearch: (searchText: string) => Promise<TR[]>; // method tyo call to searcgh gpr parenttr
    cancel: () => void; // method tyo call to searcgh gpr parenttr
    select: (Id: number, Title: string) => void; // method tyo call to searcgh gpr parenttr


}

export default class TRPicker extends React.Component<iTrPickerProps, iTrPickerState> {

    constructor(props: iTrPickerProps) {
        super(props);
        debugger;
        this.state = {
            searchText: null,
            searchRusults: []
        };
        this.renderSelect = this.renderSelect.bind(this);
 

    }

    public cancel() {

        return false; // stop postback

    }

    public renderSelect(item?: any, index?: number, column?: IColumn): JSX.Element {
        
        return (<i
            onClick={(e) => { debugger; this.props.select(item.Id,item.Title); return false; }}
            className="ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i>);
        /*return (<a href="#" onClick={(e) => { debugger; this.selectChildTR(item.Id) }}>
          {item[column.fieldName]}
        </a>);*/
    }
    public renderDate(item?: any, index?: number, column?: IColumn): any {

        return moment(item[column.fieldName]).format("MMM Do YYYY");
    }
    public doSearch(newValue: any): void {
 debugger;
        this.props.callSearch(newValue).then((results) => {
            this.state.searchRusults = results;
            this.setState(this.state);
        });
    }
    public render(): React.ReactElement<iTrPickerProps> {

        return (
            <Modal isOpen={this.props.isOpen}>
                
                    <SearchBox onSearch={this.doSearch.bind(this)} />
                    <DetailsList
                        selectionMode={SelectionMode.none}
                        layoutMode={DetailsListLayoutMode.fixedColumns}
                        items={this.state.searchRusults}
                        setKey="id"
                        columns={[
                            { key: "Select", onRender: this.renderSelect, name: "", fieldName: "Title", minWidth: 20, },
                            { key: "Title", name: "Request #", fieldName: "Title", minWidth: 80, },
                            { key: "CER", name: "CER", fieldName: "CER", minWidth: 90 },
                            { key: "Customer", name: "Customer", fieldName: "Customer", minWidth: 80 },
                            { key: "Site", name: "Site", fieldName: "Site", minWidth: 80 },

                        ]}
                    />
               <Link href="#" onClick={this.props.cancel} style={{ border: 5, backgroundColor: 'lightBlue', fontSize: 'large' }}>
            <i className="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
            Cancel
        </Link>
            </Modal>
        );
    }
}
