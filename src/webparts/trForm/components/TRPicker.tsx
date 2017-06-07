/** FAbric */
import {
    Modal, IModalProps
} from 'office-ui-fabric-react/lib/Modal';
import {
    SearchBox, ISearchBoxProps
} from 'office-ui-fabric-react/lib/SearchBox';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { PrimaryButton, ButtonType } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { DetailsList, IDetailsListProps, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { List } from 'office-ui-fabric-react/lib/List';
import {  IStyle,  ITheme,  getTheme,  mergeStyles} from '@uifabric/styling';

/** Framework */
import * as React from 'react';

/**Custom Stuff */
import styles from './TrForm.module.scss';
import { TR } from "../dataModel";

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
            onClick={(e) => { debugger; this.props.select(item.Id, item.Title); return false; }}
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
            <Modal isOpen={this.props.isOpen} >

                <SearchBox onSearch={this.doSearch.bind(this)} />
                <List style={{ "width": "700px" }}


                    items={this.state.searchRusults}
                    onRenderCell={(item, index) => (
                        <div style={{ "width": "700px" }} onClick={(e) => { debugger; this.props.select(item.Id, item.Title); return false; }} >

                            <div >
                                <span style={{ "padding-right": "80px", "display": "block", "fontSize": "21px", "color": "#212121" }}>{item.Title}</span>
                                <span style={{ "font-size": "14px", "color": "#333333" }}>
                                    <Label style={{ "display": "inline" }} >Customer:</Label> <Label style={{ "display": "inline" }} >{item.CustomerId}</Label>
                                </span>
                                <div style={{ "width": "550px", "white-space": "normal" }} className='ms-ListItem-tertiaryText'>{item.Description}</div>
                            </div>

                        </div>
                    )}

                />
                <PrimaryButton  theme={getTheme()} href="#" onClick={this.props.cancel} icon="ms-Icon--Cancel">
                    <i className="ms-Icon ms-Icon--Cancel" aria-hidden="true"></i>
                    Cancel
        </PrimaryButton>

            </Modal>
        );
    }
}
