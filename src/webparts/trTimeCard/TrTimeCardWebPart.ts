import * as React from 'react';
import * as ReactDom from 'react-dom';
import { TechnicalRequest } from "./dataModel";
import { Version } from '@microsoft/sp-core-library';
import pnp from "sp-pnp-js";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import * as moment from "moment";
import * as strings from 'trTimeCardStrings';
import TrTimeCard from './components/TrTimeCard';
import { ITrTimeCardProps } from './components/ITrTimeCardProps';
import { ITrTimeCardState } from './components/ITrTimeCardState';
import { ITrTimeCardWebPartProps } from './ITrTimeCardWebPartProps';
import * as _ from 'lodash';
export default class TrTimeCardWebPart extends BaseClientSideWebPart<ITrTimeCardWebPartProps> {
  debugger;
  private reactElement: React.ReactElement<ITrTimeCardProps>;
  public render(): void {
    debugger;
    var defaultWeekEndDate: Date = new Date(moment().utc().endOf('isoWeek').startOf('day'));
    let props: ITrTimeCardProps = {
      activeTRs: [],
      userName: this.context.pageContext.user.displayName,
      initialState: {
        weekEndingDate: defaultWeekEndDate,
        timeSpents: [],

      }
    }
    let batch = pnp.sp.createBatch();
    let filterString = `(TRAssignedTo/EMail eq '${this.context.pageContext.user.email}') and (TRStatus ne 'Completed')`;
    pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.expand("TRAssignedTo")
      .select("Title,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail")
      .filter(filterString)
      .orderBy('RequiredDate')
      .inBatch(batch)
      .get()
      .then((items) => {
        debugger;
        props.activeTRs = _.map(items, (item) => {
          return {
            id: item["Id"],
            title: item["Title"],
            status: item["TRStatus"],
            requiredDate: item["RequiredDate"],
          }
        });
      }).catch((error) => {
        console.log("ERROR, An error occured fetching TRS");
        debugger;
        console.log(error.message);
      });
    let tsFilter = `(TechmicalSpecialist/EMail eq '${this.context.pageContext.user.email}') and (WeekEndingDate eq datetime'${defaultWeekEndDate.toISOString()}')`;

    pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.expand("TechmicalSpecialist,TR")
      .select("Id,WeekEndingDate,TRId,HoursSpent,TechmicalSpecialist/Id,TechmicalSpecialist/EMail,TR/Title,TR/Id")
      .filter(tsFilter)
      // .orderBy('RequiredDate')
      .inBatch(batch)
      .get()
      .then((items) => {
        debugger;
        props.initialState.timeSpents = _.map(items, (item) => {
          let tr: TechnicalRequest = {
            id: item["TRId"],
            title: item["TR"]["Title"],
            status: item["TR"]["Status"],
            requiredDate: item["TR"]["RequiredDate"],
          }
          return {
            Id: item["Id"],
            TechnicalSpecialist: item["TechnicalSpecialist"],
            TR: tr,
            WeekEndingDate: item["WeekEndingDate"],
            HoursSpent: item["HoursSpent"]
          }
        });
      }).catch((error) => {
        console.log("ERROR, An error occured fetching timeSpents");
        debugger;
        console.log(error.message);
      });
    batch.execute()
      .then((data) => {
        debugger;
        this.reactElement = React.createElement(TrTimeCard, props);
        var formComponent: TrTimeCard = ReactDom.render(this.reactElement, this.domElement) as TrTimeCard;//render the component
      })
      .catch((error) => {
        console.log("ERROR, An error occured executing the batch");
        debugger;
        console.log(error.message);

      })


  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
