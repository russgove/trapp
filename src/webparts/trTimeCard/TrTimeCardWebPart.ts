import * as React from 'react';
import * as ReactDom from 'react-dom';
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
    var defaultWeekEndDate: Date = new Date(moment().endOf('isoWeek').startOf('day'));
    let props: ITrTimeCardProps = {
      activeTRs: [], initialState: {
        weekEndingDate: defaultWeekEndDate,
        timeSpents: []
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

    pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.expand("TechmicalSpecialist")
      .select("Id,WeekEndingDate,TRId,HoursSpent,TechmicalSpecialist/Id,TechmicalSpecialist/EMail")
      .filter(tsFilter)
      // .orderBy('RequiredDate')
      .inBatch(batch)
      .get()
      .then((items) => {
        props.initialState.timeSpents = _.map(items, (item) => {
          return {
            Id: item["Id"],
            TechnicalSpecialist: item["TechnicalSpecialist"],
            TR: item["TR"],
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
