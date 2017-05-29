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
    var defaultWeekEndDate: Date = new Date(moment().endOf('isoWeek'));
    let props: ITrTimeCardProps = {
      activeTRs: [], initialState: {
        weekEndingDate: defaultWeekEndDate
      }
    }
    let batch = pnp.sp.createBatch();
    // pnp.sp.web.currentUser.get().then((user) => {
    pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.expand("TRAssignedTo")
      .select("Title,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail")
      .filter("TRAssignedTo/EMail eq '" + this.context.pageContext.user.email + "'")
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
        this.reactElement = React.createElement(TrTimeCard, props);
        var formComponent: TrTimeCard = ReactDom.render(this.reactElement, this.domElement) as TrTimeCard;//render the component

      }).catch((error) => {
        console.log("ERROR, An error occured fetching TRS");
        debugger;
        console.log(error.message);
      });
    // })
    // .catch((error) => {
    //   console.log("ERROR, An error occured fetching currentuser");
    //   console.log(error.message);

    // });


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
