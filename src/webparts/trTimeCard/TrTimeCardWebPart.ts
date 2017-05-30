import * as React from 'react';
import * as ReactDom from 'react-dom';
import { TimeSpent, TechnicalRequest } from "./dataModel";
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

  private reactElement: React.ReactElement<ITrTimeCardProps>;
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context,
      });
      //  this.loadData();
    });
  }
  public AddTimeSpent(batch, timeSpent: TimeSpent) {
    pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.inBatch(batch).add({
      TRId: timeSpent.trId,
      TechmicalSpecialistId: timeSpent.technicalSpecialist,
      WeekEndingDate: timeSpent.weekEndingDate,
      HoursSpent: timeSpent.hoursSpent
    }).then((response) => {
      debugger;
    })
      .catch((error) => {
        console.log("ERROR, An error occured adding timespent");
        console.log(JSON.stringify(error));
      });
  }
  public UpdateTimeSpent(batch, timeSpent: TimeSpent) {
    pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.getById(timeSpent.tsId).inBatch(batch).update({
      TRId: timeSpent.trId,
      TechmicalSpecialistId: timeSpent.technicalSpecialist,
      WeekEndingDate: timeSpent.weekEndingDate,
      HoursSpent: timeSpent.hoursSpent
    }).then((response) => {
      debugger;
    })
      .catch((error) => {
        console.log("ERROR, An error occured Updateing TimeSPent");
        console.log(JSON.stringify(error));
      });
  }
  public save(timeSpents: Array<TimeSpent>): Promise<any> {
    debugger
    let batch = pnp.sp.createBatch();
    for (const timeSpent of timeSpents) {
      if (timeSpent.tsId === null) {
        this.AddTimeSpent(batch, timeSpent);
      }
      else {
        this.UpdateTimeSpent(batch, timeSpent);
      }

    }
    return batch.execute();
  }
  public render(): void {
    debugger;
    var defaultWeekEndDate: Date = new Date(moment().utc().endOf('isoWeek').startOf('day'));
    let props: ITrTimeCardProps = {
      userName: this.context.pageContext.user.displayName,
      userId: null,
      save: this.save.bind(this),
      initialState: {
        weekEndingDate: defaultWeekEndDate,
        timeSpents: [],
      }
    }
    // mainBatch is used to fetch TimeSpents and users TRs
    let mainBatch = pnp.sp.createBatch();
    // trBatc is used to fetch TRS associated with the timeSPents, needs to execute after we get all  the timespents
    let trBatch = pnp.sp.createBatch();

    pnp.sp.web.currentUser.inBatch(mainBatch).get()
      .then((user) => {
        props.userId = user.Id;
      })
      .catch((error) => {
        console.log("ERROR, An error occured fetching currentUser");
        console.log(JSON.stringify(error));
      });

    let activeTRs: Array<TechnicalRequest> = [];
    // get the Active TRS Assigned to the user. These need to be shown in the timesheet
    let filterString = `(TRAssignedTo/EMail eq '${this.context.pageContext.user.email}') and (TRStatus ne 'Completed')`;
    pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.expand("TRAssignedTo")
      .select("Title,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail")
      .filter(filterString)
      .orderBy('RequiredDate')
      .inBatch(mainBatch)
      .get()
      .then((items) => {
        debugger;
        activeTRs = _.map(items, (item) => {
          return {
            trId: item["Id"],
            title: item["Title"],
            status: item["TRStatus"],
            requiredDate: item["RequiredDate"],
            priority: item["TRPriority"],
          }
        });
      }).catch((error) => {
        console.log("ERROR, An error occured fetching TRS");

        console.log(JSON.stringify(error));
      });
    // get the Existing TimeSpensts for the user in the selected weeek
    let tsFilter = `(TechmicalSpecialist/EMail eq '${this.context.pageContext.user.email}') and (WeekEndingDate eq datetime'${defaultWeekEndDate.toISOString()}')`;
    pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.expand("TechmicalSpecialist,TR")
      .select("Id,WeekEndingDate,TRId,HoursSpent,TechmicalSpecialist/Id,TechmicalSpecialist/EMail,TR/Title,TR/Id,TR/RequiredDate")
      .filter(tsFilter)
      // .orderBy('RequiredDate')
      .inBatch(mainBatch)
      .get()
      .then((items) => {

        props.initialState.timeSpents = _.map(items, (item) => {
          // queue up a reqest in the secon batch to get the TR associaated with the TS
          // this batch gets executed after the fisr batch is finisher
          debugger;
          pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.getById(item["TRId"])
            .expand("TRAssignedTo")
            .select("Title,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail")
            .inBatch(trBatch)
            .get()
            .then((item) => {
              debugger;
              let timeSpent = _.find(props.initialState.timeSpents, (ts) => { return ts.trId = item.Id });
              if (timeSpent) {
                timeSpent.trPriority = item["TRPriority"];
                timeSpent.trStatus = item["TRStatus"];
                timeSpent.trRequiredDate = item["RequiredDate"];
                timeSpent.trTitle = item["Title"];
              }
              else {
                console.log("ERROR, Timespent not found");
              }
            })
            .catch((error) => {
              debugger;
              console.log("ERROR, An error occured fetching Technical Requests associated with TimeSPent");
              console.log(JSON.stringify(error));
            })
          return {
            tsId: item["Id"],
            technicalSpecialist: item["TechnicalSpecialist"],
            trId: item["TRId"],
            weekEndingDate: item["WeekEndingDate"],
            hoursSpent: item["HoursSpent"],
            trTitle: null,
            trStatus: null,
            trPriority: null,
            trRequiredDate: null,
          }
        });
      }).catch((error) => {
        console.log("ERROR, An error occured fetching timeSpents");

        console.log(JSON.stringify(error));
      });
    mainBatch.execute()
      .then((data) => {
        for (const tr of activeTRs) {
          // add a row for any active projects not on list
          const itemIndex = _.findIndex(props.initialState.timeSpents, (item) => { return item.trId === tr.trId });
          if (itemIndex === -1) {
            props.initialState.timeSpents.push({
              trId: tr.trId,
              technicalSpecialist: props.userId,
              weekEndingDate: defaultWeekEndDate,
              hoursSpent: 0,
              tsId: null,
              trTitle: tr.title,
              trStatus: tr.status,
              trRequiredDate: tr.requiredDate,
              trPriority: tr.priority
            });
          }
          else {
            props.initialState.timeSpents[itemIndex].trPriority = tr.priority;
            props.initialState.timeSpents[itemIndex].trRequiredDate = tr.requiredDate;
            props.initialState.timeSpents[itemIndex].trStatus = tr.status;

          }
        }
        // execute the trBatch to get the Technical Requests for the TimeSheeets (cannot get Lookup  values with Expand!)
        trBatch.execute()
          .then((data) => {
            this.reactElement = React.createElement(TrTimeCard, props);
            var formComponent: TrTimeCard = ReactDom.render(this.reactElement, this.domElement) as TrTimeCard;//render the component
          })
      })

      .catch((error) => {
        console.log("ERROR, An error occured executing the batch");

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
