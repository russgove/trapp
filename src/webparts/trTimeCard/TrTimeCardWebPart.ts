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
    });
  }
  public AddTimeSpent(batch, timeSpent: TimeSpent): Promise<number> {
    return pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.inBatch(batch).add({
      TRId: timeSpent.trId,
      TechmicalSpecialistId: timeSpent.technicalSpecialist,
      WeekEndingDate: timeSpent.weekEndingDate,
      HoursSpent: timeSpent.hoursSpent
    }).then((response) => {
      debugger;
      return response.data.Id;
      // CAPTURE ID TO RETURN
    })

  }
  public UpdateTimeSpent(batch, timeSpent: TimeSpent) {
    pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.getById(timeSpent.tsId).inBatch(batch).update({
      TRId: timeSpent.trId,
      TechmicalSpecialistId: timeSpent.technicalSpecialist,
      WeekEndingDate: timeSpent.weekEndingDate,
      HoursSpent: timeSpent.hoursSpent
    }).then((response) => {

    })
      .catch((error) => {
        console.log("ERROR, An error occured Updateing TimeSPent");
        console.log(JSON.stringify(error));
      });
  }
  public save(timeSpents: Array<TimeSpent>): Promise<Array<TimeSpent>> {

    let batch = pnp.sp.createBatch();
    for (const timeSpent of timeSpents) {
      if (timeSpent.tsId === null) {
        this.AddTimeSpent(batch, timeSpent).then((id)=>{
          timeSpent.tsId=id;
        });
      }
      else {
        this.UpdateTimeSpent(batch, timeSpent);
      }

    }
    return batch.execute().then((x)=>{
      return timeSpents;
    });
  }
  public getAssignedTrs(batch?: any): Promise<Array<TechnicalRequest>> {

    // get the Active TRS Assigned to the user. These need to be shown in the timesheet
    let filterString = `(TRAssignedTo/EMail eq '${this.context.pageContext.user.email}') and (TRStatus ne 'Completed')`;
    let command = pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items.expand("TRAssignedTo")
      .select("Title,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail")
      .filter(filterString)
      .orderBy('RequiredDate');
    if (batch) {
      command.inBatch(batch);
    }

    return command.get().then((items) => {

      return _.map(items, (item) => {
        return {
          trId: item["Id"],
          title: item["Title"],
          status: item["TRStatus"],
          requiredDate: item["RequiredDate"],
          priority: item["TRPriority"],
        }
      });
    });
  }
  public getTR(trId: number, batch?: any): Promise<TechnicalRequest> {

    // get the Active TRS Assigned to the user. These need to be shown in the timesheet
    let command = pnp.sp.web.lists.getByTitle(this.properties.technicalRequestListName).items
      .getById(trId)
      .expand("TRAssignedTo")
      .select("Title,TRStatus,RequiredDate,Id,TRAssignedTo/Id,TRAssignedTo/EMail");

    if (batch) {
      command.inBatch(batch);
    }
    return command.get().then((item) => {

      return {
        trId: item["Id"],
        title: item["Title"],
        status: item["TRStatus"],
        requiredDate: item["RequiredDate"],
        priority: item["TRPriority"],
      }
    });

  }
  public getExistingTimeSpent(weekEndingDate: Date, batch?: any): Promise<Array<TimeSpent>> {
    let tsFilter = `(TechmicalSpecialist/EMail eq '${this.context.pageContext.user.email}') and (WeekEndingDate eq datetime'${weekEndingDate.toISOString()}')`;
    let command = pnp.sp.web.lists.getByTitle(this.properties.timeSpentListName).items.expand("TechmicalSpecialist,TR")
      .select("Id,WeekEndingDate,TRId,HoursSpent,TechmicalSpecialist/Id,TechmicalSpecialist/EMail,TR/Title,TR/Id,TR/RequiredDate")
      .filter(tsFilter);
    if (batch) {
      command.inBatch(batch);
    }
    return command.get()
      .then((items) => {
        return _.map(items, (item) => {

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
      });

  }
  public render(): void {

    var defaultWeekEndDate: Date = new Date(moment().utc().endOf('isoWeek').startOf('day'));
    let props: ITrTimeCardProps = {
      userName: this.context.pageContext.user.displayName,
      userId: null,
     
      save: this.save.bind(this),
      initialState: {
        weekEndingDate: defaultWeekEndDate,
        timeSpents: [],
         message:"",
      }
    }
    // mainBatch is used to fetch TimeSpents and users TRs
    let mainBatch = pnp.sp.createBatch();


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
    this.getAssignedTrs(mainBatch).then((items) => {
      activeTRs = items;
    })
    // get the Existing TimeSpents for the user in the selected weeek
    this.getExistingTimeSpent(defaultWeekEndDate, mainBatch).then(items => {
      props.initialState.timeSpents = items;
    });
    mainBatch.execute()
      .then((data) => {

        // trBatc is used to fetch TRS associated with the timeSPents, needs to execute after we get all  the timespents
        let trBatch = pnp.sp.createBatch();
        // 
        for (let timeSpent of props.initialState.timeSpents) {
          this.getTR(timeSpent.trId, trBatch)
            .then((tr) => {

              timeSpent.trPriority = tr.priority;
              timeSpent.trStatus = tr.status;
              timeSpent.trTitle = tr.title;
              timeSpent.trRequiredDate = tr.requiredDate;
            })
            .catch((error) => {
              console.log("ERROR, An error occured fetching the TRs for the TS");

              console.log(error.message);
            });

        }
        // add any trs user has not reported time for yet
        trBatch.execute().then((x) => {

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
